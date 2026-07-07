"""
dat_to_sqlserver.py
Load a Concordance/Relativity .dat load file directly into SQL Server.

Field format:  þvalueþ\x14þvalueþ...
    þ  (thorn, \xFE) = text qualifier wrapping each field
    \x14 (ASCII 20)  = column separator

Features:
    * Uses the header row for column names
    * Creates the table with ALL columns as NVARCHAR(MAX) (no truncation)
    * Loads EVERY row, in batches
    * Live progress bar with % complete + estimated time remaining (ETA)
    * Auto-generates a report .txt: row/column counts, inferred data types,
      max lengths, empty counts, rows loaded, and any errors / missed rows.

Requires:  pip install pyodbc

------------------------------------------------------------------------------
RUN COMMANDS
------------------------------------------------------------------------------
    pip install pyodbc

    # Table name can include the database. Accepted forms:
    #   Table                    -> goes to DEFAULT_DATABASE.dbo.Table
    #   schema.Table             -> goes to DEFAULT_DATABASE.schema.Table
    #   Database.schema.Table    -> goes to the database you name
    python dat_to_sqlserver.py "Objects_1000125_export Part 2.dat" MyDatabase.dbo.Objects_1000125
------------------------------------------------------------------------------
Set SERVER (and SQL login if not using Windows auth) below before running.
The target DATABASE must already exist; the script creates the TABLE.
Keep real credentials out of source control.
"""
import os
import re
import sys
import time
from pathlib import Path
import pyodbc

# ---- EDIT THESE CONNECTION SETTINGS ----
SERVER = r"localhost\SQLEXPRESS"   # e.g. "MYPC\SQLEXPRESS" or "10.0.0.5"
DEFAULT_DATABASE = "StagingDB"     # used only if you don't include a DB in the table name
USE_WINDOWS_AUTH = True            # False -> use SQL login below
SQL_USER = "youruser"
SQL_PASSWORD = "yourpass"          # keep real credentials out of source control


def build_conn_str(database: str) -> str:
    base = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={SERVER};DATABASE={database};"
    if USE_WINDOWS_AUTH:
        return base + "Trusted_Connection=yes;"
    return base + f"UID={SQL_USER};PWD={SQL_PASSWORD};"


DELIMITER = "þ\x14þ"
BATCH_SIZE = 1000
# -----------------------------------------

INT_RE = re.compile(r"^-?\d+$")
DEC_RE = re.compile(r"^-?\d{1,3}(,\d{3})*(\.\d+)?$|^-?\d+(\.\d+)?$")
DATE_RE = re.compile(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}$|^\d{1,2}[-/]\d{1,2}[-/]\d{4}$")
DATETIME_RE = re.compile(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}[ T]\d{1,2}:\d{2}(:\d{2})?")


class ColumnStats:
    def __init__(self, name):
        self.name = name
        self.non_empty = 0
        self.empty = 0
        self.max_len = 0
        self.is_int = True
        self.is_dec = True
        self.is_date = True
        self.is_datetime = True

    def observe(self, value):
        if value == "" or value is None:
            self.empty += 1
            return
        self.non_empty += 1
        self.max_len = max(self.max_len, len(value))
        if self.is_int and not INT_RE.match(value):
            self.is_int = False
        if self.is_dec and not DEC_RE.match(value):
            self.is_dec = False
        if self.is_datetime and not DATETIME_RE.match(value):
            self.is_datetime = False
        if self.is_date and not DATE_RE.match(value):
            self.is_date = False

    def inferred_type(self):
        if self.non_empty == 0:
            return "EMPTY (no data)"
        if self.is_int:
            return "INTEGER"
        if self.is_datetime:
            return "DATETIME"
        if self.is_date:
            return "DATE"
        if self.is_dec:
            return "DECIMAL/NUMERIC"
        return f"TEXT (max len {self.max_len})"


def clean_col(name: str, idx: int) -> str:
    name = (name or "").strip("þ ").strip()
    name = re.sub(r"[^\w]", "_", name)
    if not name or name[0].isdigit():
        name = f"col_{idx}_{name}".rstrip("_")
    return name[:128]


def split_table(target: str):
    """Parse a target name into (database, schema, table).

    Accepts:
        Table                       -> (DEFAULT_DATABASE, dbo,    Table)
        schema.Table                -> (DEFAULT_DATABASE, schema, Table)
        Database.schema.Table       -> (Database,         schema, Table)
    """
    parts = [p.strip("[] ") for p in target.split(".")]
    if len(parts) == 3:
        database, schema, tbl = parts
    elif len(parts) == 2:
        database, schema, tbl = DEFAULT_DATABASE, parts[0], parts[1]
    else:
        database, schema, tbl = DEFAULT_DATABASE, "dbo", parts[0]
    return database, schema, tbl


def render_progress(processed, total, start_time, rows):
    pct = (processed / total * 100) if total else 0
    elapsed = time.time() - start_time
    rate = processed / elapsed if elapsed > 0 else 0
    eta = (total - processed) / rate if rate > 0 else 0
    bar_len = 30
    filled = int(bar_len * min(pct, 100) / 100)
    bar = "#" * filled + "-" * (bar_len - filled)
    sys.stdout.write(
        f"\r[{bar}] {pct:5.1f}% | {rows:,} rows | "
        f"elapsed {elapsed:4.0f}s | ETA {eta:4.0f}s"
    )
    sys.stdout.flush()


def write_report(report_path, input_file, full_table, cols, stats,
                 total_rows, rows_loaded, padded, trimmed, blank_skipped, elapsed):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    with open(report_path, "w", encoding="utf-8") as r:
        r.write("=" * 70 + "\n")
        r.write("DAT -> SQL SERVER IMPORT REPORT\n")
        r.write("=" * 70 + "\n")
        r.write(f"Generated       : {ts}\n")
        r.write(f"Source file     : {input_file}\n")
        r.write(f"Target table    : {full_table}  (all columns NVARCHAR(MAX))\n")
        r.write(f"Source size     : {os.path.getsize(input_file):,} bytes\n")
        r.write(f"Processing time : {elapsed:.1f} s\n")
        r.write(f"Total data rows : {total_rows:,}\n")
        r.write(f"Rows loaded     : {rows_loaded:,}\n")
        r.write(f"Total columns   : {len(cols)}\n")
        r.write(f"Blank rows skip : {blank_skipped:,}\n\n")

        r.write("-" * 70 + "\n")
        r.write("COLUMNS  (SQL type is NVARCHAR(MAX); inferred type is advisory)\n")
        r.write("-" * 70 + "\n")
        for st in stats:
            r.write(
                f"{st.name[:40]:<40} | {st.inferred_type():<22} | "
                f"filled={st.non_empty:,} | empty={st.empty:,} | maxlen={st.max_len}\n"
            )
        r.write("\n")

        r.write("-" * 70 + "\n")
        r.write("DATA ERRORS / MISSED DATA\n")
        r.write("-" * 70 + "\n")
        if total_rows != rows_loaded:
            r.write(
                f"[WARN] Row mismatch: parsed {total_rows:,} but loaded {rows_loaded:,}.\n"
            )
        if not padded and not trimmed:
            r.write("None. All rows had exactly the expected column count.\n")
        if padded:
            r.write(
                f"[WARN] {len(padded):,} row(s) had FEWER columns than the header "
                f"and were padded with NULLs (no data lost).\n"
                f"       First line numbers: {padded[:50]}\n"
            )
        if trimmed:
            r.write(
                f"[MISSED DATA] {len(trimmed):,} row(s) had MORE columns than the "
                f"header. Extra trailing values were dropped to keep columns aligned.\n"
                f"       Review these lines in the source file:\n"
                f"       First line numbers: {trimmed[:50]}\n"
            )
        r.write("\n" + "=" * 70 + "\n")
    print(f"\nReport written:   {report_path}")


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        print("ERROR: usage -> python dat_to_sqlserver.py <file.dat> <[Database.][schema.]Table>")
        sys.exit(1)

    dat_path = Path(sys.argv[1])
    database, schema, table = split_table(sys.argv[2])
    full_table = f"[{database}].[{schema}].[{table}]"
    report_path = dat_path.with_name(dat_path.stem + "_import_report.txt")
    total_bytes = os.path.getsize(dat_path)
    start_time = time.time()

    with open(dat_path, "r", encoding="utf-8-sig", errors="ignore") as fin:
        header = fin.readline().strip()
        raw_headers = header.split(DELIMITER)
        cols = [clean_col(c, i) for i, c in enumerate(raw_headers)]

        seen = {}
        for i, c in enumerate(cols):
            if c in seen:
                seen[c] += 1
                cols[i] = f"{c}_{seen[c]}"
            else:
                seen[c] = 0

        ncols = len(cols)
        stats = [ColumnStats(c) for c in cols]
        print(f"Connecting to {SERVER} / {database} ...")
        print(f"{ncols} columns detected. Target: {full_table}")

        conn = pyodbc.connect(build_conn_str(database), autocommit=False)
        cur = conn.cursor()
        cur.fast_executemany = True

        col_defs = ",\n  ".join(f"[{c}] NVARCHAR(MAX) NULL" for c in cols)
        cur.execute(f"""
            IF OBJECT_ID(N'[{database}].[{schema}].[{table}]', N'U') IS NOT NULL
                DROP TABLE {full_table};
            CREATE TABLE {full_table} (
              {col_defs}
            );
        """)

        placeholders = ",".join("?" for _ in cols)
        insert_sql = f"INSERT INTO {full_table} ([{'],['.join(cols)}]) VALUES ({placeholders})"

        processed = len(header.encode("utf-8", "ignore"))
        batch = []
        total_rows = rows_loaded = blank_skipped = 0
        padded, trimmed = [], []

        for line_no, line in enumerate(fin, start=2):
            processed += len(line.encode("utf-8", "ignore"))
            if not line.strip():
                blank_skipped += 1
                continue

            values = [v.strip("þ \r\n") for v in line.split(DELIMITER)]
            if len(values) < ncols:
                padded.append(line_no)
                values += [""] * (ncols - len(values))
            elif len(values) > ncols:
                trimmed.append(line_no)

            values = values[:ncols]
            for i in range(ncols):
                stats[i].observe(values[i])
            batch.append([v if v != "" else None for v in values])
            total_rows += 1

            if len(batch) >= BATCH_SIZE:
                cur.executemany(insert_sql, batch)
                rows_loaded += len(batch)
                batch.clear()

            if total_rows % 1000 == 0:
                render_progress(processed, total_bytes, start_time, total_rows)

        if batch:
            cur.executemany(insert_sql, batch)
            rows_loaded += len(batch)

        conn.commit()
        cur.close()
        conn.close()

    render_progress(total_bytes, total_bytes, start_time, total_rows)
    elapsed = time.time() - start_time
    print(f"\n[OK] Loaded {rows_loaded:,} rows into {full_table}.")
    write_report(report_path, dat_path, full_table, cols, stats,
                 total_rows, rows_loaded, padded, trimmed, blank_skipped, elapsed)


if __name__ == "__main__":
    main()
