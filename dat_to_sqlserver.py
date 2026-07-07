"""
dat_to_sqlserver.py
Load a Concordance/Relativity .dat load file directly into SQL Server.

Field format:  þvalueþ\x14þvalueþ...
    þ  (thorn, \xFE) = text qualifier wrapping each field
    \x14 (ASCII 20)  = column separator

Features:
    * Loads into an EXISTING table (does NOT drop it). It reads the table's
      real columns from SQL Server and maps the .dat header onto them by name
      (case-insensitive). Unmatched .dat columns are reported and skipped;
      table columns with no source are left NULL. If the table does not exist,
      it is created right-sized as a fallback.
    * RESUMABLE / crash-safe: each committed batch is recorded in a checkpoint
      file (<name>_checkpoint.txt). On any DB error or Ctrl+C the batch is
      retried a few times; if it still fails the checkpoint is saved and the
      script stops cleanly. Re-run the SAME command and it skips everything
      already loaded and continues from the last committed row -- nothing
      missed, nothing duplicated. The checkpoint is deleted on a clean finish.
    * PASS 1 analyses the file for the report (and to size a table it must
      create). PASS 2 loads rows in batches, committing each batch.
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
DEFAULT_DATABASE = "STS_NL"        # used only if you don't include a DB in the table name
USE_WINDOWS_AUTH = False           # False -> use the SQL login below (same as SSMS 'sa')
SQL_USER = "sa"
# Never hardcode a real password. Preferred: set an environment variable
#   PowerShell:  $env:SQL_PASSWORD = 'yourpassword'
# and leave the line below as-is. It falls back to the env var when present.
SQL_PASSWORD = os.environ.get("SQL_PASSWORD", "PUT_YOUR_SA_PASSWORD_HERE")


def build_conn_str(database: str) -> str:
    base = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={SERVER};DATABASE={database};"
    if USE_WINDOWS_AUTH:
        return base + "Trusted_Connection=yes;"
    return base + f"UID={SQL_USER};PWD={SQL_PASSWORD};"


DELIMITER = "þ\x14þ"
BATCH_SIZE = 1000
MAX_RETRIES = 3          # per-batch retry attempts before pausing (resumable)
RETRY_WAIT_SEC = 5       # wait between retries on a transient DB error
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


def write_report(report_path, input_file, full_table, cols, col_types, stats,
                 total_rows, rows_loaded, padded, trimmed, blank_skipped, elapsed):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    with open(report_path, "w", encoding="utf-8") as r:
        r.write("=" * 70 + "\n")
        r.write("DAT -> SQL SERVER IMPORT REPORT\n")
        r.write("=" * 70 + "\n")
        r.write(f"Generated       : {ts}\n")
        r.write(f"Source file     : {input_file}\n")
        r.write(f"Target table    : {full_table}  (columns auto-sized, no truncation)\n")
        r.write(f"Source size     : {os.path.getsize(input_file):,} bytes\n")
        r.write(f"Processing time : {elapsed:.1f} s\n")
        r.write(f"Total data rows : {total_rows:,}\n")
        r.write(f"Rows loaded     : {rows_loaded:,}\n")
        r.write(f"Total columns   : {len(cols)}\n")
        r.write(f"Blank rows skip : {blank_skipped:,}\n\n")

        r.write("-" * 70 + "\n")
        r.write("COLUMNS  (SQL column type | inferred data type | fill counts | max length)\n")
        r.write("-" * 70 + "\n")
        for st, t in zip(stats, col_types):
            r.write(
                f"{st.name[:34]:<34} | {t:<15} | {st.inferred_type():<22} | "
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


def sql_type(max_len: int) -> str:
    """Pick the smallest NVARCHAR that safely holds this column's data.
    Sizes to the observed max (+ small buffer). Uses NVARCHAR(MAX) only when
    the data genuinely exceeds 4000 characters, so nothing is ever truncated."""
    if max_len <= 0:
        return "NVARCHAR(1)"          # column is entirely empty
    if max_len > 4000:
        return "NVARCHAR(MAX)"        # true long text (e.g. extracted body)
    buffered = ((max_len // 50) + 1) * 50   # round up to next 50 for headroom
    return f"NVARCHAR({min(buffered, 4000)})"


def read_header(fin):
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
    return header, cols


def read_checkpoint(path: Path) -> int:
    """How many data rows were already committed on a previous run (0 = fresh)."""
    try:
        return int(path.read_text(encoding="utf-8").strip())
    except (FileNotFoundError, ValueError):
        return 0


def write_checkpoint(path: Path, rows_committed: int):
    """Record progress atomically so a crash mid-write can't corrupt it."""
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(str(rows_committed), encoding="utf-8")
    os.replace(tmp, path)   # atomic on Windows and POSIX


def get_existing_columns(cur, database, schema, table):
    """Return the target table's real column names in order, or [] if it doesn't exist."""
    cur.execute(
        f"SELECT COLUMN_NAME FROM [{database}].INFORMATION_SCHEMA.COLUMNS "
        f"WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ? ORDER BY ORDINAL_POSITION",
        schema, table,
    )
    return [row[0] for row in cur.fetchall()]


def commit_batch(cur, conn, insert_sql, batch):
    """Insert + commit one batch, retrying transient errors. Raises if it can't."""
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            cur.executemany(insert_sql, batch)
            conn.commit()
            return
        except pyodbc.Error as exc:
            try:
                conn.rollback()
            except pyodbc.Error:
                pass
            if attempt == MAX_RETRIES:
                raise
            sys.stdout.write(
                f"\n[retry {attempt}/{MAX_RETRIES}] batch failed: {exc}. "
                f"waiting {RETRY_WAIT_SEC}s...\n"
            )
            sys.stdout.flush()
            time.sleep(RETRY_WAIT_SEC)


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        print("ERROR: usage -> python dat_to_sqlserver.py <file.dat> <[Database.][schema.]Table>")
        sys.exit(1)

    dat_path = Path(sys.argv[1])
    database, schema, table = split_table(sys.argv[2])
    full_table = f"[{database}].[{schema}].[{table}]"
    report_path = dat_path.with_name(dat_path.stem + "_import_report.txt")
    checkpoint_path = dat_path.with_name(dat_path.stem + "_checkpoint.txt")
    total_bytes = os.path.getsize(dat_path)
    start_time = time.time()

    # ---------------- PASS 1: analyse (measure max length per column) --------
    print("PASS 1/2  Analysing file to size columns (no data written yet)...")
    with open(dat_path, "r", encoding="utf-8-sig", errors="ignore") as fin:
        header, cols = read_header(fin)
        ncols = len(cols)
        stats = [ColumnStats(c) for c in cols]
        print(f"{ncols} columns detected. Target: {full_table}")

        processed = len(header.encode("utf-8", "ignore"))
        total_rows = blank_skipped = 0
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
            total_rows += 1
            if total_rows % 1000 == 0:
                render_progress(processed, total_bytes, start_time, total_rows)
    render_progress(total_bytes, total_bytes, start_time, total_rows)
    print(f"\nAnalysis done: {total_rows:,} rows. Creating right-sized table...")

    # ---------------- Connect + resolve target columns -----------------------
    col_types = [sql_type(st.max_len) for st in stats]
    conn = pyodbc.connect(build_conn_str(database), autocommit=False)
    cur = conn.cursor()
    cur.fast_executemany = True

    existing = get_existing_columns(cur, database, schema, table)
    if existing:
        # Table already exists -> DO NOT drop. Map the .dat header onto the real
        # columns by name (case-insensitive). Anything unmatched is reported.
        by_lower = {c.lower(): c for c in existing}
        load_idx, load_cols = [], []       # source col index -> real table col
        unmatched_src = []
        for i, c in enumerate(cols):
            real = by_lower.get(c.lower())
            if real is not None:
                load_idx.append(i)
                load_cols.append(real)
            else:
                unmatched_src.append(c)
        if not load_cols:
            print(f"\n[ERROR] None of the .dat columns match table {full_table}.")
            print(f"        .dat columns : {cols[:20]}")
            print(f"        table columns: {existing[:20]}")
            sys.exit(1)
        print(f"Existing table found: mapping {len(load_cols)}/{len(cols)} columns.")
        if unmatched_src:
            print(f"[WARN] {len(unmatched_src)} .dat column(s) not in table (skipped): "
                  f"{unmatched_src[:20]}")
    else:
        # Table does not exist -> create it right-sized (no drop needed).
        print("Target table not found -> creating it.")
        col_defs = ",\n  ".join(f"[{c}] {t} NULL" for c, t in zip(cols, col_types))
        cur.execute(f"CREATE TABLE {full_table} (\n  {col_defs}\n);")
        conn.commit()
        load_idx = list(range(ncols))
        load_cols = cols

    placeholders = ",".join("?" for _ in load_cols)
    insert_sql = (f"INSERT INTO {full_table} "
                  f"([{'],['.join(load_cols)}]) VALUES ({placeholders})")

    # How many rows are already loaded from a previous (interrupted) run?
    resume_from = read_checkpoint(checkpoint_path)
    if resume_from:
        print(f"[RESUME] {resume_from:,} rows already committed -> skipping them, "
              f"continuing from row {resume_from + 1:,}.")

    # ---------------- PASS 2: load (commit + checkpoint each batch) ----------
    print("PASS 2/2  Loading rows into SQL Server...")
    load_start = time.time()
    seen = 0            # data rows scanned so far (matches the checkpoint counter)
    rows_loaded = resume_from
    try:
        with open(dat_path, "r", encoding="utf-8-sig", errors="ignore") as fin:
            fin.readline()  # skip header
            processed = len(header.encode("utf-8", "ignore"))
            batch = []
            for line in fin:
                processed += len(line.encode("utf-8", "ignore"))
                if not line.strip():
                    continue
                seen += 1
                if seen <= resume_from:
                    continue                # already loaded on a previous run
                values = [v.strip("þ \r\n") for v in line.split(DELIMITER)]
                if len(values) < ncols:
                    values += [""] * (ncols - len(values))
                values = values[:ncols]
                row = [values[i] for i in load_idx]
                batch.append([v if v != "" else None for v in row])
                if len(batch) >= BATCH_SIZE:
                    commit_batch(cur, conn, insert_sql, batch)   # retries transient errors
                    rows_loaded = seen
                    write_checkpoint(checkpoint_path, rows_loaded)   # resumable point
                    batch.clear()
                    render_progress(processed, total_bytes, load_start, rows_loaded)
            if batch:
                commit_batch(cur, conn, insert_sql, batch)
                rows_loaded = seen
                write_checkpoint(checkpoint_path, rows_loaded)
    except (pyodbc.Error, KeyboardInterrupt) as exc:
        # Everything up to `rows_loaded` is committed and recorded in the
        # checkpoint file. Re-run the SAME command to continue from here.
        cur.close()
        conn.close()
        print(f"\n[STOPPED] {type(exc).__name__}: {exc}")
        print(f"[SAFE] {rows_loaded:,} rows are committed and saved to "
              f"{checkpoint_path.name}.")
        print("        Re-run the exact same command to resume without losing "
              "or duplicating any rows.")
        sys.exit(2)

    render_progress(total_bytes, total_bytes, load_start, rows_loaded)
    cur.close()
    conn.close()

    # Full run finished cleanly -> the checkpoint is no longer needed.
    try:
        checkpoint_path.unlink()
    except FileNotFoundError:
        pass

    elapsed = time.time() - start_time
    print(f"\n[OK] Loaded {rows_loaded:,} rows into {full_table}.")
    write_report(report_path, dat_path, full_table, cols, col_types, stats,
                 total_rows, rows_loaded, padded, trimmed, blank_skipped, elapsed)


if __name__ == "__main__":
    main()
