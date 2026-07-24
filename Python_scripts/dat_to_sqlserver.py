"""
dat_to_sqlserver.py
Load a Concordance/Relativity .dat load file (or a standard .csv file)
directly into SQL Server. Format is auto-detected from the file extension.

.dat field format:  þvalueþ\x14þvalueþ...
    þ  (thorn, \xFE) = text qualifier wrapping each field
    \x14 (ASCII 20)  = column separator

.csv field format: standard RFC4180 -- comma-separated, double-quote
    qualifier, "" as an escaped quote. Parsed with Python's csv module,
    so quoted commas and embedded newlines inside a field are handled
    correctly (not just split line-by-line).

Features:
    * Loads into an EXISTING table (does NOT drop it). It reads the table's
      real columns from SQL Server and maps the .dat header onto them by name
      (case-insensitive). Unmatched .dat columns are reported and skipped;
      table columns with no source are left NULL. If the table does not exist,
      it is created right-sized as a fallback.
    * RESUMABLE / crash-safe: the resume point is the ACTUAL number of rows
      currently in the target table (SELECT COUNT_BIG(*)). On any DB error or
      Ctrl+C the batch is retried a few times; if it still fails the script
      stops cleanly. Re-run the SAME command and it skips exactly the rows
      already in the table and continues from there -- nothing missed, nothing
      duplicated. Because the resume point is read live from the table, it
      stays correct even if you TRUNCATE/clear the table between runs (it
      simply starts over from row 1). Rows are inserted in file order, so the
      table's row count always equals the number of source rows loaded.
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
    python dat_to_sqlserver.py "export.csv" MyDatabase.dbo.MyTable
------------------------------------------------------------------------------
Set SERVER (and SQL login if not using Windows auth) below before running.
The target DATABASE must already exist; the script creates the TABLE.
Keep real credentials out of source control.
"""
import csv
import os
import re
import sys
import time
from pathlib import Path
import pyodbc

# ---- EDIT THESE CONNECTION SETTINGS ----
SERVER = r"prdenvfdevm-3\MSSQLSERVER01"   # Developer edition instance (no 10 GB limit)
DEFAULT_DATABASE = "sts_db"        # used only if you don't include a DB in the table name
USE_WINDOWS_AUTH = False           # True -> Windows/domain login (like SSMS); False -> SQL login (sa) below
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


# Concordance/Relativity load-file delimiters:
#   \x14 (DC4, ASCII 20) = column separator  <-- what we split on
#   þ    (thorn, \xFE)   = text qualifier wrapping each field, stripped per-field
# Splitting on the DC4 separator alone (not "þ\x14þ") is robust: it still works
# even if the thorn byte is dropped during decoding (ANSI-exported files).
DELIMITER = "\x14"
BATCH_SIZE = 1000
MAX_RETRIES = 3          # per-batch retry attempts before pausing (resumable)
RETRY_WAIT_SEC = 5       # wait between retries on a transient DB error

# How to line up the .dat columns with an EXISTING table's columns:
#   "auto"     -> match by name; if the names barely match AND both sides have
#                 the same number of columns, fall back to matching by position
#                 (1st .dat col -> 1st table col, ...). Best default.
#   "name"     -> match strictly by (cleaned) column name, case-insensitive
#   "position" -> ignore names, match purely by column order
# Use "position" if your .dat header names differ from the table but the
# columns are in the same order (the usual Concordance/Relativity case).
COLUMN_MATCH = "auto"
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


def is_csv_file(path: Path) -> bool:
    return path.suffix.lower() == ".csv"


def dedupe_cols(cols):
    seen = {}
    out = []
    for c in cols:
        if c in seen:
            seen[c] += 1
            out.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            out.append(c)
    return out


def read_header(fin):
    header = fin.readline().strip()
    raw_headers = header.split(DELIMITER)
    cols = dedupe_cols([clean_col(c, i) for i, c in enumerate(raw_headers)])
    return header, cols


def read_header_csv(fin):
    """Read the header row of a .csv file via the csv module (handles a
    quoted header name that itself contains a comma).

    Reads the line with fin.readline() rather than handing fin straight to
    csv.reader(). Once a text file has been iterated via next()/for-loop,
    Python permanently disables fin.tell() on that handle (a long-standing
    CPython quirk) -- readline() doesn't trigger it, so byte-progress
    tracking (see CountingLineReader) keeps working for the rest of the file.
    """
    raw_headers = next(csv.reader([fin.readline()]))
    return dedupe_cols([clean_col(c, i) for i, c in enumerate(raw_headers)])


class CountingLineReader:
    """Feeds csv.reader lines via fin.readline() while tracking bytes
    consumed manually, since fin.tell() is unusable once anything iterates
    fin with next()/for-loop (see read_header_csv)."""

    def __init__(self, fin):
        self.fin = fin
        self.processed = 0

    def __iter__(self):
        return self

    def __next__(self):
        line = self.fin.readline()
        if not line:
            raise StopIteration
        self.processed += len(line.encode("utf-8", "ignore"))
        return line


def get_row_count(cur, full_table) -> int:
    """How many rows are already in the target table right now. This is the
    resume point: rows are inserted in file order, so the count equals the
    number of source rows already loaded. Reading it live from the table means
    the resume point can never get out of sync with the data -- if the table is
    truncated/cleared, the count is 0 and loading starts over from row 1."""
    cur.execute(f"SELECT COUNT_BIG(*) FROM {full_table}")
    return int(cur.fetchone()[0])


def get_existing_columns(cur, database, schema, table):
    """Return the target table's real column names in order, or [] if it doesn't exist."""
    cur.execute(
        f"SELECT COLUMN_NAME FROM [{database}].INFORMATION_SCHEMA.COLUMNS "
        f"WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ? ORDER BY ORDINAL_POSITION",
        schema, table,
    )
    return [row[0] for row in cur.fetchall()]


def build_mapping(dat_cols, table_cols, mode):
    """Decide which source column index feeds which real table column.

    Returns (load_idx, load_cols): parallel lists where row value at
    load_idx[k] is inserted into table column load_cols[k].

      mode="name"     -> match by cleaned name (case-insensitive)
      mode="position" -> match by order (Nth .dat col -> Nth table col)
      mode="auto"     -> try name; if it matches poorly AND both sides have the
                         same count, switch to position and say so.
    """
    def by_name():
        by_lower = {c.lower(): c for c in table_cols}
        idx, cols_out, unmatched = [], [], []
        for i, c in enumerate(dat_cols):
            real = by_lower.get(c.lower())
            if real is not None:
                idx.append(i)
                cols_out.append(real)
            else:
                unmatched.append(c)
        return idx, cols_out, unmatched

    def by_position():
        m = min(len(dat_cols), len(table_cols))
        return list(range(m)), list(table_cols[:m])

    if mode == "position":
        idx, cols_out = by_position()
        print(f"[MAP] Using POSITION matching ({len(idx)} columns by order).")
        if len(dat_cols) != len(table_cols):
            print(f"[WARN] Column counts differ (.dat={len(dat_cols)}, "
                  f"table={len(table_cols)}); extra columns on the longer side "
                  f"are ignored/left NULL. Check the pairs printed below.")
        return idx, cols_out

    idx, cols_out, unmatched = by_name()
    good = len(idx)
    if mode == "name":
        print(f"[MAP] Using NAME matching ({good}/{len(dat_cols)} matched).")
        if unmatched:
            print(f"[WARN] {len(unmatched)} .dat column(s) not in table (NULL): "
                  f"{unmatched[:20]}")
        return idx, cols_out

    # mode == "auto"
    enough = good >= 0.8 * min(len(dat_cols), len(table_cols))
    if enough:
        print(f"[MAP] Auto -> NAME matching ({good}/{len(dat_cols)} matched).")
        if unmatched:
            print(f"[WARN] {len(unmatched)} .dat column(s) not in table (NULL): "
                  f"{unmatched[:20]}")
        return idx, cols_out
    if len(dat_cols) == len(table_cols):
        pidx, pcols = by_position()
        print(f"[MAP] Auto -> POSITION matching (names matched only {good}/"
              f"{len(dat_cols)}, but column counts are equal so order is used).")
        return pidx, pcols
    # Poor name match and counts differ -> best effort by name, warn loudly.
    print(f"[MAP] Auto -> NAME matching, but only {good}/{len(dat_cols)} matched "
          f"and counts differ (.dat={len(dat_cols)}, table={len(table_cols)}).")
    print("[WARN] Many columns will be NULL. If the columns are actually in the "
          "same ORDER, set COLUMN_MATCH = \"position\" near the top and re-run.")
    return idx, cols_out


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
    total_bytes = os.path.getsize(dat_path)
    start_time = time.time()

    is_csv = is_csv_file(dat_path)
    open_kwargs = dict(encoding="utf-8-sig", errors="ignore")
    if is_csv:
        open_kwargs["newline"] = ""  # required by the csv module
    print(f"Detected format: {'CSV' if is_csv else 'DAT'} (from file extension)")

    # ---------------- PASS 1: analyse (measure max length per column) --------
    print("PASS 1/2  Analysing file to size columns (no data written yet)...")
    header = None
    with open(dat_path, "r", **open_kwargs) as fin:
        if is_csv:
            cols = read_header_csv(fin)
            processed = 0  # negligible vs. file size; refined once the loop starts
        else:
            header, cols = read_header(fin)
            processed = len(header.encode("utf-8", "ignore"))
        ncols = len(cols)
        stats = [ColumnStats(c) for c in cols]
        print(f"{ncols} columns detected. Target: {full_table}")

        total_rows = blank_skipped = 0
        padded, trimmed = [], []

        if is_csv:
            line_reader = CountingLineReader(fin)
            for line_no, raw_values in enumerate(csv.reader(line_reader), start=2):
                processed = line_reader.processed
                values = [v.strip() for v in raw_values]
                if not any(values):
                    blank_skipped += 1
                    continue
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
        else:
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
        # Table already exists -> DO NOT drop. Line up the .dat columns with the
        # real table columns using COLUMN_MATCH ("auto" / "name" / "position").
        load_idx, load_cols = build_mapping(cols, existing, COLUMN_MATCH)
        if not load_idx:
            print(f"\n[ERROR] Could not map any .dat column to table {full_table}.")
            print(f"        .dat columns : {cols[:20]}")
            print(f"        table columns: {existing[:20]}")
            sys.exit(1)
        print(f"Existing table found. Mapping {len(load_idx)} columns "
              f"(.dat={len(cols)}, table={len(existing)}).")
        print("First column pairs (.dat header -> table column):")
        for src_i, tbl_c in list(zip(load_idx, load_cols))[:8]:
            print(f"   {cols[src_i][:30]:<30} -> {tbl_c}")
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

    # Resume point = rows already in the table (0 if empty / just truncated).
    resume_from = get_row_count(cur, full_table)
    if resume_from:
        print(f"[RESUME] Table already has {resume_from:,} rows -> skipping them, "
              f"continuing from row {resume_from + 1:,}.")
    else:
        print("[FRESH] Table is empty -> loading from row 1.")

    # ---------------- PASS 2: load (commit + checkpoint each batch) ----------
    print("PASS 2/2  Loading rows into SQL Server...")
    load_start = time.time()
    seen = 0            # data rows scanned so far (matches the checkpoint counter)
    rows_loaded = resume_from
    try:
        with open(dat_path, "r", **open_kwargs) as fin:
            if is_csv:
                line_reader = CountingLineReader(fin)
                reader = csv.reader(line_reader)
                next(reader)  # skip header
                processed = line_reader.processed
            else:
                fin.readline()  # skip header
                processed = len(header.encode("utf-8", "ignore"))
            batch = []

            if is_csv:
                for raw_values in reader:
                    processed = line_reader.processed
                    values = [v.strip() for v in raw_values]
                    if not any(values):
                        continue
                    seen += 1
                    if seen <= resume_from:
                        continue            # already loaded on a previous run
                    if len(values) < ncols:
                        values += [""] * (ncols - len(values))
                    values = values[:ncols]
                    row = [values[i] for i in load_idx]
                    batch.append([v if v != "" else None for v in row])
                    if len(batch) >= BATCH_SIZE:
                        commit_batch(cur, conn, insert_sql, batch)   # retries transient errors
                        rows_loaded = seen                           # committed -> resumable
                        batch.clear()
                        render_progress(processed, total_bytes, load_start, rows_loaded)
            else:
                for line in fin:
                    processed += len(line.encode("utf-8", "ignore"))
                    if not line.strip():
                        continue
                    seen += 1
                    if seen <= resume_from:
                        continue            # already loaded on a previous run
                    values = [v.strip("þ \r\n") for v in line.split(DELIMITER)]
                    if len(values) < ncols:
                        values += [""] * (ncols - len(values))
                    values = values[:ncols]
                    row = [values[i] for i in load_idx]
                    batch.append([v if v != "" else None for v in row])
                    if len(batch) >= BATCH_SIZE:
                        commit_batch(cur, conn, insert_sql, batch)   # retries transient errors
                        rows_loaded = seen                           # committed -> resumable
                        batch.clear()
                        render_progress(processed, total_bytes, load_start, rows_loaded)
            if batch:
                commit_batch(cur, conn, insert_sql, batch)
                rows_loaded = seen
    except (pyodbc.Error, KeyboardInterrupt) as exc:
        # Everything up to `rows_loaded` is committed to the table. Re-run the
        # SAME command to continue -- it reads the table's row count and picks
        # up from exactly there.
        cur.close()
        conn.close()
        print(f"\n[STOPPED] {type(exc).__name__}: {exc}")
        print(f"[SAFE] {rows_loaded:,} rows are committed to the table.")
        print("        Re-run the exact same command to resume without losing "
              "or duplicating any rows.")
        sys.exit(2)

    render_progress(total_bytes, total_bytes, load_start, rows_loaded)
    cur.close()
    conn.close()

    elapsed = time.time() - start_time
    print(f"\n[OK] Loaded {rows_loaded:,} rows into {full_table}.")
    write_report(report_path, dat_path, full_table, cols, col_types, stats,
                 total_rows, rows_loaded, padded, trimmed, blank_skipped, elapsed)


if __name__ == "__main__":
    main()
