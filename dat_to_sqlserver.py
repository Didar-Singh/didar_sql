"""
dat_to_sqlserver.py
Load a delimited .dat file directly into SQL Server.

- Auto-detects the delimiter
- Uses the header row for column names
- Creates the table with ALL columns as NVARCHAR(MAX) (no data loss / truncation)

Requires:  pip install pyodbc

Usage:
    python dat_to_sqlserver.py "C:\\path\\to\\file.dat" dbo.MyTable
"""
import csv
import sys
import re
from pathlib import Path
import pyodbc

# ---- EDIT THESE CONNECTION SETTINGS ----
SERVER = r"localhost\SQLEXPRESS"   # e.g. "MYPC\SQLEXPRESS" or "10.0.0.5"
DATABASE = "StagingDB"
# Windows auth (recommended). For SQL auth, see the alt string below.
CONN_STR = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    f"SERVER={SERVER};DATABASE={DATABASE};Trusted_Connection=yes;"
)
# SQL auth alternative (keep real credentials out of source control):
# CONN_STR = ("DRIVER={ODBC Driver 17 for SQL Server};"
#             f"SERVER={SERVER};DATABASE={DATABASE};UID=youruser;PWD=yourpass;")

BATCH_SIZE = 1000
# -----------------------------------------


def detect_delimiter(sample: str) -> str:
    try:
        return csv.Sniffer().sniff(sample, delimiters=",\t|;").delimiter
    except csv.Error:
        return ","


def clean_col(name: str, idx: int) -> str:
    """Make a safe, non-empty SQL column name."""
    name = (name or "").strip()
    name = re.sub(r"[^\w]", "_", name)          # replace non-word chars
    if not name or name[0].isdigit():
        name = f"col_{idx}_{name}".rstrip("_")
    return name[:128]  # SQL Server identifier max length


def split_table(table: str):
    if "." in table:
        schema, tbl = table.split(".", 1)
    else:
        schema, tbl = "dbo", table
    return schema.strip("[]"), tbl.strip("[]")


def main():
    if len(sys.argv) < 3:
        print("Usage: python dat_to_sqlserver.py <file.dat> <schema.Table>")
        sys.exit(1)

    dat_path = Path(sys.argv[1])
    schema, table = split_table(sys.argv[2])
    full_table = f"[{schema}].[{table}]"

    with open(dat_path, "r", encoding="utf-8-sig", errors="replace", newline="") as fin:
        sample = fin.read(8192)
        fin.seek(0)
        delim = detect_delimiter(sample)
        reader = csv.reader(fin, delimiter=delim)

        header = next(reader)
        cols = [clean_col(c, i) for i, c in enumerate(header)]
        # de-duplicate column names
        seen = {}
        for i, c in enumerate(cols):
            if c in seen:
                seen[c] += 1
                cols[i] = f"{c}_{seen[c]}"
            else:
                seen[c] = 0

        print(f"Delimiter={delim!r}, {len(cols)} columns detected.")

        conn = pyodbc.connect(CONN_STR, autocommit=False)
        cur = conn.cursor()
        cur.fast_executemany = True

        # Create table: every column NVARCHAR(MAX)
        col_defs = ",\n  ".join(f"[{c}] NVARCHAR(MAX) NULL" for c in cols)
        cur.execute(f"""
            IF OBJECT_ID(N'{schema}.{table}', N'U') IS NOT NULL
                DROP TABLE {full_table};
            CREATE TABLE {full_table} (
              {col_defs}
            );
        """)

        placeholders = ",".join("?" for _ in cols)
        insert_sql = f"INSERT INTO {full_table} ([{'],['.join(cols)}]) VALUES ({placeholders})"

        batch, total = [], 0
        ncols = len(cols)
        for row in reader:
            # pad/trim so ragged rows don't fail
            if len(row) < ncols:
                row = row + [None] * (ncols - len(row))
            elif len(row) > ncols:
                row = row[:ncols]
            batch.append([v if v != "" else None for v in row])
            if len(batch) >= BATCH_SIZE:
                cur.executemany(insert_sql, batch)
                total += len(batch)
                batch.clear()
        if batch:
            cur.executemany(insert_sql, batch)
            total += len(batch)

        conn.commit()
        cur.close()
        conn.close()
        print(f"[OK] Loaded {total} rows into {full_table}.")


if __name__ == "__main__":
    main()
