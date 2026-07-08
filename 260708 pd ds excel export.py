"""
============================================================
Export two SQL Server results to Excel
  Sheet 1: Space  -> rows with trim/space/clean issues
  Sheet 2: FlipName -> rows where Flip Name Count > 1
Read-only. Source table is never modified.

Requirements (install once):
    pip install pyodbc pandas openpyxl

The "Space" query is built dynamically in Python by reading the
table's column list, so it scales to 50+ columns automatically.
============================================================
"""

import os
import sys
import pandas as pd
import pyodbc

# ------------------------------------------------------------
# 1) CONFIG  - edit these values
#    (connection settings mirror the working dat_to_sqlserver.py)
# ------------------------------------------------------------
SERVER   = r"prdenvfdevm-3\MSSQLSERVER01"  # same instance as dat_to_sqlserver.py
DATABASE = "sts_db"                        # <-- your database
SCHEMA   = "dbo"
TABLE    = "sts_master_core"               # <-- your table
DRIVER   = "ODBC Driver 17 for SQL Server"

# Column names used by the Flip Name query (edit to match your table)
UNIQUE_ID_COL = "Unique_ID"                # <-- your real unique id column
DOCID_COL     = "DOCID"
FIRST_NAME    = "First Name"
LAST_NAME     = "Last Name"

# Output workbook (save to your Global Insider folder, not the desktop)
OUTPUT_XLSX = "260708 re ds space and flip name.xlsx"
EXCEL_MAX_ROWS = 1_048_576                 # per-sheet hard limit

# ------------------------------------------------------------
# AUTH - same approach as the working dat_to_sqlserver.py
#   USE_WINDOWS_AUTH = True  -> Trusted_Connection (like SSMS Windows login)
#   USE_WINDOWS_AUTH = False -> SQL login; set the password via env var:
#       PowerShell:  $env:SQL_PASSWORD = 'yourpassword'
# ------------------------------------------------------------
USE_WINDOWS_AUTH = True
SQL_USER = "sa"
SQL_PASSWORD = os.environ.get("SQL_PASSWORD", "")   # never hard-code a real password

_base = f"DRIVER={{{DRIVER}}};SERVER={SERVER};DATABASE={DATABASE};"
if USE_WINDOWS_AUTH:
    CONN_STR = _base + "Trusted_Connection=yes;"
else:
    CONN_STR = _base + f"UID={SQL_USER};PWD={SQL_PASSWORD};"


def q(name: str) -> str:
    """Safely bracket-quote an identifier."""
    return "[" + name.replace("]", "]]") + "]"


FULL_TABLE = f"{q(SCHEMA)}.{q(TABLE)}"


# ------------------------------------------------------------
# 2) Build the "Space" query dynamically (all text columns)
# ------------------------------------------------------------
def build_space_query(conn) -> str:
    cols = pd.read_sql(
        """
        SELECT COLUMN_NAME
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ?
          AND DATA_TYPE IN ('char','varchar','nchar','nvarchar','text','ntext')
        ORDER BY ORDINAL_POSITION
        """,
        conn, params=[SCHEMA, TABLE],
    )["COLUMN_NAME"].tolist()

    if not cols:
        raise RuntimeError("No character columns found to check.")

    def dirty_case(col: str) -> str:
        c = q(col)
        lit = col.replace("'", "''")
        return (
            f"        CASE WHEN {c} IS NOT NULL AND ("
            f"{c} <> LTRIM(RTRIM({c}))"
            f" OR {c} LIKE '%  %'"
            f" OR {c} LIKE '%' + CHAR(9) + '%'"
            f" OR {c} LIKE '%' + CHAR(13) + '%'"
            f" OR {c} LIKE '%' + CHAR(10) + '%'"
            f" OR {c} LIKE '%' + CHAR(160) + '%')"
            f" THEN '{lit}' END"
        )

    cases = ",\n".join(dirty_case(c) for c in cols)
    return (
        "SELECT * FROM (\n"
        "    SELECT *,\n"
        "        CONCAT_WS('; ',\n"
        f"{cases}\n"
        "        ) AS [Space]\n"
        f"    FROM {FULL_TABLE}\n"
        ") AS x\n"
        "WHERE x.[Space] <> ''"
    )


# ------------------------------------------------------------
# 3) Build the "Flip Name" query
# ------------------------------------------------------------
def build_flip_query() -> str:
    uid, doc = q(UNIQUE_ID_COL), q(DOCID_COL)
    fn, ln = q(FIRST_NAME), q(LAST_NAME)
    return f"""
SELECT
    q.{uid}, q.{doc},
    q.[Full Name Count], q.[Flip Name Count],
    q.[Full Name], q.[Flip Name],
    q.*
FROM (
    SELECT
        t.*,
        COUNT(*) OVER (PARTITION BY fn.FullName) AS [Full Name Count],
        COUNT(*) OVER (PARTITION BY fx.FlipName) AS [Flip Name Count],
        fn.FullName AS [Full Name],
        fx.FlipName AS [Flip Name]
    FROM {FULL_TABLE} AS t
    CROSS APPLY (
        SELECT LTRIM(RTRIM(REPLACE(REPLACE(
                 LTRIM(RTRIM(ISNULL(t.{fn}, ''))) + N' ' +
                 LTRIM(RTRIM(ISNULL(t.{ln},  ''))),
               CHAR(160), N' '), CHAR(9), N' '))) AS FullName
    ) AS fn
    CROSS APPLY (
        SELECT STRING_AGG(s.[value], N' ') WITHIN GROUP (ORDER BY s.[value]) AS FlipName
        FROM STRING_SPLIT(fn.FullName, N' ') AS s
        WHERE s.[value] <> N''
    ) AS fx
) AS q
WHERE q.[Flip Name Count] > 1
ORDER BY q.[Flip Name], q.[Full Name], q.{uid}
"""


# ------------------------------------------------------------
# 4) Run and export
# ------------------------------------------------------------
def dedupe_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Excel/openpyxl dislikes duplicate headers; suffix repeats."""
    seen, new_cols = {}, []
    for col in df.columns:
        if col in seen:
            seen[col] += 1
            new_cols.append(f"{col}.{seen[col]}")
        else:
            seen[col] = 0
            new_cols.append(col)
    df.columns = new_cols
    return df


def main() -> None:
    print(f"Connecting to {SERVER} / {DATABASE} ...")
    with pyodbc.connect(CONN_STR) as conn:
        print("Building + running SPACE query ...")
        df_space = pd.read_sql(build_space_query(conn), conn)
        print(f"  Space rows: {len(df_space):,}")

        print("Running FLIP NAME query ...")
        df_flip = dedupe_columns(pd.read_sql(build_flip_query(), conn))
        print(f"  Flip Name rows: {len(df_flip):,}")

    results = {"Space": df_space, "FlipName": df_flip}

    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as xl:
        wrote_any = False
        for sheet, df in results.items():
            if len(df) > EXCEL_MAX_ROWS:
                csv_name = OUTPUT_XLSX.replace(".xlsx", f" {sheet}.csv")
                df.to_csv(csv_name, index=False, encoding="utf-8-sig")
                print(f"  {sheet}: {len(df):,} rows exceed Excel limit -> {csv_name}")
            else:
                df.to_excel(xl, sheet_name=sheet, index=False)
                wrote_any = True
        if not wrote_any:
            # ExcelWriter needs at least one sheet
            pd.DataFrame({"note": ["All results exported to CSV (too large for Excel)"]}
                         ).to_excel(xl, sheet_name="README", index=False)

    print(f"Done -> {OUTPUT_XLSX}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
