"""
compare_columns.py -- diagnose why some columns load as NULL.

It compares the .dat file's header columns (after the same name-cleaning that
dat_to_sqlserver.py does) against the real columns of your SQL Server table,
and prints which ones match and which stay NULL.

Run (Windows auth):
    python compare_columns.py

Edit the settings below to match your load if needed.
"""
import re
import pyodbc

# ---- match these to your load ----
SERVER   = r"localhost\SQLEXPRESS"
DRIVER   = "{ODBC Driver 17 for SQL Server}"
DATABASE = "sts_db"
SCHEMA   = "dbo"
TABLE    = "sts_master"
DAT      = "input.dat"
# ----------------------------------

DELIM = "þ\x14þ"   # þ \x14 þ


def clean_col(name, idx):
    """Same column-name cleaning dat_to_sqlserver.py uses."""
    name = (name or "").strip("þ ").strip()
    name = re.sub(r"[^\w]", "_", name)
    if not name or name[0].isdigit():
        name = f"col_{idx}_{name}".rstrip("_")
    return name[:128]


def main():
    with open(DAT, "r", encoding="utf-8-sig", errors="ignore") as f:
        raw = f.readline().strip().split(DELIM)
    dat_cols = [clean_col(c, i) for i, c in enumerate(raw)]

    conn_str = (f"DRIVER={DRIVER};SERVER={SERVER};DATABASE={DATABASE};"
                f"Trusted_Connection=yes;")
    cn = pyodbc.connect(conn_str, timeout=5)
    cur = cn.cursor()
    cur.execute(
        "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
        "WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ? ORDER BY ORDINAL_POSITION",
        SCHEMA, TABLE)
    tbl_cols = [r[0] for r in cur.fetchall()]
    cn.close()

    tbl_lower = {c.lower(): c for c in tbl_cols}
    print(f"{len(dat_cols)} .dat columns  vs  {len(tbl_cols)} table columns\n")
    print(f"{'#':>3} | {'DAT raw header':<30} | {'cleaned':<25} | match?")
    print("-" * 90)
    matched = 0
    for i, (r, c) in enumerate(zip(raw, dat_cols)):
        r = r.strip("þ ")
        hit = tbl_lower.get(c.lower())
        if hit:
            matched += 1
        status = f"YES -> {hit}" if hit else "NO (stays NULL)"
        print(f"{i:>3} | {r[:30]:<30} | {c[:25]:<25} | {status}")

    print(f"\nMatched {matched}/{len(dat_cols)}")
    print("\nTable columns NOT filled by any .dat column:")
    dat_lower = {c.lower() for c in dat_cols}
    for c in tbl_cols:
        if c.lower() not in dat_lower:
            print("   -", c)


if __name__ == "__main__":
    main()
