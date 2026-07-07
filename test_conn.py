"""
test_conn.py  --  diagnose the SQL Server connection for dat_to_sqlserver.py

Run:
    python test_conn.py
Optionally set the sa password first (recommended, don't hardcode it):
    $env:SQL_PASSWORD = 'your_actual_sa_password'
    python test_conn.py

It prints which drivers are installed and which auth method connects, then
tells you what to use in dat_to_sqlserver.py.
"""
import os
import pyodbc

# ---- match these to dat_to_sqlserver.py ----
SERVER   = r"localhost\SQLEXPRESS"   # same value as SERVER in the main script
DATABASE = "sts_db"                  # the database your table lives in
SQL_USER = "sa"
SQL_PASSWORD = os.environ.get("SQL_PASSWORD", "")   # set $env:SQL_PASSWORD first
DRIVER = "{ODBC Driver 17 for SQL Server}"
# --------------------------------------------


def show_drivers():
    print("Installed ODBC drivers:")
    for d in pyodbc.drivers():
        print("   -", d)
    print()


def try_conn(label, conn_str):
    # Never print the password.
    safe = conn_str.replace(SQL_PASSWORD, "***") if SQL_PASSWORD else conn_str
    print(f"[TEST] {label}")
    print(f"       {safe}")
    try:
        cn = pyodbc.connect(conn_str, timeout=5)
        cur = cn.cursor()
        cur.execute("SELECT SUSER_SNAME(), DB_NAME(), @@VERSION")
        who, db, ver = cur.fetchone()
        print(f"       [OK] connected as '{who}' to database '{db}'")
        print(f"       {ver.splitlines()[0]}")
        cn.close()
        return True
    except pyodbc.Error as e:
        print(f"       [FAIL] {e}")
        return False
    finally:
        print()


def main():
    show_drivers()

    base = f"DRIVER={DRIVER};SERVER={SERVER};DATABASE={DATABASE};"

    win_ok = try_conn(
        "Windows authentication (USE_WINDOWS_AUTH = True)",
        base + "Trusted_Connection=yes;",
    )

    sql_ok = False
    if SQL_PASSWORD:
        sql_ok = try_conn(
            f"SQL login '{SQL_USER}' (USE_WINDOWS_AUTH = False)",
            base + f"UID={SQL_USER};PWD={SQL_PASSWORD};",
        )
    else:
        print("[SKIP] SQL login test -- no password set.")
        print("       Set it first:  $env:SQL_PASSWORD = 'your_sa_password'\n")

    print("=" * 60)
    if win_ok:
        print("USE THIS -> set USE_WINDOWS_AUTH = True  (no password needed)")
    elif sql_ok:
        print("USE THIS -> set USE_WINDOWS_AUTH = False and set $env:SQL_PASSWORD")
    else:
        print("Neither worked. Check the notes below.")
        print(" - Is SERVER correct? Try 'localhost', '.\\SQLEXPRESS', or the")
        print("   machine name. In SSMS, copy the exact 'Server name' you use.")
        print(" - Login failed (18456) with the right password usually means")
        print("   SQL auth is disabled (server is Windows-only) -> use Windows auth.")
        print(" - Is the SQL Server service running? Is the instance name right?")
    print("=" * 60)


if __name__ == "__main__":
    main()
