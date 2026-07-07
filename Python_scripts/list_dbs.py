"""
list_dbs.py -- show which databases (and your target table) you can access,
using Windows authentication.

Run:
    python list_dbs.py
"""
import pyodbc

SERVER = r"localhost\SQLEXPRESS"   # same value as SERVER in dat_to_sqlserver.py
DRIVER = "{ODBC Driver 17 for SQL Server}"
TABLE_TO_FIND = "sts_master"       # the table you want to load into

# Connect to 'master' -- every login can open it.
conn_str = f"DRIVER={DRIVER};SERVER={SERVER};DATABASE=master;Trusted_Connection=yes;"
cn = pyodbc.connect(conn_str, timeout=5)
cur = cn.cursor()

print("Databases you can see on this server:")
cur.execute("SELECT name FROM sys.databases ORDER BY name")
dbs = [r[0] for r in cur.fetchall()]
for name in dbs:
    print("   -", name)
print()

print(f"Looking for a table named '{TABLE_TO_FIND}' in each database...")
for db in dbs:
    try:
        cur.execute(
            f"SELECT TABLE_SCHEMA, TABLE_NAME FROM [{db}].INFORMATION_SCHEMA.TABLES "
            f"WHERE TABLE_NAME = ?", TABLE_TO_FIND)
        for schema, tbl in cur.fetchall():
            print(f"   FOUND -> {db}.{schema}.{tbl}")
    except pyodbc.Error:
        pass   # no permission on that db -- skip quietly

cn.close()
print("\nUse the database name shown next to FOUND when you run the load, e.g.:")
print(f"   python dat_to_sqlserver.py input.dat <DatabaseName>.dbo.{TABLE_TO_FIND}")
