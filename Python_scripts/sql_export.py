"""
sql_export.py
Export SQL Server data (a table or a custom query) to CSV, and optionally to
Excel split into parts of at most 800,000 (8 lakh) rows each.

------------------------------------------------------------------------------
CONNECTION  (same settings as test_conn.py)
------------------------------------------------------------------------------
    Windows auth (default):   just run it.
    SQL login:                set USE_WINDOWS_AUTH = False and
                                  $env:SQL_PASSWORD = 'your_sa_password'

------------------------------------------------------------------------------
RUN COMMANDS
------------------------------------------------------------------------------
    # export a whole table to CSV  (CSV is the default when no format given)
    python sql_export.py --table dbo.MyTable

    # export the result of a query to CSV
    python sql_export.py --query "SELECT * FROM dbo.MyTable WHERE Year = 2025"

    # Excel only, split into parts of <= 800,000 rows each
    python sql_export.py --table dbo.MyTable --excel

    # both CSV and Excel
    python sql_export.py --table dbo.MyTable --csv --excel

    # choose the output folder / base name and Excel chunk size
    python sql_export.py --table dbo.MyTable --excel ^
        --out C:\\exports\\mytable --rows-per-part 800000

    # open the point-and-click window (no arguments needed)
    python sql_export.py --gui
    python sql_export.py            (also opens the GUI when no args)

Output:
    <base>.csv                     always written (all rows, one file)
    <base>_Set_0001.xlsx           only with --excel; each <= rows-per-part
    <base>_Set_0002.xlsx
    ...
    <base>_report.txt              summary of the export

NOTE (data handling): exported files may contain confidential or personal
data. Save them only to the appropriate secured location -- never to the
desktop -- and handle through authorized systems only.
------------------------------------------------------------------------------
"""
import argparse
import csv
import os
import sys
import time
from pathlib import Path

import pyodbc

# ---- connection settings (match test_conn.py) ----
SERVER   = r"localhost\SQLEXPRESS"
DATABASE = "sts_db"
SQL_USER = "sa"
DRIVER   = "{ODBC Driver 17 for SQL Server}"
USE_WINDOWS_AUTH = True
SQL_PASSWORD = os.environ.get("SQL_PASSWORD", "")
# ---------------------------------------------------

# Excel hard limit is 1,048,576 rows/sheet; keep parts well under it.
DEFAULT_ROWS_PER_PART = 800_000        # 8 lakh
FETCH_BATCH = 5_000                    # rows pulled from the cursor at a time


def connect(args):
    base = f"DRIVER={DRIVER};SERVER={args.server};DATABASE={args.database};"
    if args.use_windows_auth:
        conn_str = base + "Trusted_Connection=yes;"
    else:
        if not args.password:
            raise RuntimeError(
                "SQL login selected but no password given "
                "(set $env:SQL_PASSWORD or fill the Password field).")
        conn_str = base + f"UID={args.user};PWD={args.password};"
    return pyodbc.connect(conn_str)


def build_sql(args):
    if args.query:
        return args.query
    # Basic guard so a stray identifier can't inject; allow schema.table only.
    tbl = args.table.strip()
    if not all(c.isalnum() or c in "._[]" for c in tbl):
        raise ValueError(f"unexpected characters in table name: {tbl!r}")
    return f"SELECT * FROM {tbl}"


def progress_line(rows, start_time, part=None):
    elapsed = time.time() - start_time
    rate = rows / elapsed if elapsed > 0 else 0
    tail = f" | part {part}" if part else ""
    return f"{rows:,} rows | elapsed {elapsed:4.0f}s | {rate:,.0f} rows/s{tail}"


def export(args, log=print, progress=None):
    """Run the export.

    log      : callback for messages (defaults to print).
    progress : optional callback(text) for the live row counter; if None,
               the counter is written to stdout in place.
    """
    def show_progress(rows, start_time, part=None):
        text = progress_line(rows, start_time, part)
        if progress is not None:
            progress(text)
        else:
            sys.stdout.write("\r" + text)
            sys.stdout.flush()

    if not args.csv and not args.excel:
        raise RuntimeError("Nothing to export: choose CSV, Excel, or both.")

    base = Path(args.out) if args.out else Path("export")
    base.parent.mkdir(parents=True, exist_ok=True)
    csv_path = base.with_suffix(".csv") if args.csv else None
    sql = build_sql(args)
    start_time = time.time()

    if args.excel:
        try:
            import openpyxl  # noqa: F401
        except ImportError:
            raise RuntimeError(
                "Excel output needs openpyxl.  Install it:  pip install openpyxl")

    cn = connect(args)
    cur = cn.cursor()
    log(f"Running query...\n  {sql}")
    cur.execute(sql)
    headers = [d[0] for d in cur.description]
    ncols = len(headers)

    total_rows = 0
    excel_parts = []

    # Excel state (write-only workbook = low memory for large exports)
    wb = ws = None
    part_no = 0
    rows_in_part = 0

    def open_new_part():
        nonlocal wb, ws, part_no, rows_in_part
        from openpyxl import Workbook
        part_no += 1
        wb = Workbook(write_only=True)
        ws = wb.create_sheet(title="Data")
        ws.append(headers)
        rows_in_part = 0

    def close_part():
        nonlocal wb
        if wb is None:
            return
        part_path = base.with_name(f"{base.name}_Set_{part_no:04d}.xlsx")
        wb.save(part_path)
        excel_parts.append((part_path, rows_in_part))
        wb = None

    # Open the CSV file only if CSV output was requested.
    fout = open(csv_path, "w", encoding="utf-8-sig", newline="") if args.csv else None
    writer = csv.writer(fout) if fout else None
    try:
        if writer:
            writer.writerow(headers)
        if args.excel:
            open_new_part()

        while True:
            batch = cur.fetchmany(FETCH_BATCH)
            if not batch:
                break
            for row in batch:
                values = list(row)
                if writer:
                    writer.writerow(values)

                if args.excel:
                    if rows_in_part >= args.rows_per_part:
                        close_part()
                        open_new_part()
                    ws.append(["" if v is None else v for v in values])
                    rows_in_part += 1

                total_rows += 1
                if total_rows % FETCH_BATCH == 0:
                    show_progress(total_rows, start_time,
                                  part_no if args.excel else None)
    finally:
        if fout:
            fout.close()

    if args.excel:
        close_part()

    cn.close()
    show_progress(total_rows, start_time, part_no if args.excel else None)
    elapsed = time.time() - start_time

    log(f"\nRows exported:    {total_rows:,}")
    log(f"Columns exported: {ncols}")
    if args.csv:
        log(f"CSV file created: {csv_path}")
    if args.excel:
        log(f"Excel parts:      {len(excel_parts)} "
            f"(<= {args.rows_per_part:,} rows each)")
        for p, n in excel_parts:
            log(f"   {p.name}  ({n:,} rows)")

    write_report(base, sql, headers, total_rows, excel_parts, args, elapsed, log)
    return csv_path, total_rows, excel_parts


def write_report(base, sql, headers, total_rows, excel_parts, args, elapsed,
                 log=print):
    report_path = base.with_name(base.name + "_report.txt")
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    with open(report_path, "w", encoding="utf-8") as r:
        r.write("=" * 70 + "\n")
        r.write("SQL -> CSV / EXCEL EXPORT REPORT\n")
        r.write("=" * 70 + "\n")
        r.write(f"Generated       : {ts}\n")
        r.write(f"Server / DB     : {args.server} / {args.database}\n")
        r.write(f"Query           : {sql}\n")
        r.write(f"Processing time : {elapsed:.1f} s\n")
        r.write(f"Total data rows : {total_rows:,}\n")
        r.write(f"Total columns   : {len(headers)}\n")
        if args.csv:
            r.write(f"CSV output      : {base.with_suffix('.csv')}\n")
        if args.excel:
            r.write(f"Rows per part   : {args.rows_per_part:,}\n")
            r.write(f"Excel parts     : {len(excel_parts)}\n")
            for p, n in excel_parts:
                r.write(f"   {p.name:<30} {n:,} rows\n")
        r.write("\n" + "-" * 70 + "\n")
        r.write("COLUMNS\n")
        r.write("-" * 70 + "\n")
        for h in headers:
            r.write(f"   {h}\n")
        r.write("\n" + "=" * 70 + "\n")
    log(f"Report written:   {report_path}")


# =============================================================================
# TKINTER GUI
# =============================================================================
def run_gui():
    import queue
    import threading
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    class Args:
        """Plain holder matching what export() expects."""

    root = tk.Tk()
    root.title("SQL -> CSV / Excel Export")
    root.geometry("720x780")
    root.minsize(720, 620)

    pad = {"padx": 8, "pady": 4}
    row = 0

    def add_label(text, r):
        tk.Label(root, text=text, anchor="w").grid(
            row=r, column=0, sticky="w", **pad)

    # --- connection ---
    tk.Label(root, text="CONNECTION", font=("Segoe UI", 10, "bold")).grid(
        row=row, column=0, sticky="w", **pad); row += 1

    add_label("Server", row)
    server_var = tk.StringVar(value=SERVER)
    tk.Entry(root, textvariable=server_var, width=55).grid(
        row=row, column=1, columnspan=2, sticky="w", **pad); row += 1

    add_label("Database", row)
    db_var = tk.StringVar(value=DATABASE)
    tk.Entry(root, textvariable=db_var, width=55).grid(
        row=row, column=1, columnspan=2, sticky="w", **pad); row += 1

    add_label("Authentication", row)
    auth_var = tk.StringVar(value="windows" if USE_WINDOWS_AUTH else "sql")
    auth_frame = tk.Frame(root)
    auth_frame.grid(row=row, column=1, columnspan=2, sticky="w", **pad)
    tk.Radiobutton(auth_frame, text="Windows", variable=auth_var,
                   value="windows").pack(side="left")
    tk.Radiobutton(auth_frame, text="SQL login", variable=auth_var,
                   value="sql").pack(side="left"); row += 1

    add_label("SQL user", row)
    user_var = tk.StringVar(value=SQL_USER)
    tk.Entry(root, textvariable=user_var, width=25).grid(
        row=row, column=1, sticky="w", **pad); row += 1

    add_label("SQL password", row)
    pwd_var = tk.StringVar(value=SQL_PASSWORD)
    tk.Entry(root, textvariable=pwd_var, width=25, show="*").grid(
        row=row, column=1, sticky="w", **pad); row += 1

    # --- source ---
    tk.Label(root, text="SOURCE", font=("Segoe UI", 10, "bold")).grid(
        row=row, column=0, sticky="w", **pad); row += 1

    add_label("Mode", row)
    mode_var = tk.StringVar(value="table")
    mode_frame = tk.Frame(root)
    mode_frame.grid(row=row, column=1, columnspan=2, sticky="w", **pad)
    tk.Radiobutton(mode_frame, text="Table", variable=mode_var,
                   value="table").pack(side="left")
    tk.Radiobutton(mode_frame, text="Query", variable=mode_var,
                   value="query").pack(side="left"); row += 1

    add_label("Table (schema.table)", row)
    table_var = tk.StringVar(value="dbo.")
    tk.Entry(root, textvariable=table_var, width=55).grid(
        row=row, column=1, columnspan=2, sticky="w", **pad); row += 1

    add_label("Query (SELECT ...)", row)
    query_text = tk.Text(root, width=55, height=3)
    query_text.grid(row=row, column=1, columnspan=2, sticky="w", **pad); row += 1

    # --- output ---
    tk.Label(root, text="OUTPUT", font=("Segoe UI", 10, "bold")).grid(
        row=row, column=0, sticky="w", **pad); row += 1

    add_label("Output base (no extension)", row)
    out_var = tk.StringVar(value="")
    tk.Entry(root, textvariable=out_var, width=45).grid(
        row=row, column=1, sticky="w", **pad)

    def browse():
        f = filedialog.asksaveasfilename(
            title="Choose output base name", defaultextension="",
            filetypes=[("All files", "*.*")])
        if f:
            out_var.set(f)
    tk.Button(root, text="Browse...", command=browse).grid(
        row=row, column=2, sticky="w", **pad); row += 1

    add_label("Formats", row)
    fmt_frame = tk.Frame(root)
    fmt_frame.grid(row=row, column=1, columnspan=2, sticky="w", **pad)
    csv_var = tk.BooleanVar(value=True)
    excel_var = tk.BooleanVar(value=False)
    tk.Checkbutton(fmt_frame, text="CSV (single file)",
                   variable=csv_var).pack(side="left")
    tk.Checkbutton(fmt_frame, text="Excel (split into Set_0001, ...)",
                   variable=excel_var).pack(side="left"); row += 1

    add_label("Rows per Excel part", row)
    rpp_var = tk.StringVar(value=str(DEFAULT_ROWS_PER_PART))
    tk.Entry(root, textvariable=rpp_var, width=15).grid(
        row=row, column=1, sticky="w", **pad); row += 1

    # --- run button (kept above the log so it is always visible) ---
    run_btn = tk.Button(root, text="Run export", height=2,
                        bg="#0a7d28", fg="white",
                        font=("Segoe UI", 10, "bold"))
    run_btn.grid(row=row, column=0, columnspan=3, sticky="we",
                 padx=8, pady=8); row += 1

    # --- log box (expands with the window) ---
    log_box = tk.Text(root, width=90, height=10, state="disabled",
                      bg="#1e1e1e", fg="#d4d4d4")
    log_box.grid(row=row, column=0, columnspan=3, sticky="nsew", **pad)
    root.rowconfigure(row, weight=1)
    root.columnconfigure(1, weight=1)
    row += 1

    msg_queue = queue.Queue()

    def enqueue(text, replace_last=False):
        msg_queue.put((text, replace_last))

    def drain_queue():
        while not msg_queue.empty():
            text, replace_last = msg_queue.get_nowait()
            log_box.config(state="normal")
            if replace_last:
                # overwrite the last line (live counter)
                log_box.delete("end-2l", "end-1l")
            log_box.insert("end", text + "\n")
            log_box.see("end")
            log_box.config(state="disabled")
        root.after(100, drain_queue)

    def worker(args):
        try:
            export(args,
                   log=lambda m: enqueue(m),
                   progress=lambda m: enqueue(m, replace_last=True))
            enqueue("\n*** DONE ***")
            enqueue("Reminder: save exports to the secured Global Insider "
                    "folder, not the desktop.")
        except Exception as e:  # surface any failure in the log + a dialog
            enqueue(f"\n[ERROR] {e}")
            root.after(0, lambda: messagebox.showerror("Export failed", str(e)))
        finally:
            root.after(0, lambda: run_btn.config(state="normal"))

    def on_run():
        args = Args()
        args.server = server_var.get().strip()
        args.database = db_var.get().strip()
        args.use_windows_auth = (auth_var.get() == "windows")
        args.user = user_var.get().strip()
        args.password = pwd_var.get()
        args.table = table_var.get().strip() if mode_var.get() == "table" else None
        args.query = (query_text.get("1.0", "end").strip()
                      if mode_var.get() == "query" else None)
        args.out = out_var.get().strip() or None
        args.csv = csv_var.get()
        args.excel = excel_var.get()
        try:
            args.rows_per_part = int(rpp_var.get())
        except ValueError:
            messagebox.showerror("Invalid input", "Rows per part must be a number.")
            return

        if not args.csv and not args.excel:
            messagebox.showerror("Missing input",
                                 "Choose at least one format: CSV, Excel, or both.")
            return
        if args.rows_per_part > 1_048_575:
            messagebox.showerror(
                "Invalid input",
                "Rows per part must be <= 1,048,575 (Excel sheet limit).")
            return
        if mode_var.get() == "table" and not (args.table and args.table != "dbo."):
            messagebox.showerror("Missing input", "Enter a table name.")
            return
        if mode_var.get() == "query" and not args.query:
            messagebox.showerror("Missing input", "Enter a query.")
            return

        run_btn.config(state="disabled")
        threading.Thread(target=worker, args=(args,), daemon=True).start()

    run_btn.config(command=on_run)

    root.after(100, drain_queue)
    root.mainloop()


def main():
    p = argparse.ArgumentParser(
        description="Export SQL Server data to CSV, and optionally to "
                    "Excel split into parts of <= 800,000 rows each.")
    src = p.add_mutually_exclusive_group()
    src.add_argument("--table", help="table to export, e.g. dbo.MyTable")
    src.add_argument("--query", help="custom SELECT query to export")
    p.add_argument("--out", help="output base path/name (default: ./export)")
    p.add_argument("--csv", action="store_true",
                   help="write CSV (single file). Default if no format is given.")
    p.add_argument("--excel", action="store_true",
                   help="write Excel, split into Set_0001, Set_0002, ...")
    p.add_argument("--rows-per-part", type=int, default=DEFAULT_ROWS_PER_PART,
                   help=f"max rows per Excel part (default {DEFAULT_ROWS_PER_PART:,})")
    p.add_argument("--gui", action="store_true", help="open the Tkinter window")
    args = p.parse_args()

    # No source given (or --gui) -> open the GUI.
    if args.gui or (not args.table and not args.query):
        run_gui()
        return

    if args.rows_per_part > 1_048_575:
        print("ERROR: --rows-per-part must be <= 1,048,575 (Excel sheet limit).")
        sys.exit(1)

    # No format flag -> default to CSV (backward compatible).
    if not args.csv and not args.excel:
        args.csv = True

    # Fill the connection fields the CLI path relies on from module defaults.
    args.server = SERVER
    args.database = DATABASE
    args.use_windows_auth = USE_WINDOWS_AUTH
    args.user = SQL_USER
    args.password = SQL_PASSWORD
    export(args)


if __name__ == "__main__":
    main()
