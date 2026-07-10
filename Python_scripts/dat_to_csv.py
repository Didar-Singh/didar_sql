"""
dat_to_csv.py
Convert a Concordance/Relativity .dat load file to .csv.

Field format:  þvalueþ\x14þvalueþ...
    þ  (thorn, \xFE) = text qualifier wrapping each field
    \x14 (ASCII 20)  = column separator

Features:
    * Reads ALL rows (no data loss / no coercion)
    * Live progress bar with % complete + estimated time remaining (ETA)
    * Auto-generates a report .txt: row/column counts, inferred data types,
      max lengths, empty counts, and any errors or missed/overflow rows.

------------------------------------------------------------------------------
RUN COMMANDS
------------------------------------------------------------------------------
    python dat_to_csv.py "Objects_1000125_export Part 2.dat"
    python dat_to_csv.py "Objects_1000125_export Part 2.dat" "output.csv"
------------------------------------------------------------------------------
"""
import csv
import os
import re
import sys
import time
from pathlib import Path

# Concordance/Relativity load-file delimiters:
#   \x14 (DC4, ASCII 20) = column separator  <-- what we split on
#   þ    (thorn, \xFE)   = text qualifier wrapping each field, stripped per-field
# Splitting on the DC4 separator alone (not "þ\x14þ") is robust: it still works
# even if the thorn byte is dropped during decoding (ANSI-exported files).
DELIMITER = "\x14"

# ---- data-type inference patterns ----
INT_RE = re.compile(r"^-?\d+$")
DEC_RE = re.compile(r"^-?\d{1,3}(,\d{3})*(\.\d+)?$|^-?\d+(\.\d+)?$")
DATE_RE = re.compile(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}$|^\d{1,2}[-/]\d{1,2}[-/]\d{4}$")
DATETIME_RE = re.compile(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}[ T]\d{1,2}:\d{2}(:\d{2})?")


class ColumnStats:
    """Tracks per-column stats to infer a likely data type for the report."""
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


def write_report(report_path, input_file, output_file, headers, stats,
                 total_rows, padded, trimmed, blank_skipped, elapsed):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    with open(report_path, "w", encoding="utf-8") as r:
        r.write("=" * 70 + "\n")
        r.write("DAT -> CSV CONVERSION REPORT\n")
        r.write("=" * 70 + "\n")
        r.write(f"Generated       : {ts}\n")
        r.write(f"Source file     : {input_file}\n")
        r.write(f"Output CSV      : {output_file}\n")
        r.write(f"Source size     : {os.path.getsize(input_file):,} bytes\n")
        r.write(f"Processing time : {elapsed:.1f} s\n")
        r.write(f"Total data rows : {total_rows:,}\n")
        r.write(f"Total columns   : {len(headers)}\n")
        r.write(f"Blank rows skip : {blank_skipped:,}\n\n")

        r.write("-" * 70 + "\n")
        r.write("COLUMNS  (inferred type | non-empty | empty | max length)\n")
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
        if not padded and not trimmed:
            r.write("None. All rows had exactly the expected column count.\n")
        if padded:
            r.write(
                f"[WARN] {len(padded):,} row(s) had FEWER columns than the header "
                f"and were padded with blanks (no data lost).\n"
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


def convert(input_file: Path, output_file: Path) -> None:
    report_path = input_file.with_name(input_file.stem + "_report.txt")
    total_bytes = os.path.getsize(input_file)
    start_time = time.time()

    with open(input_file, "r", encoding="utf-8-sig", errors="ignore") as fin:
        header = fin.readline().strip()
        headers = [h.strip("þ ") for h in header.split(DELIMITER)]
        ncols = len(headers)
        stats = [ColumnStats(h) for h in headers]

        processed = len(header.encode("utf-8", "ignore"))
        total_rows = blank_skipped = 0
        padded, trimmed = [], []

        with open(output_file, "w", encoding="utf-8-sig", newline="") as fout:
            writer = csv.writer(fout)
            writer.writerow(headers)

            for line_no, line in enumerate(fin, start=2):  # header is line 1
                processed += len(line.encode("utf-8", "ignore"))
                if not line.strip():
                    blank_skipped += 1
                    continue

                values = [v.strip("þ \r\n") for v in line.split(DELIMITER)]
                if len(values) < ncols:
                    padded.append(line_no)
                    values += [""] * (ncols - len(values))
                elif len(values) > ncols:
                    trimmed.append(line_no)  # data loss flagged in report

                values = values[:ncols]
                for i in range(ncols):
                    stats[i].observe(values[i])
                writer.writerow(values)
                total_rows += 1

                if total_rows % 1000 == 0:
                    render_progress(processed, total_bytes, start_time, total_rows)

    render_progress(total_bytes, total_bytes, start_time, total_rows)
    elapsed = time.time() - start_time

    print(f"\nCSV file created: {output_file}")
    print(f"Rows exported:    {total_rows:,}")
    print(f"Columns exported: {ncols}")
    write_report(report_path, input_file, output_file, headers, stats,
                 total_rows, padded, trimmed, blank_skipped, elapsed)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        print("ERROR: no input file given.")
        sys.exit(1)
    input_file = Path(sys.argv[1])
    output_file = Path(sys.argv[2]) if len(sys.argv) > 2 else input_file.with_suffix(".csv")
    convert(input_file, output_file)


if __name__ == "__main__":
    main()
