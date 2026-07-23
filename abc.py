"""
Merge repeated "Unique ID" / "File Name" rows into one row per Unique ID.

Input CSV must have headers:
    Unique ID, File Name

For each Unique ID, all its File Name values are concatenated (joined with
SEPARATOR) into a single cell. If the merged text would exceed MAX_CELL_LEN
characters, it overflows into additional columns (File Name 2, File Name 3, ...)
which are added dynamically to the output — only as many extra columns as the
worst-case Unique ID actually needs.

Usage:
    python merge_unique_id_filenames.py input.csv output.csv
    python merge_unique_id_filenames.py input.csv output.csv --sep "; " --max-len 25000
"""

import argparse
import csv
import sys

EXCEL_HARD_LIMIT = 32767  # Excel's absolute max characters per cell


def print_progress(stage, current, total):
    pct = 100 if total == 0 else min(100, int(current * 100 / total))
    bar_width = 30
    filled = int(bar_width * pct / 100)
    bar = "#" * filled + "-" * (bar_width - filled)
    sys.stderr.write(f"\r{stage:<22} [{bar}] {pct:3d}%")
    sys.stderr.flush()
    if pct >= 100:
        sys.stderr.write("\n")


def build_chunks(file_names, separator, max_len):
    """Pack file_names into as few chunks as possible, each <= max_len chars."""
    chunks = []
    current = ""

    for name in file_names:
        if not current:
            candidate = name
        else:
            candidate = current + separator + name

        if len(candidate) <= max_len:
            current = candidate
        else:
            if current:
                chunks.append(current)
            # Single file name longer than max_len on its own: keep it whole
            # in its own chunk rather than truncating/losing data.
            current = name

    if current:
        chunks.append(current)

    return chunks or [""]


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("input_csv", help="Path to source CSV with 'Unique ID' and 'File Name' columns")
    parser.add_argument("output_csv", help="Path to write the merged CSV")
    parser.add_argument("--sep", default="; ", help="Separator used to join file names within a cell (default: '; ')")
    parser.add_argument(
        "--max-len",
        type=int,
        default=25000,
        help="Max characters per merged cell before overflowing into a new column (default: 25000)",
    )
    parser.add_argument(
        "--id-col", default="Unique_ID", help="Name of the unique id column (default: 'Unique_ID')"
    )
    parser.add_argument(
        "--name-col", default="File_Name", help="Name of the file name column (default: 'File_Name')"
    )
    args = parser.parse_args()

    if args.max_len > EXCEL_HARD_LIMIT:
        print(
            f"Warning: --max-len {args.max_len} exceeds Excel's hard limit of "
            f"{EXCEL_HARD_LIMIT} characters per cell. Clamping to {EXCEL_HARD_LIMIT}.",
            file=sys.stderr,
        )
        args.max_len = EXCEL_HARD_LIMIT

    # Count data rows up front so read progress can show a real percentage.
    total_rows = sum(1 for _ in open(args.input_csv, encoding="utf-8-sig")) - 1
    total_rows = max(total_rows, 0)

    # Preserve first-seen order of Unique IDs, and collect file names per id.
    grouped = {}
    order = []

    with open(args.input_csv, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        if args.id_col not in reader.fieldnames or args.name_col not in reader.fieldnames:
            sys.exit(
                f"Input CSV must contain columns '{args.id_col}' and '{args.name_col}'. "
                f"Found: {reader.fieldnames}"
            )
        for i, row in enumerate(reader, start=1):
            uid = row[args.id_col]
            name = row[args.name_col]
            if uid not in grouped:
                grouped[uid] = []
                order.append(uid)
            if name:
                grouped[uid].append(name)
            if i % 500 == 0 or i == total_rows:
                print_progress("Reading rows", i, total_rows)
    if total_rows == 0:
        print_progress("Reading rows", 0, 0)

    # Build chunks per id, tracking the max number of columns needed.
    id_chunks = {}
    max_cols = 1
    total_ids = len(order)
    for i, uid in enumerate(order, start=1):
        chunks = build_chunks(grouped[uid], args.sep, args.max_len)
        id_chunks[uid] = chunks
        max_cols = max(max_cols, len(chunks))
        if i % 200 == 0 or i == total_ids:
            print_progress("Merging IDs", i, total_ids)
    if total_ids == 0:
        print_progress("Merging IDs", 0, 0)

    # Dynamically name columns: File Name, File Name 2, File Name 3, ...
    name_columns = [args.name_col] + [f"{args.name_col} {i}" for i in range(2, max_cols + 1)]

    with open(args.output_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow([args.id_col] + name_columns)
        for i, uid in enumerate(order, start=1):
            chunks = id_chunks[uid]
            padded = chunks + [""] * (max_cols - len(chunks))
            writer.writerow([uid] + padded)
            if i % 200 == 0 or i == total_ids:
                print_progress("Writing output", i, total_ids)
    if total_ids == 0:
        print_progress("Writing output", 0, 0)

    print(f"Done. {len(order)} unique IDs written to '{args.output_csv}' using {max_cols} file name column(s).")


if __name__ == "__main__":
    main()
