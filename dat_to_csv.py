"""
dat_to_csv.py
Convert delimited .dat file(s) to .csv, auto-detecting the delimiter.

Usage:
    python dat_to_csv.py "C:\\path\\to\\file.dat"
    python dat_to_csv.py "C:\\path\\to\\folder"   (converts every .dat in the folder)
"""
import csv
import sys
from pathlib import Path


def detect_delimiter(sample: str) -> str:
    try:
        return csv.Sniffer().sniff(sample, delimiters=",\t|;").delimiter
    except csv.Error:
        return ","  # fallback


def convert(dat_path: Path) -> Path:
    csv_path = dat_path.with_suffix(".csv")
    # utf-8-sig + errors="replace" so no row/char is lost on odd encodings
    with open(dat_path, "r", encoding="utf-8-sig", errors="replace", newline="") as fin:
        sample = fin.read(8192)
        fin.seek(0)
        delim = detect_delimiter(sample)
        reader = csv.reader(fin, delimiter=delim)
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as fout:
            writer = csv.writer(fout)  # writes proper comma-quoted CSV
            row_count = 0
            for row in reader:
                writer.writerow(row)
                row_count += 1
    print(f"[OK] {dat_path.name}: delimiter={delim!r}, rows={row_count} -> {csv_path.name}")
    return csv_path


def main():
    if len(sys.argv) < 2:
        print("Usage: python dat_to_csv.py <file.dat | folder>")
        sys.exit(1)
    target = Path(sys.argv[1])
    files = list(target.glob("*.dat")) if target.is_dir() else [target]
    if not files:
        print("No .dat files found.")
        return
    for f in files:
        convert(f)


if __name__ == "__main__":
    main()
