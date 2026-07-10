"""
============================================================
260710 pd ds merge data.py
Person de-duplication / merge for [cng_db].[dbo].[cng_dedup]

Reads First/Middle/Last/Suffix, Full DOB (MM/DD/YYYY) and SSN, groups
records that represent the SAME PERSON into clusters (a person_id), and
writes:
    Sheet "Clusters" -> every input row + person_id + why it matched
    Sheet "Merged"   -> one "golden" (most complete) row per person

Matching rules (a pair of rows is the same person if ANY hold):
    R1  Same SSN (both present, equal)               -> same, even if names differ
    R2  Same fuzzy name + SSN present on one, blank on other
    R3  Same fuzzy name + one has DOB, other has SSN
    R4  Same fuzzy name (First+Middle+Last+Suffix)
    R6  Middle initial vs full ("H" / "Harish"), rest of name equal
    (R5 = same SSN with/without suffix -> already covered by R1)

"Fuzzy name" = Last equal AND First equal-or-prefix (min 3 chars) AND
               Middle equal-or-blank-or-prefix AND Suffix equal-or-blank.

Clustering is TRANSITIVE via union-find: A~B by SSN and B~C by name+DOB
puts A, B and C in one person.

READ-ONLY: the source table is never modified. Output goes to a workbook
you save in the secured Global Insider folder (never the desktop).

Install once:
    pip install pyodbc pandas openpyxl

Run:
    # Windows auth (default). For SQL login:  $env:SQL_PASSWORD = 'your_sa_password'
    python "260710 pd ds merge data.py"
============================================================
"""

import os
import re
import sys
import pandas as pd
import pyodbc

# ------------------------------------------------------------
# 1) CONFIG  - edit these to match your environment
# ------------------------------------------------------------
SERVER   = r"prdenvfdevm-3\MSSQLSERVER01"   # same instance style as your other scripts
DATABASE = "cng_db"
SCHEMA   = "dbo"
TABLE    = "cng_dedup"
DRIVER   = "ODBC Driver 17 for SQL Server"

# Column names exactly as they appear in the table
COL_FIRST  = "First Name"
COL_LAST   = "Last Name"
COL_MIDDLE = "Middle Name"
COL_SUFFIX = "Suffix"
COL_DOB    = "Full Date of Birth (MM/DD/YYYY)"
COL_SSN    = "Social Security Number"

# Tuning knobs
FIRST_MIN_PREFIX = 3      # "Saj" (>=3) counts as a prefix of "Sajan"; "Jo" would not

OUTPUT_XLSX = "260710 re ds merge data results.xlsx"
EXCEL_MAX_ROWS = 1_048_576

# ------------------------------------------------------------
# 2) AUTH - Windows auth first, SQL login (sa) fallback
#    (mirrors your dat_to_sqlserver.py / excel export.py)
# ------------------------------------------------------------
USE_WINDOWS_AUTH = True
SQL_USER = "sa"
SQL_PASSWORD = os.environ.get("SQL_PASSWORD", "")   # never hard-code a real password

_base = f"DRIVER={{{DRIVER}}};SERVER={SERVER};DATABASE={DATABASE};"
CONN_WINDOWS = _base + "Trusted_Connection=yes;"
CONN_SQL = _base + f"UID={SQL_USER};PWD={SQL_PASSWORD};"


def connect():
    attempts = []
    if USE_WINDOWS_AUTH:
        attempts.append(("Windows auth", CONN_WINDOWS))
    if SQL_PASSWORD:
        attempts.append((f"SQL login ({SQL_USER})", CONN_SQL))
    if not attempts:
        raise SystemExit(
            "No usable auth. Keep USE_WINDOWS_AUTH = True, or set the password:\n"
            "  $env:SQL_PASSWORD = 'yourSApassword'  then re-run."
        )
    last_err = None
    for label, conn_str in attempts:
        try:
            print(f"Trying {label} ...")
            return pyodbc.connect(conn_str)
        except pyodbc.Error as exc:
            last_err = exc
            print(f"  {label} failed: {exc.args[1] if len(exc.args) > 1 else exc}")
    raise SystemExit(f"\nCould not connect. Last error:\n  {last_err}")


def q(name: str) -> str:
    """Safely bracket-quote an identifier."""
    return "[" + name.replace("]", "]]") + "]"


FULL_TABLE = f"{q(SCHEMA)}.{q(TABLE)}"


# ------------------------------------------------------------
# 3) Normalisation helpers
# ------------------------------------------------------------
def norm_text(v) -> str:
    """Upper-case, trim, collapse internal whitespace. NULL/blank -> ''."""
    if v is None:
        return ""
    s = str(v).replace(" ", " ")          # non-breaking space -> space
    s = re.sub(r"\s+", " ", s).strip().upper()
    return "" if s.lower() in ("nan", "none", "null") else s


def norm_ssn(v) -> str:
    """Keep digits only. Reject blanks and obvious non-SSNs (all same digit,
    not 9 digits) so junk SSNs never merge two different people."""
    if v is None:
        return ""
    digits = re.sub(r"\D", "", str(v))
    if len(digits) != 9:
        return ""
    if digits == "000000000" or len(set(digits)) == 1:
        return ""
    return digits


def norm_dob(v) -> str:
    """Return a canonical YYYYMMDD string, or '' if not parseable."""
    if v is None:
        return ""
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return ""
    ts = pd.to_datetime(s, errors="coerce")
    return "" if pd.isna(ts) else ts.strftime("%Y%m%d")


# ------------------------------------------------------------
# 4) Fuzzy name comparison
# ------------------------------------------------------------
def first_match(a: str, b: str) -> bool:
    """Equal, or one is a prefix of the other (min FIRST_MIN_PREFIX chars)."""
    if not a or not b:
        return False
    if a == b:
        return True
    short, long = (a, b) if len(a) <= len(b) else (b, a)
    return len(short) >= FIRST_MIN_PREFIX and long.startswith(short)


def part_match(a: str, b: str) -> bool:
    """Middle-name style: equal, either blank, or one is a prefix/initial of
    the other (covers 'H' vs 'HARISH')."""
    if a == b:
        return True
    if not a or not b:
        return True                     # one side blank -> not a conflict
    short, long = (a, b) if len(a) <= len(b) else (b, a)
    return long.startswith(short)       # H -> HARISH, HAR -> HARISH


def suffix_match(a: str, b: str) -> bool:
    """Equal, or one side blank (suffix present on one record only)."""
    return a == b or not a or not b


def name_match(r1, r2) -> bool:
    """Full fuzzy-name equality used by R2/R3/R4/R6."""
    return (
        r1["last"] and r1["last"] == r2["last"]        # last name must be equal
        and first_match(r1["first"], r2["first"])
        and part_match(r1["mid"], r2["mid"])
        and suffix_match(r1["suf"], r2["suf"])
    )


# ---- conflict guards (added per rule refinement #1) ----
def name_conflict(r1, r2) -> bool:
    """True when BOTH rows have a real name and they do NOT fuzzy-match.
    This stops a shared SSN from merging two clearly different people
    (Didar Singh vs Harish Singh). A blank name can't conflict."""
    if not (r1["last"] and r2["last"] and r1["first"] and r2["first"]):
        return False
    return not name_match(r1, r2)


def ssn_conflict(r1, r2) -> bool:
    """True when both SSNs are present and different (two different people)."""
    return bool(r1["ssn"]) and bool(r2["ssn"]) and r1["ssn"] != r2["ssn"]


def dob_conflict(r1, r2) -> bool:
    """True when both DOBs are present and different."""
    return bool(r1["dob"]) and bool(r2["dob"]) and r1["dob"] != r2["dob"]


# ------------------------------------------------------------
# 5) Union-Find (disjoint set) for transitive clustering
# ------------------------------------------------------------
class UnionFind:
    def __init__(self, n):
        self.parent = list(range(n))
        self.rank = [0] * n

    def find(self, x):
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]   # path halving
            x = self.parent[x]
        return x

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra == rb:
            return
        if self.rank[ra] < self.rank[rb]:
            ra, rb = rb, ra
        self.parent[rb] = ra
        if self.rank[ra] == self.rank[rb]:
            self.rank[ra] += 1


# ------------------------------------------------------------
# 6) Build clusters
# ------------------------------------------------------------
def build_records(df: pd.DataFrame):
    recs = []
    for i, row in df.iterrows():
        recs.append({
            "idx":   i,
            "first": norm_text(row[COL_FIRST]),
            "last":  norm_text(row[COL_LAST]),
            "mid":   norm_text(row[COL_MIDDLE]),
            "suf":   norm_text(row[COL_SUFFIX]),
            "dob":   norm_dob(row[COL_DOB]),
            "ssn":   norm_ssn(row[COL_SSN]),
        })
    return recs


def _link(uf, reasons, r1, r2, tag):
    uf.union(r1["idx"], r2["idx"])
    reasons[r1["idx"]].add(tag)
    reasons[r2["idx"]].add(tag)


def cluster(recs, uf, reasons):
    """Union records that are the same person; record a reason per link.

    A pair merges only when nothing CONFLICTS (name / SSN / DOB) AND there is
    a positive signal:
      * same SSN  (R1)  -- but only if the names do not conflict
      * fuzzy name match (R2/R3/R4/R6) -- but only if SSN/DOB do not conflict
    """
    # ---- R1: same SSN, blocked by SSN.  Merge within a group only when the
    #          names are compatible (Sajan/Saj) and DOBs don't conflict.
    #          Different names on the same SSN stay SEPARATE (Didar vs Harish).
    by_ssn = {}
    for r in recs:
        if r["ssn"]:
            by_ssn.setdefault(r["ssn"], []).append(r)
    for ssn, group in by_ssn.items():
        m = len(group)
        for a in range(m):
            for b in range(a + 1, m):
                r1, r2 = group[a], group[b]
                if name_conflict(r1, r2) or dob_conflict(r1, r2):
                    continue
                _link(uf, reasons, r1, r2, "R1 same-SSN")

    # ---- R2/R3/R4/R6: fuzzy name inside (last, first-initial) blocks.
    #      Blocking on last + first initial keeps comparisons small while
    #      still allowing prefix matches (a prefix shares the first initial).
    #      Skip if SSN or DOB conflict (same name but clearly two people).
    blocks = {}
    for r in recs:
        if not r["last"] or not r["first"]:
            continue
        key = (r["last"], r["first"][0])
        blocks.setdefault(key, []).append(r)

    for key, group in blocks.items():
        m = len(group)
        for a in range(m):
            for b in range(a + 1, m):
                r1, r2 = group[a], group[b]
                if uf.find(r1["idx"]) == uf.find(r2["idx"]):
                    continue                       # already linked (e.g. by SSN)
                if ssn_conflict(r1, r2) or dob_conflict(r1, r2):
                    continue                       # different SSN/DOB -> not same
                if name_match(r1, r2):
                    _link(uf, reasons, r1, r2, classify_name_link(r1, r2))


def classify_name_link(r1, r2) -> str:
    """Label which rule explains a name-based link (for the audit column)."""
    if r1["dob"] and r2["dob"] and r1["dob"] == r2["dob"]:
        base = "R4 same-name+DOB"
    elif (r1["ssn"] or r2["ssn"]) and not (r1["ssn"] and r2["ssn"]):
        base = "R2/R3 same-name+one-id"      # one has SSN/DOB, other blank
    else:
        base = "R4 same-name"
    if r1["mid"] != r2["mid"] and (r1["mid"] and r2["mid"]):
        base += " +R6 middle-initial"
    if r1["suf"] != r2["suf"]:
        base += " +suffix-diff"
    return base


# ------------------------------------------------------------
# 7) Unique ID generator  ->  3 letters + 4-digit number
#    AAA1000, AAA1001, ... AAA9999, AAB1000, ...
#    (same person gets the same Unique ID; NO rows are merged/removed)
# ------------------------------------------------------------
UID_FIRST_NUMBER = 1000        # 4-digit numbers start at 1000
UID_NUMBER_SPAN = 9000         # 1000..9999 inclusive


def make_uid(seq: int) -> str:
    """seq is 0-based order of the person. Numbers roll 1000->9999, then the
    3-letter prefix advances AAA->AAB->...->ZZZ (over 234 million IDs)."""
    letter_index = seq // UID_NUMBER_SPAN
    number = UID_FIRST_NUMBER + (seq % UID_NUMBER_SPAN)
    if letter_index >= 26 * 26 * 26:
        raise RuntimeError("Ran out of Unique IDs (max ZZZ9999).")
    a = letter_index // (26 * 26)
    b = (letter_index // 26) % 26
    c = letter_index % 26
    letters = chr(65 + a) + chr(65 + b) + chr(65 + c)
    return f"{letters}{number}"


# ------------------------------------------------------------
# 8) Main
# ------------------------------------------------------------
def main() -> None:
    print(f"Connecting to {SERVER} / {DATABASE} ...")
    with connect() as conn:
        # DISTINCT collapses exact duplicates first (30k -> ~10k), matching the
        # SELECT DISTINCT step you already run; the fuzzy rules below then
        # catch the NEAR-duplicates that DISTINCT cannot (Sajan vs Saj, etc.).
        sql = (
            f"SELECT DISTINCT {q(COL_FIRST)}, {q(COL_LAST)}, {q(COL_MIDDLE)}, "
            f"{q(COL_SUFFIX)}, {q(COL_DOB)}, {q(COL_SSN)} FROM {FULL_TABLE}"
        )
        print("Reading rows (SELECT DISTINCT) ...")
        df = pd.read_sql(sql, conn)
    print(f"  {len(df):,} rows read.")

    df = df.reset_index(drop=True)
    recs = build_records(df)
    uf = UnionFind(len(recs))
    reasons = [set() for _ in recs]

    print("Clustering ...")
    cluster(recs, uf, reasons)

    # Assign a Unique ID per cluster root. Same person -> same Unique ID.
    # NO rows are merged or removed; every input row is kept.
    root_to_uid = {}
    unique_id = [""] * len(recs)
    next_seq = 0
    for r in recs:                        # preserves original row order
        root = uf.find(r["idx"])
        if root not in root_to_uid:
            root_to_uid[root] = make_uid(next_seq)
            next_seq += 1
        unique_id[r["idx"]] = root_to_uid[root]

    df_out = df.copy()
    df_out.insert(0, "Unique ID", unique_id)
    df_out["Duplicate Count"] = df_out.groupby("Unique ID")["Unique ID"].transform("size")
    df_out["match_reason"] = [", ".join(sorted(s)) if s else "unique" for s in reasons]

    n_people = df_out["Unique ID"].nunique()
    n_dups = (df_out["Duplicate Count"] > 1).sum()
    print(f"  {len(df):,} rows kept  ->  {n_people:,} distinct Unique IDs "
          f"({n_dups:,} rows share an ID with at least one other row).")

    # Write results: ALL rows, with the new Unique ID column. No merged sheet.
    print(f"Writing {OUTPUT_XLSX} ...")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as xl:
        _write_sheet(xl, "Data with Unique ID",
                     df_out.sort_values(["Duplicate Count", "Unique ID"],
                                        ascending=[False, True]))
    print(f"Done -> {OUTPUT_XLSX}")
    print("Reminder: save the output to the secured Global Insider folder, "
          "not the desktop. It contains SSN/DOB -- handle via authorized systems only.")


def _write_sheet(xl, sheet, df):
    if len(df) > EXCEL_MAX_ROWS:
        csv_name = OUTPUT_XLSX.replace(".xlsx", f" {sheet}.csv")
        df.to_csv(csv_name, index=False, encoding="utf-8-sig")
        print(f"  {sheet}: {len(df):,} rows exceed Excel limit -> {csv_name}")
        pd.DataFrame({"note": [f"{sheet} exported to {csv_name} (too large)"]}
                     ).to_excel(xl, sheet_name=sheet, index=False)
    else:
        df.to_excel(xl, sheet_name=sheet, index=False)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
