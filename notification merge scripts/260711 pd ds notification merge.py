"""
260711 pd ds notification merge.py

Merge PII/PHI person records from an Excel export into one row per
confirmed person, for the notification report. Implements the rules
documented in "260711 re ds notification merge summary.md" (Base rule
+ Rules 1-9).

INPUT  : an Excel workbook with the columns listed in EXPECTED_COLS below.
OUTPUT : a new Excel workbook, ONE ROW PER CONFIRMED PERSON. Rows that
         match are combined; PII/PHI fields keep every distinct value
         seen, joined with "; ".

This script does not touch the input file. Save the output only to the
secured/authorized folder for this data (never a desktop) - it contains
SSN, DOB, and other PII/PHI.

Designed for large row counts (uses "blocking" - only compares rows
that already share an exact SSN, exact (Last Name, First Initial), or
exact PII value - instead of comparing every row to every other row).

Install once:
    pip install pandas openpyxl

Run:
    python "260711 pd ds notification merge.py"
"""

import sys
import re
import itertools
from collections import defaultdict

import pandas as pd

# ------------------------------------------------------------
# 1) CONFIG - edit these to match your workbook
# ------------------------------------------------------------
INPUT_XLSX  = "Data_01.xlsx"
INPUT_SHEET = 0
OUTPUT_XLSX = "260711 re ds notification merge output.xlsx"

COL_DOCID  = "DOCIDs"
COL_FIRST  = "First Name"
COL_LAST   = "Last Name"
COL_MIDDLE = "Middle Name"
COL_SUFFIX = "Suffix"
COL_DOB    = "Full Date of Birth (MM/DD/YYYY)"
COL_SSN    = "Social Security Number"

# PII fields used for Rule 3 (matching PII overrides a blank/"[Unknown]" name)
COL_DL      = "Driver's License Number"
COL_PASSPORT = "Passport Number"
COL_GOVID   = "Government-Issued ID Number"

# Every other column in the sheet - these get semicolon-merged as-is.
# Edit this list if your real headers differ.
OTHER_MERGE_COLS = [
    "Data Subject Type",
    "Birth Information",
    "Residential Address",
    "City",
    "State of Residence (if US)",
    "Province of Residence (if Canada)",
    "Zip Code",
    "Country of Residence",
    "Address Comments",
    "Email Address - Personal",
    "Phone Number",
    "Contact Information",
    COL_DL,
    "DL Issuing Country",
    "DL Issuing Province (if Canada)",
    "DL Issuing State (if US)",
    "Passport Country",
    COL_PASSPORT,
    "Government ID Issuing Country",
    "Government- Issued Identification",
    COL_GOVID,
    "Health Related Information",
    "Employee Identification Number",
    "Work-Related Information",
    "Family Information",
    "Financial Account Information",
    "Student-Related Information",
    "Demographic Information",
    "Biometric Data",
    "PI Notes",
    "Access Credentials (Non-Financial Account)",
]

EXPECTED_COLS = (
    [COL_DOCID, COL_FIRST, COL_LAST, COL_MIDDLE, COL_SUFFIX, COL_DOB, COL_SSN]
    + OTHER_MERGE_COLS
)

MERGE_SEP = "; "
EXCEL_MAX_ROWS = 1_048_576

# Placeholder name values treated as blank (never match/conflict on their own;
# a real name always supersedes these - Rule 3). Checked after stripping
# brackets/parens/periods, so "[Unknown]", "(unknown)", "N/A" all match.
NAME_PLACEHOLDERS = {
    "UNKNOWN", "UNK", "UNKN", "NA", "NONE", "NULL", "NIL",
    "TEST", "XXX", "XX", "X", "NMN", "NONAME", "NOTGIVEN", "NOTPROVIDED",
}

# Fake / junk SSNs that must never be used to match people.
SSN_PLACEHOLDERS = {
    "123456789", "987654321", "111223333", "123121234", "456789123",
    "078051120", "219099999", "457555462",
}


# ------------------------------------------------------------
# 2) Normalization helpers
# ------------------------------------------------------------
def norm_text(v) -> str:
    if v is None:
        return ""
    s = str(v).replace(" ", " ")
    s = re.sub(r"\s+", " ", s).strip().upper()
    return "" if s.lower() in ("nan", "none", "null") else s


def norm_name(v) -> str:
    """Upper/trim; placeholder values ('[Unknown]', 'N/A', ...) become ''
    so they never out-compete or conflict with a real name (Rule 3)."""
    s = norm_text(v)
    core = re.sub(r"[^A-Z0-9]", "", s)
    if not core or core in NAME_PLACEHOLDERS:
        return ""
    return s


def norm_suffix(v) -> str:
    return norm_text(v).replace(".", "")


def norm_ssn(v) -> str:
    """9 numeric digits only, not a known junk/placeholder value, or ''."""
    if v is None:
        return ""
    digits = re.sub(r"[^0-9]", "", str(v))
    if len(digits) != 9:
        return ""
    if len(set(digits)) == 1:               # 000000000, 111111111, ...
        return ""
    if digits in SSN_PLACEHOLDERS:
        return ""
    return digits


def norm_dob(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return ""
    ts = pd.to_datetime(s, errors="coerce")
    return "" if pd.isna(ts) else ts.strftime("%Y%m%d")


def norm_pii(v) -> str:
    return norm_text(v)


# ------------------------------------------------------------
# 3) Pairwise matching rules (Base + Rules 1-9 from the summary doc)
# ------------------------------------------------------------
def first_compat(a: str, b: str) -> bool:
    """Exact, or one is a prefix of the other (covers 'H'->'Harish' and
    'Did'->'Didar' with no minimum length - identity is corroborated by
    SSN/DOB in every rule that calls this, so a short prefix is safe)."""
    if not a or not b:
        return False
    if a == b:
        return True
    short, long_ = (a, b) if len(a) <= len(b) else (b, a)
    return long_.startswith(short)


def last_compat(a: str, b: str) -> bool:
    """Exact, or a >=3 char prefix/typo-tolerant match (covers 'Sin'/'Sing'
    -> 'Singh')."""
    if not a or not b:
        return False
    if a == b:
        return True
    short, long_ = (a, b) if len(a) <= len(b) else (b, a)
    return len(short) >= 3 and long_.startswith(short)


def middle_compat(a: str, b: str) -> bool:
    """Equal, either blank, or one is a prefix of the other (Rule 7)."""
    if a == b or not a or not b:
        return True
    short, long_ = (a, b) if len(a) <= len(b) else (b, a)
    return long_.startswith(short)


def name_match(r1, r2) -> bool:
    return (
        r1["last"] and r2["last"] and last_compat(r1["last"], r2["last"])
        and first_compat(r1["first"], r2["first"])
        and middle_compat(r1["mid"], r2["mid"])
    )


def suffix_conflict(r1, r2) -> bool:
    """Rule 8: a real, differing suffix blocks EVERY rule below, even the
    base SSN+DOB match. Rule 9 (blank vs. a real suffix) is NOT a conflict."""
    return bool(r1["suf"]) and bool(r2["suf"]) and r1["suf"] != r2["suf"]


def pii_match(r1, r2) -> bool:
    return (
        (r1["dl"] and r1["dl"] == r2["dl"])
        or (r1["passport"] and r1["passport"] == r2["passport"])
        or (r1["govid"] and r1["govid"] == r2["govid"])
    )


def is_match(r1, r2) -> bool:
    if suffix_conflict(r1, r2):
        return False

    ssn_same = bool(r1["ssn"]) and r1["ssn"] == r2["ssn"]
    dob_same = bool(r1["dob"]) and r1["dob"] == r2["dob"]

    if ssn_same and dob_same:                      # Base rule
        return True

    nm = name_match(r1, r2)
    if ssn_same and nm:                             # Rule 1 (DOB may differ)
        return True
    if nm and (ssn_same or dob_same):               # Rules 4-7
        return True

    # Rule 2: same exact name, complementary SSN/DOB (one has SSN only,
    # the other DOB only)
    if r1["first"] and r1["first"] == r2["first"] and r1["last"] and r1["last"] == r2["last"]:
        ssn_complementary = bool(r1["ssn"]) != bool(r2["ssn"])
        dob_complementary = bool(r1["dob"]) != bool(r2["dob"])
        if ssn_complementary and dob_complementary:
            return True

    # Rule 3: matching PII, one side's name is entirely blank/"[Unknown]"
    if pii_match(r1, r2):
        r1_blank = not r1["first"] and not r1["last"]
        r2_blank = not r2["first"] and not r2["last"]
        if r1_blank != r2_blank:
            return True

    return False


# ------------------------------------------------------------
# 4) Union-Find (disjoint set) for transitive clustering
# ------------------------------------------------------------
class UnionFind:
    def __init__(self, n):
        self.parent = list(range(n))

    def find(self, x):
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[max(ra, rb)] = min(ra, rb)


# ------------------------------------------------------------
# 5) Blocking - only test pairs that already share something exact, so we
#    never do a full N x N comparison. Candidate pairs come from three
#    cheap buckets; is_match() then applies the real rules within each.
# ------------------------------------------------------------
MAX_BUCKET_SIZE = 300   # safety valve: skip pairwise test inside a bucket


def build_records(df: pd.DataFrame):
    recs = []
    for i, row in df.iterrows():
        recs.append({
            "idx":      i,
            "first":    norm_name(row[COL_FIRST]),
            "last":     norm_name(row[COL_LAST]),
            "mid":      norm_name(row[COL_MIDDLE]),
            "suf":      norm_suffix(row[COL_SUFFIX]),
            "dob":      norm_dob(row[COL_DOB]),
            "ssn":      norm_ssn(row[COL_SSN]),
            "dl":       norm_pii(row[COL_DL]),
            "passport": norm_pii(row[COL_PASSPORT]),
            "govid":    norm_pii(row[COL_GOVID]),
        })
    return recs


def bucket_candidate_pairs(recs):
    buckets = defaultdict(list)
    for r in recs:
        if r["ssn"]:
            buckets[("ssn", r["ssn"])].append(r["idx"])
        if r["last"] and r["first"]:
            buckets[("name", r["last"], r["first"][0])].append(r["idx"])
        if r["dl"]:
            buckets[("dl", r["dl"])].append(r["idx"])
        if r["passport"]:
            buckets[("passport", r["passport"])].append(r["idx"])
        if r["govid"]:
            buckets[("govid", r["govid"])].append(r["idx"])

    pairs = set()
    skipped = 0
    for key, idxs in buckets.items():
        if len(idxs) < 2:
            continue
        if len(idxs) > MAX_BUCKET_SIZE:
            skipped += 1
            print(f"  WARNING: bucket {key[:1]}... has {len(idxs):,} rows "
                  f"(> {MAX_BUCKET_SIZE}) - skipping exhaustive compare, "
                  f"likely a junk/shared value. Review manually.")
            continue
        for a, b in itertools.combinations(sorted(idxs), 2):
            pairs.add((a, b))
    if skipped:
        print(f"  {skipped} oversized bucket(s) skipped - see warnings above.")
    return pairs


# ------------------------------------------------------------
# 6) Merge helpers for building the output
# ------------------------------------------------------------
def semicolon_merge(values) -> str:
    """Distinct, non-blank values joined with '; ', first-seen order,
    original casing preserved; dedup key is upper/trimmed."""
    seen = set()
    out = []
    for v in values:
        raw = "" if v is None else str(v).strip()
        key = norm_text(raw)
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(raw)
    return MERGE_SEP.join(out)


def fullest_value(values, norm_values) -> str:
    """Longest non-blank/non-placeholder value (Rules 4/5/6/7/9)."""
    best, best_norm_len = "", -1
    for raw, norm in zip(values, norm_values):
        if not norm:
            continue
        raw = str(raw).strip()
        if len(raw) > best_norm_len:
            best, best_norm_len = raw, len(raw)
    return best


def majority_dob(rows) -> str:
    """Rule 1: most frequent normalized DOB wins; '' if none present."""
    counts = defaultdict(int)
    raw_for = {}
    for r in rows:
        if r["dob"]:
            counts[r["dob"]] += 1
            raw_for.setdefault(r["dob"], r["dob_raw"])
    if not counts:
        return ""
    best = max(counts.items(), key=lambda kv: (kv[1], kv[0]))[0]
    return raw_for[best]


# ------------------------------------------------------------
# 7) Main
# ------------------------------------------------------------
def main() -> None:
    print(f"Reading {INPUT_XLSX} ...")
    df = pd.read_excel(INPUT_XLSX, sheet_name=INPUT_SHEET, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in EXPECTED_COLS if c not in df.columns]
    if missing:
        raise SystemExit(
            f"These expected columns were not found in {INPUT_XLSX}:\n"
            f"  {missing}\nColumns present:\n  {list(df.columns)}\n"
            "Fix the COL_*/OTHER_MERGE_COLS names in the CONFIG block."
        )
    print(f"  {len(df):,} rows read.")

    recs = build_records(df)
    for r in recs:
        r["dob_raw"] = df.loc[r["idx"], COL_DOB]

    # UnionFind is positional (0..n-1); map DataFrame index -> position
    uf = UnionFind(len(recs))
    pos_of = {r["idx"]: i for i, r in enumerate(recs)}
    recs_by_pos = recs

    def is_match_pos(a_idx, b_idx):
        return is_match(recs_by_pos[pos_of[a_idx]], recs_by_pos[pos_of[b_idx]])

    print("Clustering (blocked comparison) ...")
    pairs = bucket_candidate_pairs(recs)
    print(f"  {len(pairs):,} candidate pairs to test.")
    for a_idx, b_idx in pairs:
        if is_match_pos(a_idx, b_idx):
            uf.union(pos_of[a_idx], pos_of[b_idx])

    groups = defaultdict(list)
    for r in recs:
        groups[uf.find(pos_of[r["idx"]])].append(r["idx"])

    print(f"  {len(df):,} rows -> {len(groups):,} merged people "
          f"({len(df) - len(groups):,} rows collapsed by a match).")

    out_rows = []
    for root, idxs in groups.items():
        sub = df.loc[idxs]
        sub_recs = [r for r in recs if r["idx"] in idxs]

        row = {
            COL_DOCID: semicolon_merge(sub[COL_DOCID]),
            COL_FIRST: fullest_value(sub[COL_FIRST], [r["first"] for r in sub_recs]),
            COL_LAST:  fullest_value(sub[COL_LAST],  [r["last"] for r in sub_recs]),
            COL_MIDDLE: fullest_value(sub[COL_MIDDLE], [r["mid"] for r in sub_recs]),
            COL_SUFFIX: fullest_value(sub[COL_SUFFIX], [r["suf"] for r in sub_recs]),
            COL_DOB: majority_dob(sub_recs),
            COL_SSN: semicolon_merge(sub[COL_SSN]),
        }
        for c in OTHER_MERGE_COLS:
            row[c] = semicolon_merge(sub[c])
        row["Rows Merged"] = len(idxs)
        row["Names Differ"] = ";" in row[COL_FIRST] or ";" in row[COL_LAST]
        out_rows.append(row)

    df_out = pd.DataFrame(out_rows)
    df_out = df_out.sort_values(["Rows Merged"], ascending=False).reset_index(drop=True)

    n_multi = (df_out["Rows Merged"] > 1).sum()
    print(f"  {n_multi:,} merged groups combine 2+ original rows.")
    biggest = df_out["Rows Merged"].max()
    print(f"  Largest merged group: {biggest:,} rows.")
    if biggest > 50:
        print("  WARNING: a group >50 rows usually means a shared junk value "
              "(e.g. a fake SSN or a common PII placeholder). Inspect the top "
              "groups below before trusting the output.")
        print(df_out.sort_values("Rows Merged", ascending=False).head(10)
              [[COL_FIRST, COL_LAST, COL_SSN, "Rows Merged"]].to_string())

    print(f"Writing {OUTPUT_XLSX} ...")
    _write_sheet(OUTPUT_XLSX, "Merged Notification Data", df_out)
    print(f"Done -> {OUTPUT_XLSX}")
    print("Reminder: save the output only to the secured/authorized folder for "
          "this data - never a desktop or personal drive. It contains SSN, "
          "DOB, and other PII/PHI.")


def _write_sheet(path, sheet, df):
    if len(df) > EXCEL_MAX_ROWS:
        csv_name = path.replace(".xlsx", f" {sheet}.csv")
        df.to_csv(csv_name, index=False, encoding="utf-8-sig")
        print(f"  {sheet}: {len(df):,} rows exceed Excel limit -> {csv_name}")
        return
    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        df.to_excel(xl, sheet_name=sheet, index=False)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
