"""
============================================================
260710 pd ds merge data.py
Assign a person Unique ID to an Excel file (Data_01.xlsx).

INPUT  : an Excel workbook with these columns:
         First Name, Middle Name, Last Name, Suffix,
         Date of Birth, Social Security Number
OUTPUT : a new Excel workbook = ALL rows kept, with a new "Unique ID"
         column where the SAME PERSON gets the SAME id.  No rows are
         merged or deleted.

Unique ID format: 3 letters + 4-digit number  ->  AAA1000, AAA1001, ...

The script first drops EXACT duplicate rows (your SELECT DISTINCT step),
then assigns ids.  Rows that are the same person but not identical
(Sajan vs Saj, "H" vs "Harish", suffix present/absent) share one id via
the rules below.

Matching rules (a pair of rows is the same person if ANY hold):
    R1  Same SSN -- but only if the names do NOT conflict
        (Didar Singh vs Harish Singh on one SSN stay separate)
    R2  Same fuzzy name + SSN present on one, blank on other
    R3  Same fuzzy name + one has DOB, other has SSN
    R4  Same fuzzy name (First+Middle+Last+Suffix)
    R6  Middle initial vs full ("H" / "Harish"), rest of name equal
Guards: identical name but DIFFERENT SSN or DIFFERENT DOB -> NOT merged.

"Fuzzy name" = Last equal AND First equal-or-prefix (min 3 chars) AND
               Middle equal-or-blank-or-prefix AND Suffix equal-or-blank.
Clustering is TRANSITIVE (union-find).

The input file is never modified. Save the output to the secured Global
Insider folder (never the desktop) -- it contains SSN/DOB.

Install once:
    pip install pandas openpyxl

Run:
    python "260710 pd ds merge data.py"
============================================================
"""

import sys
import re
import pandas as pd

# ------------------------------------------------------------
# 1) CONFIG  - edit these to match your files
# ------------------------------------------------------------
INPUT_XLSX  = "Data_01.xlsx"     # input workbook (same folder as this script)
INPUT_SHEET = 0                  # sheet name or 0 for the first sheet
OUTPUT_XLSX = "260710 re ds data with unique id.xlsx"

# Column names exactly as they appear in the header row
COL_FIRST  = "First Name"
COL_LAST   = "Last Name"
COL_MIDDLE = "Middle Name"
COL_SUFFIX = "Suffix"
COL_DOB    = "Date of Birth"
COL_SSN    = "Social Security Number"

# Tuning knobs
FIRST_MIN_PREFIX = 3      # "Saj" (>=3) counts as a prefix of "Sajan"; "Jo" would not
EXCEL_MAX_ROWS = 1_048_576

# Placeholder name values that are NOT a real name. These are treated as blank,
# so "[Unknown]" rows never match each other by name -- they can only share a
# Unique ID via a real, equal SSN. Compared AFTER stripping brackets/punctuation
# and spaces, so "[Unknown]", "(unknown)", "N/A", "UN KNOWN" all match "UNKNOWN".
# Add your own placeholder cores here (letters/digits only, upper-case).
NAME_PLACEHOLDERS = {
    "UNKNOWN", "UNK", "UNKN", "NA", "NONE", "NULL", "NIL",
    "TEST", "XXX", "XX", "X", "NMN", "NONAME", "NOTGIVEN", "NOTPROVIDED",
}

# Fake / placeholder SSNs that must NEVER be used to merge people. These are
# structurally valid but are common test/default values entered when the real
# SSN is unknown. Add any others you see in your data (9 digits, no dashes).
SSN_PLACEHOLDERS = {
    "123456789", "987654321", "111223333", "123121234", "456789123",
    "078051120",   # famous Woolworth wallet SSN
    "219099999", "457555462",
}


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


def norm_name(v) -> str:
    """Like norm_text, but placeholder values ('[Unknown]', 'N/A', ...) become
    '' so they are never treated as a real, matchable name. The check strips
    everything except letters/digits first, so brackets and punctuation don't
    hide a placeholder."""
    s = norm_text(v)
    core = re.sub(r"[^A-Z0-9]", "", s)          # drop [] () - / . spaces etc.
    if not core or core in NAME_PLACEHOLDERS:
        return ""
    return s


SSN_MIN_KNOWN_OVERLAP = 4     # min matching KNOWN digits to accept a masked SSN


def norm_ssn(v) -> str:
    """Return a 9-character SSN pattern of digits and 'X' (X = redacted digit),
    or '' if unusable. Mask characters *, #, ? are treated as X, e.g.
        123-45-6789 -> '123456789'
        123-45-XXXX -> '12345XXXX'
        XXX-XX-6789 -> 'XXXXX6789'
    Rejected (returns ''): not 9 chars, all-X (no info), a fully-known junk SSN
    (all same digit / 000000000), or a masked SSN with fewer than
    SSN_MIN_KNOWN_OVERLAP known digits."""
    if v is None:
        return ""
    s = str(v).upper().replace("*", "X").replace("#", "X").replace("?", "X")
    kept = re.sub(r"[^0-9X]", "", s)
    if len(kept) != 9:
        return ""
    if "X" not in kept:                                   # fully known
        return "" if is_junk_ssn(kept) else kept
    known = sum(c != "X" for c in kept)                   # masked
    return kept if known >= SSN_MIN_KNOWN_OVERLAP else ""


def is_junk_ssn(d: str) -> bool:
    """True for a 9-digit SSN that cannot belong to a real, single person:
    invalid per SSA structure, all-same-digit, sequential, or a known fake."""
    if len(set(d)) == 1:                       # 000000000, 111111111, ...
        return True
    if d in SSN_PLACEHOLDERS:
        return True
    if d in ("123456789", "234567890", "345678901", "456789012", "567890123",
             "012345678", "987654321", "098765432"):
        return True
    area, group, serial = d[:3], d[3:5], d[5:]
    if area in ("000", "666") or area >= "900":  # SSA-invalid area numbers
        return True
    if group == "00" or serial == "0000":        # SSA-invalid group / serial
        return True
    return False


def ssn_cmp(a: str, b: str) -> str:
    """Compare two 9-char SSN patterns treating 'X' as a wildcard.
    Returns 'diff'    -> a known digit disagrees (definitely different people),
            'same'    -> known digits agree on >= SSN_MIN_KNOWN_OVERLAP positions,
            'unknown' -> compatible but too few overlapping known digits to be
                         sure (e.g. '12345XXXX' vs 'XXXXX6789')."""
    overlap = 0
    for ca, cb in zip(a, b):
        if ca != "X" and cb != "X":
            if ca != cb:
                return "diff"
            overlap += 1
    return "same" if overlap >= SSN_MIN_KNOWN_OVERLAP else "unknown"


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
    """True only when both SSNs are present and a KNOWN digit disagrees.
    Masked SSNs whose known digits agree ('123456789' vs '12345XXXX') do NOT
    conflict; non-overlapping masks are 'unknown', not a conflict."""
    return bool(r1["ssn"]) and bool(r2["ssn"]) and ssn_cmp(r1["ssn"], r2["ssn"]) == "diff"


def dob_conflict(r1, r2) -> bool:
    """True when both DOBs are present and different."""
    return bool(r1["dob"]) and bool(r2["dob"]) and r1["dob"] != r2["dob"]


def group_is_coherent(group) -> bool:
    """True if NO two rows in the group have conflicting names or DOBs.
    A shared SSN whose rows carry clearly different names (Blankenship vs
    Bradford) is unreliable -- often a fake/default SSN -- so we refuse to
    merge the whole group and let the safer name rules decide instead. This
    also stops a blank '[Unknown]' row from bridging two conflicting people."""
    n = len(group)
    for a in range(n):
        for b in range(a + 1, n):
            if name_conflict(group[a], group[b]) or dob_conflict(group[a], group[b]):
                return False
    return True


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
            "first": norm_name(row[COL_FIRST]),
            "last":  norm_name(row[COL_LAST]),
            "mid":   norm_name(row[COL_MIDDLE]),
            "suf":   norm_name(row[COL_SUFFIX]),
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
    # ---- R1: same SSN. A shared full SSN merges its rows ONLY if the whole
    #          group is coherent (no two rows have conflicting names/DOBs).
    #          A group with conflicting names means the SSN is unreliable
    #          (usually a fake/default value) -> we skip it entirely so a blank
    #          '[Unknown]' row cannot bridge two different people.
    ssn_recs = [r for r in recs if r["ssn"]]
    masked = [r for r in ssn_recs if "X" in r["ssn"]]
    skipped_ssns = 0

    by_ssn = {}
    for r in ssn_recs:
        if "X" not in r["ssn"]:
            by_ssn.setdefault(r["ssn"], []).append(r)
    for ssn, group in by_ssn.items():
        if len(group) < 2:
            continue
        if not group_is_coherent(group):
            skipped_ssns += 1
            continue                              # unreliable SSN -> don't merge
        root = group[0]
        for other in group[1:]:                   # coherent -> all one person
            _link(uf, reasons, root, other, "R1 same-SSN")

    # Masked SSNs (123-45-XXXX): a mask reveals only a few digits, so we require
    # a POSITIVE real-name match (both names present and fuzzy-equal). This
    # prevents a mask from bridging '[Unknown]' rows across different people.
    for mr in masked:
        if not (mr["first"] and mr["last"]):
            continue                              # blank name: too weak to trust a mask
        for r in ssn_recs:
            if r["idx"] == mr["idx"] or uf.find(r["idx"]) == uf.find(mr["idx"]):
                continue
            if not (r["first"] and r["last"]):
                continue
            if ssn_cmp(mr["ssn"], r["ssn"]) != "same":
                continue
            if dob_conflict(mr, r) or not name_match(mr, r):
                continue
            _link(uf, reasons, mr, r, "R1 same-SSN (masked)")

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
EXPECTED_COLS = [COL_FIRST, COL_MIDDLE, COL_LAST, COL_SUFFIX, COL_DOB, COL_SSN]


def main() -> None:
    print(f"Reading {INPUT_XLSX} ...")
    df = pd.read_excel(INPUT_XLSX, sheet_name=INPUT_SHEET, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]   # tidy header whitespace

    missing = [c for c in EXPECTED_COLS if c not in df.columns]
    if missing:
        raise SystemExit(
            f"These expected columns were not found in {INPUT_XLSX}:\n"
            f"  {missing}\nColumns present:\n  {list(df.columns)}\n"
            "Fix the COL_* names in the CONFIG block to match your header row."
        )
    print(f"  {len(df):,} rows read.")

    # Drop EXACT duplicate rows first (your SELECT DISTINCT step). The fuzzy
    # rules below then catch the NEAR-duplicates DISTINCT cannot (Sajan vs Saj).
    before = len(df)
    df = df.drop_duplicates(subset=EXPECTED_COLS).reset_index(drop=True)
    print(f"  {before - len(df):,} exact-duplicate rows removed -> {len(df):,} rows.")
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

    # ---- SANITY CHECK: a real person should be a handful of rows, not
    #      thousands. Print the biggest clusters so over-merging is obvious.
    biggest = df_out["Duplicate Count"].max()
    print(f"  Largest cluster: {biggest:,} rows.")
    if biggest > 50:
        print("  WARNING: a cluster >50 rows almost always means a shared "
              "junk/placeholder value. Inspect these Unique IDs:")
        top = (df_out[df_out["Duplicate Count"] > 50]
               .groupby("Unique ID")
               .agg(rows=("Unique ID", "size"),
                    sample_last=(COL_LAST, lambda s: s.head(3).tolist()),
                    sample_ssn=(COL_SSN, lambda s: s.head(3).tolist()))
               .sort_values("rows", ascending=False).head(10))
        print(top.to_string())
        print("  If a cluster mixes different real names, add its SSN to "
              "SSN_PLACEHOLDERS (or its name to NAME_PLACEHOLDERS) and re-run.")

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
