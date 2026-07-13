"""
260711 pd ds notification merge.py

Merge PII/PHI person records from an Excel export into one row per
confirmed person, for the notification report. Three rules, built step by
step:

  Rule 1 (SSN Exists): rows with the same (non-blank, non-junk) SSN
                        are merged.
  Rule 2 (Exact Name, DOB): rows with the same First Name + Last Name
                        (exact text match) AND the same DOB are merged,
                        as long as their SSNs don't conflict (blank on
                        either/both sides is fine).
  Rule 3 (Name-Only, No Identifiers): if EITHER row has no usable SSN and
                        no DOB at all (nothing to corroborate identity),
                        fall back to matching on Last Name (exact) +
                        compatible First Name (exact match, or blank on
                        either side). This lets a bare-bones record with
                        no SSN/DOB still attach to a fuller record by name
                        alone.

Either rule matching is enough to merge, and the match is transitive (if
A matches B and B matches C, all three end up in one merged row, even if
A and C don't directly match each other) - EXCEPT that a merge is refused
whenever it would combine 2+ DIFFERENT known SSNs into one group (this can
happen via a blank-SSN "bridge" row that matches two otherwise-unrelated
people). Rather than un-merging the whole cluster, only the specific union
that would cross real SSNs is refused - so e.g. 5 rows sharing one SSN and
3 rows sharing a different SSN, bridged by a blank-SSN row, still end up as
2 clean groups instead of 8 separate rows (see group_ssn/try_union() in
main()).

INPUT  : an Excel workbook with the columns listed in EXPECTED_COLS below.
OUTPUT : a new Excel workbook with one sheet, "Merged Notification Data":
         ONE ROW PER CONFIRMED PERSON.
         - First Name, Middle Name, Last Name, and SSN: the single fullest/
           most complete value among the merged rows (placeholder values
           like "[Unknown]" are never picked).
         - Suffix: the single fullest non-blank value.
         - DOB: the most frequent (majority) value among the merged rows.
         - Every OTHER column (DOCIDs, Driver's License, Gov ID, Employee ID,
           etc.): every distinct value seen, joined with "; ".
         - Address fields (Residential Address, City, State, Province, Zip,
           Country) are kept TOGETHER as one unit: the most common address
           stays in those columns as-is, and every OTHER distinct address
           goes into a new "Other Address" column as one combined string per
           address, semicolon-joined.

This script does not touch the input file. Save the output only to the
secured/authorized folder for this data (never a desktop) - it contains
SSN, DOB, and other PII/PHI.

Designed for large row counts (uses "blocking" - only compares rows that
already share an exact SSN, or an exact First+Last Name+DOB - instead of
comparing every row to every other row).

Install once:
    pip install pandas openpyxl

Run:
    python "260711 pd ds notification merge.py"
"""

import sys
import re
import itertools
import unicodedata
from collections import defaultdict
from multiprocessing import Pool

import pandas as pd

# Worker processes for the pairwise clustering step (the slowest phase on
# large files). Plain `threading` would NOT help here - is_match() is pure
# CPU-bound Python, and the GIL serializes threads so they'd just take turns
# instead of running in parallel. Separate processes actually parallelize it.
PARALLEL_WORKERS = 4
# Below this many candidate pairs, run single-process instead - spinning up
# worker processes has real overhead (~tens of ms each) that isn't worth it
# for a small/quick run.
PARALLEL_THRESHOLD = 20_000

# ------------------------------------------------------------
# Progress bar - one line, updated in place, no per-item explanation.
# ------------------------------------------------------------
_last_pct = {}


def progress(label: str, current: int, total: int, extra: str = "") -> None:
    """Prints '[label]  42% |########------------|  420,000/1,000,000  extra'
    on a single line, overwriting itself. Only redraws when the whole percent
    value changes, so it's cheap to call on every loop iteration."""
    pct = 100 if total <= 0 else min(100, current * 100 // total)
    if _last_pct.get(label) == pct and current != total:
        return
    _last_pct[label] = pct
    bar = "#" * (pct // 5) + "-" * (20 - pct // 5)
    tail = f"  {extra}" if extra else ""
    end = "\n" if current >= total else ""
    print(f"\r  [{label}] {pct:3d}% |{bar}| {current:,}/{total:,}{tail}   ",
          end=end, flush=True)


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

# Not used for matching (Steps 1-2 only use SSN and Name+DOB) - kept as
# plain semicolon-merged columns in OTHER_MERGE_COLS below.
COL_DL       = "Driver's License Number"
COL_PASSPORT = "Passport Number"
COL_GOVID    = "Government-Issued ID Number"
COL_EMPID    = "Employee Identification Number"

# Address fields are handled specially (see ADDRESS_COLS below), not as
# plain semicolon-merged columns - they need to stay together as one unit
# per address, not be shuffled independently per field.
COL_ADDR    = "Residential Address"
COL_CITY    = "City"
COL_STATE   = "State of Residence (if US)"
COL_PROVINCE = "Province of Residence (if Canada)"
COL_ZIP     = "Zip Code"
COL_COUNTRY = "Country of Residence"

# The full set of fields that make up "one address". The MAJORITY (most
# common) address among a merged person's rows is kept in these columns as-is;
# every OTHER distinct address goes into the "Other Address" output column
# as one combined string per address, semicolon-joined.
ADDRESS_COLS = [COL_ADDR, COL_CITY, COL_STATE, COL_PROVINCE, COL_ZIP, COL_COUNTRY]

# Every other column in the sheet - these get semicolon-merged as-is.
# Edit this list if your real headers differ.
OTHER_MERGE_COLS = [
    "Data Subject Type",
    "Birth Information",
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
    COL_EMPID,
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
    + ADDRESS_COLS + OTHER_MERGE_COLS
)

MERGE_SEP = "; "
EXCEL_MAX_ROWS = 1_048_576

# Placeholder name values treated as blank (never match/conflict on their own;
# a real name always supersedes these - Rule 3). Checked after stripping
# brackets/parens/periods, so "[Unknown]", "(unknown)", "N/A" all match.
NAME_PLACEHOLDERS = {
    "UNKNOWN", "UNK", "UNKN", "NA", "NONE", "NULL", "NIL",
    "XXX", "XX", "X", "NMN", "NONAME", "NOTGIVEN", "NOTPROVIDED",
}

# Fake / junk SSNs that must never be used to match people.
SSN_PLACEHOLDERS = {
    "123456789", "987654321", "111223333", "123121234", "456789123",
    "078051120", "219099999", "457555462",
}


# ------------------------------------------------------------
# 2) Normalization helpers
# ------------------------------------------------------------
# Zero-width space (U+200B), zero-width non-joiner (U+200C), zero-width
# joiner (U+200D), word joiner (U+2060), BOM (U+FEFF), non-breaking space
# (U+00A0), soft hyphen (U+00AD) - common invisible characters in Excel/CSV
# exports that make two values which LOOK identical fail an exact-string
# comparison. Built from chr() codes (not literal characters pasted into
# this file), so the source itself never contains an actual invisible char.
_INVISIBLE_CODEPOINTS = (0x200B, 0x200C, 0x200D, 0x2060, 0xFEFF, 0x00A0, 0x00AD)
_INVISIBLE_CHARS = re.compile("[" + "".join(chr(c) for c in _INVISIBLE_CODEPOINTS) + "]")


def norm_text(v) -> str:
    """Used as the DEDUP KEY everywhere (semicolon_merge, address_key, name
    matching). Normalizes unicode (NFKC - so visually-identical characters
    with different encodings compare equal), strips invisible characters
    (zero-width space/joiner, BOM, soft hyphen, non-breaking space) that are
    common in Excel/CSV exports and otherwise make two values that LOOK the
    same fail an exact-string dedup check, then collapses whitespace and
    upper-cases."""
    if v is None:
        return ""
    s = unicodedata.normalize("NFKC", str(v))
    s = _INVISIBLE_CHARS.sub(" ", s)
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


SSN_MIN_KNOWN_OVERLAP = 4   # min matching KNOWN digits to trust a masked SSN


def is_junk_ssn(d: str) -> bool:
    """True for a 9-digit (fully known) SSN that can't belong to a real
    person: all-same-digit, or a known fake/placeholder value."""
    return len(set(d)) == 1 or d in SSN_PLACEHOLDERS


def norm_ssn(v) -> str:
    """Return a 9-character pattern of digits and 'X' (X = redacted digit),
    or '' if unusable. Mask characters *, #, ? are treated as X, so
        123-45-6789 -> '123456789'
        123-45-XXXX -> '12345XXXX'
        123-45-6XXX -> '123456XXX'
    Rejected (-> ''): not 9 characters, a fully-known junk SSN (all-same-
    digit / known placeholder), or a masked SSN with fewer than
    SSN_MIN_KNOWN_OVERLAP known digits (too little information to trust)."""
    if v is None:
        return ""
    s = str(v).upper().replace("*", "X").replace("#", "X").replace("?", "X")
    kept = re.sub(r"[^0-9X]", "", s)
    if len(kept) != 9:
        return ""
    if "X" not in kept:                                    # fully known
        return "" if is_junk_ssn(kept) else kept
    known = sum(c != "X" for c in kept)                     # masked
    return kept if known >= SSN_MIN_KNOWN_OVERLAP else ""


def ssn_cmp(a: str, b: str) -> str:
    """Compare two 9-char SSN patterns, 'X' = wildcard. Returns:
        'diff'    - a known digit disagrees (definitely different SSNs)
        'same'    - known digits agree on >= SSN_MIN_KNOWN_OVERLAP positions
        'unknown' - compatible but not enough overlap to be sure"""
    overlap = 0
    for ca, cb in zip(a, b):
        if ca != "X" and cb != "X":
            if ca != cb:
                return "diff"
            overlap += 1
    return "same" if overlap >= SSN_MIN_KNOWN_OVERLAP else "unknown"


def norm_dob(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return ""
    ts = pd.to_datetime(s, errors="coerce")
    return "" if pd.isna(ts) else ts.strftime("%Y%m%d")


# ------------------------------------------------------------
# 3) Record type - __slots__ for fast attribute access at scale
#    (this loop runs millions of times, so dict-key lookups add up)
# ------------------------------------------------------------
class Rec:
    __slots__ = ("idx", "first", "last", "mid", "dob", "dob_raw", "ssn")

    def __init__(self, idx):
        self.idx = idx


def build_records(df: pd.DataFrame):
    """df MUST already have a 0..n-1 RangeIndex (see main()) - idx doubles
    as the record's position, so no separate index->position map is needed.
    Uses positional numpy-array access (df.values) rather than itertuples(),
    since itertuples() mangles column names with spaces/punctuation."""
    recs = []
    col_pos = {c: p for p, c in enumerate(df.columns)}
    values = df.values  # numpy object array, fast positional access
    fi, la, mi, do, ss = (col_pos[COL_FIRST], col_pos[COL_LAST], col_pos[COL_MIDDLE],
                          col_pos[COL_DOB], col_pos[COL_SSN])
    for i in range(len(df)):
        row = values[i]
        r = Rec(i)
        r.first = norm_name(row[fi])
        r.last = norm_name(row[la])
        r.mid = norm_name(row[mi])
        r.dob_raw = row[do]
        r.dob = norm_dob(row[do])
        r.ssn = norm_ssn(row[ss])
        recs.append(r)
    return recs


# ------------------------------------------------------------
# 4) Pairwise matching rules - built STEP BY STEP.
#    Add each newly confirmed rule here as its own small function, then
#    call it from is_match() below.
# ------------------------------------------------------------
def _name_compat(a: str, b: str) -> bool:
    """Loose compatibility check used ONLY by identity_conflict() below:
    blank on either side, exact match, or one is a text prefix of the other
    (covers initials/nicknames like 'D'/'Didar', 'Jon'/'Jonathan'). This is
    intentionally more forgiving than Rule 2's exact-only matching, since
    here it's just deciding whether a name difference LOOKS like ordinary
    spelling variation vs. a genuinely different name."""
    if not a or not b:
        return True
    if a == b:
        return True
    short, long_ = (a, b) if len(a) <= len(b) else (b, a)
    return long_.startswith(short)


def identity_conflict(r1: Rec, r2: Rec) -> bool:
    """True when EVERY other available signal disagrees: both First and
    Last Name are present on both sides and genuinely unrelated (not exact,
    not an initial/prefix match - e.g. 'James Ebersole' vs 'Jose Gallegos'),
    AND both DOBs are present and different. This means a shared SSN is
    more likely itself wrong/reused/fake than proof of the same identity, so
    it blocks Step 1 (SSN Exists) for this specific pair. If DOB is blank on
    either side, or the names are merely spelling variants, this does NOT
    fire - Rule 1 still merges purely on a shared SSN as usual."""
    if not (r1.first and r1.last and r2.first and r2.last):
        return False
    if _name_compat(r1.first, r2.first) and _name_compat(r1.last, r2.last):
        return False
    return bool(r1.dob) and bool(r2.dob) and r1.dob != r2.dob


def ssn_exists_match(r1: Rec, r2: Rec) -> bool:
    """Step 1 - 'SSN Exists': both rows have a usable (non-blank, non-junk)
    SSN and it's the same value. Name is NOT considered here at all, UNLESS
    identity_conflict() also finds the DOB disagreeing on top of a
    genuinely unrelated name - in that case the shared SSN alone isn't
    trusted (see identity_conflict())."""
    if identity_conflict(r1, r2):
        return False
    return bool(r1.ssn) and bool(r2.ssn) and r1.ssn == r2.ssn


def ssn_conflict(r1: Rec, r2: Rec) -> bool:
    """True when BOTH rows have a usable, known SSN and it genuinely
    disagrees. Used to block Step 2 (Name+DOB) from merging two different
    real people who happen to share a name and DOB - without this guard,
    Union-Find's transitive clustering can chain unrelated SSNs together
    through a shared-name-and-DOB row (e.g. A<->B via SSN, B<->C via
    Name+DOB even though A and C have different SSNs), contaminating a
    single merged record with several different people's SSNs."""
    return bool(r1.ssn) and bool(r2.ssn) and r1.ssn != r2.ssn


def exact_name_dob_match(r1: Rec, r2: Rec) -> bool:
    """Step 2 - 'Exact Name, DOB': both rows have the same First Name AND
    the same Last Name (exact text match, no typo/prefix tolerance) AND the
    same DOB, AND their SSNs don't conflict (blank on either/both sides is
    fine, but two different known SSNs block the merge). Middle Name/Suffix
    are NOT part of the match - they just get semicolon-merged like every
    other column when this rule fires."""
    if ssn_conflict(r1, r2):
        return False
    return (
        bool(r1.first) and bool(r1.last) and bool(r1.dob)
        and r1.first == r2.first and r1.last == r2.last and r1.dob == r2.dob
    )


def has_no_identifiers(r: Rec) -> bool:
    """True when a row has NEITHER a usable SSN NOR a DOB - nothing at all
    to corroborate identity with."""
    return not r.ssn and not r.dob


def no_identifier_name_match(r1: Rec, r2: Rec) -> bool:
    """Step 3 - 'Name-Only, No Identifiers': only applies when at least one
    side has no SSN and no DOB at all (see has_no_identifiers()) - otherwise
    Rule 2 already covers rows that both have a DOB. Matches on Last Name
    (exact) plus a compatible First Name (exact match, or blank on either
    side - a blank First Name is never treated as a conflict). This is
    intentionally weaker evidence than Rules 1/2 (name alone, no SSN/DOB
    backing it up), so it's a last-resort fallback for bare-bones records.
    The group-level SSN-consistency check in main() still applies - if this
    rule ever bridges two records that turn out to have different real
    SSNs, the whole group is un-merged again."""
    if not (has_no_identifiers(r1) or has_no_identifiers(r2)):
        return False
    if not r1.last or not r2.last or r1.last != r2.last:
        return False
    if r1.first and r2.first and r1.first != r2.first:
        return False
    return True


def is_match(r1: Rec, r2: Rec) -> bool:
    return (
        ssn_exists_match(r1, r2)
        or exact_name_dob_match(r1, r2)
        or no_identifier_name_match(r1, r2)
    )


# ------------------------------------------------------------
# 4b) Multiprocessing workers for the pairwise clustering step. Must be
#     module-level functions (not closures) so they can be pickled and sent
#     to worker processes. _worker_recs is set once per worker via the Pool
#     initializer, instead of re-sending the full `recs` list with every
#     chunk.
# ------------------------------------------------------------
_worker_recs = None


def _init_worker(recs):
    global _worker_recs
    _worker_recs = recs


def _match_chunk(chunk):
    """Runs in a worker process: tests every pair in this chunk, returns only
    the ones that matched."""
    return [(a, b) for a, b in chunk if is_match(_worker_recs[a], _worker_recs[b])]


def _chunk_pairs(pairs_list, target_chunks=200, min_chunk_size=200):
    """Splits pairs_list into roughly target_chunks pieces (never smaller
    than min_chunk_size), so progress can update ~target_chunks times
    regardless of how many pairs there are in total."""
    n = len(pairs_list)
    if n == 0:
        return []
    chunk_size = max(min_chunk_size, -(-n // target_chunks))   # ceil division
    return [pairs_list[i:i + chunk_size] for i in range(0, n, chunk_size)]


# ------------------------------------------------------------
# 5) Union-Find (disjoint set) for transitive clustering
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
# 6) Blocking - only test pairs that already share something exact, so we
#    never do a full N x N comparison. Candidate pairs come from one bucket
#    per active rule; is_match() then applies the real rules within each.
# ------------------------------------------------------------
MAX_BUCKET_SIZE = 300   # safety valve: skip pairwise test inside a bucket


def bucket_candidate_pairs(recs):
    buckets = defaultdict(list)
    # Rule 3's bucket needs special handling: bucketing by Last Name ALONE
    # (as a naive approach would) makes the bucket as big as every row that
    # happens to share a common surname - easily 300+ in a large file, which
    # then gets skipped outright by the MAX_BUCKET_SIZE safety valve, so even
    # obvious exact duplicates (same First+Last, e.g. 80 copies of "Harish
    # Kuntz") never get compared at all. Instead: bucket named rows by the
    # FULL (Last, First) pair - keeps buckets small, scoped to one specific
    # name, however common the surname is. Rows with a BLANK First Name
    # (compatible with any First Name under Rule 3) get their own smaller
    # "blank" bucket per Last Name, which also picks up one representative
    # row from each distinct (Last, First) group under that surname - enough
    # to bridge a blank-First row to a whole named group via Union-Find,
    # without needing a full comparison against every member of it.
    lastfirst_reps = {}
    for r in recs:
        if r.ssn:                                    # Rule 1: SSN Exists
            buckets[("ssn", r.ssn)].append(r.idx)
        if r.first and r.last and r.dob:              # Rule 2: Exact Name, DOB
            buckets[("namedob", r.first, r.last, r.dob)].append(r.idx)
        if r.last:                                    # Rule 3: Name-Only, No Identifiers
            if r.first:
                key = ("lastfirst", r.last, r.first)
                lastfirst_reps.setdefault(key, r.idx)
                buckets[key].append(r.idx)
            else:
                buckets[("lastblank", r.last)].append(r.idx)

    for (kind, last, first), rep_idx in lastfirst_reps.items():
        blank_key = ("lastblank", last)
        if blank_key in buckets:
            buckets[blank_key].append(rep_idx)

    pairs = set()
    skipped_buckets = 0
    skipped_rows = 0
    total = len(buckets)
    for n, (key, idxs) in enumerate(buckets.items(), 1):
        progress("Bucketing", n, total, extra=f"skipped={skipped_buckets}")
        if len(idxs) < 2:
            continue
        if len(idxs) > MAX_BUCKET_SIZE:
            skipped_buckets += 1
            skipped_rows += len(idxs)
            continue
        for a, b in itertools.combinations(sorted(idxs), 2):
            pairs.add((a, b))
    if skipped_buckets:
        print(f"  Skipped {skipped_buckets:,} oversized bucket(s) "
              f"({skipped_rows:,} rows, likely a shared junk value) - review manually.")
    return pairs


# ------------------------------------------------------------
# 7) Merge helpers for building the output
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


def fullest_value(raw_values, norm_values) -> str:
    """Longest raw value whose norm form is non-blank (placeholders/blanks
    are skipped); '' if every value is blank/placeholder. Used for First/
    Middle/Last Name, Suffix, and SSN - these keep ONE final value per
    merged person, not every variant."""
    best, best_len = "", -1
    for raw, norm in zip(raw_values, norm_values):
        if not norm:
            continue
        raw = "" if raw is None else str(raw).strip()
        if len(raw) > best_len:
            best, best_len = raw, len(raw)
    return best


def majority_dob(sub_recs) -> str:
    """Most frequent normalized DOB among the group wins (Rule 1 doesn't
    require DOB to match, so a merged group can legitimately contain more
    than one DOB value); '' if none present."""
    counts = defaultdict(int)
    raw_for = {}
    for r in sub_recs:
        if r.dob:
            counts[r.dob] += 1
            raw_for.setdefault(r.dob, r.dob_raw)
    if not counts:
        return ""
    best = max(counts.items(), key=lambda kv: (kv[1], kv[0]))[0]
    return raw_for[best]


def has_variation(raw_values) -> bool:
    """True if the group had 2+ distinct real values for this field, even
    though only one (the fullest) was kept in the output - used for the
    'Names Differ' review flag."""
    return len({norm_text(v) for v in raw_values if norm_text(v)}) > 1


def address_key(values) -> tuple:
    """Normalized tuple used to tell whether two rows have the SAME address
    (all fields blank-insensitive) - used to find the majority address."""
    return tuple(norm_text(v) for v in values)


def format_full_address(values) -> str:
    """One address, all its parts combined into a single readable string,
    e.g. '123 ABC Ln, Springfield, IL, 62701, USA'. Blank parts are skipped."""
    parts = []
    for v in values:
        s = "" if v is None else str(v).strip()
        if s and s.lower() not in ("nan", "none", "null"):
            parts.append(s)
    return ", ".join(parts)


def split_addresses(df, group_idxs):
    """Returns (majority_values, other_address_string) for one merged group.
    majority_values: the ADDRESS_COLS values (as originally entered) for the
    address that appears most often among this group's rows - these go into
    the normal Residential Address/City/State/Zip/Country columns unchanged.
    other_address_string: every OTHER distinct address in this group,
    combined into one string per address and semicolon-joined, for the
    'Other Address' column. A row with no address at all doesn't count as
    a "real" address unless it's the only kind of address in the group."""
    counts = defaultdict(int)
    first_values = {}
    first_seen_order = []
    for idx in group_idxs:
        values = tuple(df.at[idx, c] for c in ADDRESS_COLS)
        key = address_key(values)
        counts[key] += 1
        if key not in first_values:
            first_values[key] = values
            first_seen_order.append(key)

    non_blank_keys = [k for k in counts if any(k)]
    candidates = non_blank_keys or list(counts.keys())
    order_rank = {k: i for i, k in enumerate(first_seen_order)}
    majority_key = max(candidates, key=lambda k: (counts[k], -order_rank[k]))

    other_strings = []
    for key in first_seen_order:
        if key == majority_key or not any(key):
            continue
        other_strings.append(format_full_address(first_values[key]))

    return first_values[majority_key], MERGE_SEP.join(other_strings)


# ------------------------------------------------------------
# 8) Main
# ------------------------------------------------------------
def main() -> None:
    print(f"Reading {INPUT_XLSX} ...")
    df = pd.read_excel(INPUT_XLSX, sheet_name=INPUT_SHEET, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.reset_index(drop=True)   # guarantees row position == record idx

    missing = [c for c in EXPECTED_COLS if c not in df.columns]
    if missing:
        raise SystemExit(
            f"These expected columns were not found in {INPUT_XLSX}:\n"
            f"  {missing}\nColumns present:\n  {list(df.columns)}\n"
            "Fix the COL_*/OTHER_MERGE_COLS names in the CONFIG block."
        )
    print(f"  {len(df):,} rows read.")

    recs = build_records(df)   # recs[i].idx == i == df row position

    print("Clustering (blocked comparison) ...")
    pairs = bucket_candidate_pairs(recs)
    print(f"  {len(pairs):,} candidate pairs to test.")

    uf = UnionFind(len(recs))
    # group_ssn[root] = every DISTINCT known SSN currently inside that root's
    # merged group (not just the two rows being compared). Guards against
    # transitive contamination: a Step 2/3 match can bridge through a row
    # with a BLANK SSN (blank never conflicts pairwise), which would
    # otherwise let two groups with different real SSNs merge into one via
    # that bridge row even though neither pairwise SSN conflicts directly.
    # Refusing the union as soon as it would combine 2 different known SSNs
    # - rather than merging everything first and un-merging afterward -
    # means the SAFE parts of a would-be conflicting cluster still merge
    # normally: e.g. 5 rows sharing SSN=111 and 3 rows sharing SSN=999,
    # bridged by one blank-SSN row, end up as 2 clean groups (not 8 separate
    # rows) - the bridge row itself lands in whichever group it's processed
    # against first (order can vary; which side "wins" isn't predictable,
    # but two different real SSNs never end up in the same group either way).
    group_ssn = [({r.ssn} if r.ssn else set()) for r in recs]
    # group_first[root] = every DISTINCT known First Name currently inside
    # that root's group - the same kind of guard, but for Step 3 (Name-Only,
    # No Identifiers). A blank-First "bridge" row is compatible with ANY
    # First Name under a shared Last Name, so without this guard it could
    # transitively glue together many genuinely different real people who
    # just happen to share a common surname (e.g. 250 distinct "Kuntz"es,
    # each with a different first name, all bridged into one group by a
    # single blank-First "Kuntz" row) - and since none of them may have an
    # SSN at all, group_ssn alone wouldn't catch it. Only enforced when the
    # match ISN'T also justified by Step 1/2 (ssn_exists_match /
    # exact_name_dob_match), since those are strong enough evidence to
    # override a name difference on their own (e.g. same SSN, different
    # spelled name, is supposed to merge).
    group_first = [({r.first} if r.first else set()) for r in recs]
    refused_ssn = 0
    refused_name = 0

    def try_union(a_idx, b_idx):
        nonlocal refused_ssn, refused_name
        ra, rb = uf.find(a_idx), uf.find(b_idx)
        if ra == rb:
            return
        sa, sb = group_ssn[ra], group_ssn[rb]
        if sa and sb and sa.isdisjoint(sb):
            refused_ssn += 1
            return   # would combine two different real SSNs - refused
        r1, r2 = recs[a_idx], recs[b_idx]
        if not (ssn_exists_match(r1, r2) or exact_name_dob_match(r1, r2)):
            fa, fb = group_first[ra], group_first[rb]
            if fa and fb and fa.isdisjoint(fb):
                refused_name += 1
                return   # weak Rule 3 evidence only - would combine two different real First Names
        uf.union(a_idx, b_idx)
        merged_root = min(ra, rb)
        group_ssn[merged_root] = sa | sb
        group_first[merged_root] = group_first[ra] | group_first[rb]

    pairs_list = list(pairs)
    total_pairs = len(pairs_list)

    if total_pairs < PARALLEL_THRESHOLD:
        # Small run - not worth the process-startup overhead, just test
        # in-process.
        for tested, (a_idx, b_idx) in enumerate(pairs_list, 1):
            if is_match(recs[a_idx], recs[b_idx]):
                try_union(a_idx, b_idx)
            progress("Clustering", tested, total_pairs)
    else:
        # Large run - spread the pairwise is_match() tests across
        # PARALLEL_WORKERS separate processes (real parallelism; a thread
        # pool would not help here, see PARALLEL_WORKERS comment above).
        # The actual union calls stay single-process afterward, since
        # Union-Find (and group_ssn) is shared, mutable state that can't be
        # split safely.
        chunks = _chunk_pairs(pairs_list)
        done = 0
        with Pool(processes=PARALLEL_WORKERS, initializer=_init_worker, initargs=(recs,)) as pool:
            for matched in pool.imap_unordered(_match_chunk, chunks):
                for a_idx, b_idx in matched:
                    try_union(a_idx, b_idx)
                done += 1
                progress("Clustering", done, len(chunks))

    groups = defaultdict(list)
    for r in recs:
        groups[uf.find(r.idx)].append(r.idx)
    groups = list(groups.values())

    print(f"  {len(df):,} rows -> {len(groups):,} merged people "
          f"({len(df) - len(groups):,} rows collapsed by a match).")
    if refused_ssn:
        print(f"  {refused_ssn:,} candidate merge(s) were refused - would have "
              f"combined 2+ different real SSNs into one group.")
    if refused_name:
        print(f"  {refused_name:,} candidate merge(s) were refused - would have "
              f"combined 2+ different real First Names into one group via the "
              f"Name-Only fallback rule.")

    print("Building merged output ...")
    SEMICOLON_COLS = [COL_DOCID] + OTHER_MERGE_COLS
    total_groups = len(groups)
    out_rows = []
    for n, group_idxs in enumerate(groups, 1):
        progress("Building output", n, total_groups)
        sub = df.iloc[group_idxs]           # O(group size), not O(n)
        sub_recs = [recs[i] for i in group_idxs]

        row = {c: semicolon_merge(sub[c]) for c in SEMICOLON_COLS}
        row[COL_FIRST] = fullest_value(sub[COL_FIRST], [r.first for r in sub_recs])
        row[COL_LAST] = fullest_value(sub[COL_LAST], [r.last for r in sub_recs])
        row[COL_MIDDLE] = fullest_value(sub[COL_MIDDLE], [r.mid for r in sub_recs])
        row[COL_SUFFIX] = fullest_value(sub[COL_SUFFIX], [norm_text(v) for v in sub[COL_SUFFIX]])
        row[COL_SSN] = fullest_value(sub[COL_SSN], [r.ssn for r in sub_recs])
        row[COL_DOB] = majority_dob(sub_recs)

        majority_addr_values, other_address = split_addresses(df, group_idxs)
        for c, v in zip(ADDRESS_COLS, majority_addr_values):
            row[c] = v
        row["Other Address"] = other_address

        row["Rows Merged"] = len(group_idxs)
        row["Names Differ"] = has_variation(sub[COL_FIRST]) or has_variation(sub[COL_LAST])
        out_rows.append(row)

    df_out = pd.DataFrame(out_rows)
    df_out = df_out.sort_values(["Rows Merged"], ascending=False).reset_index(drop=True)

    n_multi = (df_out["Rows Merged"] > 1).sum()
    print(f"  {n_multi:,} merged groups combine 2+ original rows.")
    biggest = df_out["Rows Merged"].max()
    print(f"  Largest merged group: {biggest:,} rows.")
    if biggest > 50:
        print("  WARNING: a group >50 rows usually means a shared junk value "
              "(e.g. a fake SSN). Inspect the top groups below before "
              "trusting the output.")
        print(df_out.sort_values("Rows Merged", ascending=False).head(10)
              [[COL_FIRST, COL_LAST, COL_SSN, "Rows Merged"]].to_string())

    print(f"Writing {OUTPUT_XLSX} ...")
    _write_workbook(OUTPUT_XLSX, {
        "Merged Notification Data": df_out,
    })
    print(f"Done -> {OUTPUT_XLSX}")
    print("Reminder: save the output only to the secured/authorized folder for "
          "this data - never a desktop or personal drive. It contains SSN, "
          "DOB, and other PII/PHI.")


def _write_workbook(path, sheets: dict):
    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        for sheet, df in sheets.items():
            if len(df) > EXCEL_MAX_ROWS:
                csv_name = path.replace(".xlsx", f" {sheet}.csv")
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
