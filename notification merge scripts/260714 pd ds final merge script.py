"""
260714 pd ds final merge script.py

SECOND PASS: takes the OUTPUT of "260711 pd ds notification merge.py" (its
"Merged Notification Data" sheet, already merged via Rules 1-9) and applies
ONE more rule - Full Name Match - to catch the same person still split
across multiple already-merged rows because each one carries a DIFFERENT
single piece of corroborating info: one row has only a DOB, another only
an SSN, another only an Address, another has nothing at all - with no
shared token or exact Name+DOB pair for Rules 1-9 to have anchored on.

Rule (Full Name Match): First + Last Name must match exactly (both
present). Middle Name and Suffix must be matching-or-blank. EVERY other
field must be non-conflicting:
    - SSN: matching or blank on both sides.
    - DOB: the first-pass output can already hold MULTIPLE semicolon-
      joined DOB values per row (Rule 1 doesn't require DOB to match) -
      treated as a set; blank on either side, or the two sets share at
      least one value, is fine. Two non-blank sets with NO overlap at all
      is a conflict.
    - Address (Residential Address/City/State/Zip/Province): reuses the
      same tolerant comparison as the main script (ZIP+4 vs plain ZIP,
      and a bare street vs. the same street + apartment/unit suffix are
      NOT conflicts) - a real, differing value in any field blocks it.
    - Employee ID, Driver's License, Passport, Phone, Email: each cell can
      already hold multiple semicolon-joined tokens - blank on either
      side, or the two token sets share at least one value, is fine; two
      non-empty, completely disjoint sets is a conflict.

    Name alone is not enough to merge, even with nothing else conflicting:
    at least ONE of SSN, DOB, Address, Driver's License, Phone, or Email
    must ACTUALLY agree (both rows non-blank and matching/overlapping,
    not merely blank-vs-present). Two rows sharing only the name, with
    every other field blank on both sides, are left unmerged.

INPUT  : the OUTPUT workbook from the main merge script - specifically its
         "Merged Notification Data" sheet.
OUTPUT : a new workbook with:
         - "Merged Notification Data": further merged, one row per
           confirmed person. "Rows Merged" is the SUM of the constituent
           first-pass rows' own "Rows Merged" counts (i.e. still counts
           back to the original raw rows, not just this pass's rows).
         - "Full Name Match Review": every pair merged in THIS second
           pass, shown side-by-side (Full Name, DOCID/SSN/DOB/Address for
           each row), so it's easy to see which single field each row
           actually contributed.
         - "Junk SSN Review", "Large Group Review", "Skipped Bucket
           Review": carried through unchanged from the input workbook, if
           present, so nothing from the first pass is lost.

This script does not touch the input file. Save the output only to the
secured/authorized folder for this data (never a desktop) - it contains
SSN, DOB, and other PII/PHI.

Designed for large row counts (uses "blocking" - only compares rows that
already share an exact First+Last Name, instead of comparing every row to
every other row).

Install once:
    pip install pandas openpyxl

Run:
    python "260714 pd ds final merge script.py"
"""

import sys
import re
import itertools
from collections import defaultdict

import pandas as pd

# ------------------------------------------------------------
# Progress bar - one line, updated in place, no per-item explanation.
# ------------------------------------------------------------
_last_pct = {}


def progress(label: str, current: int, total: int, extra: str = "") -> None:
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
# 1) CONFIG - edit these to match the main script's output headers
# ------------------------------------------------------------
INPUT_XLSX  = "260711 re ds notification merge output.xlsx"
INPUT_SHEET = "Merged Notification Data"
OUTPUT_XLSX = "260714 re ds final merge output.xlsx"

COL_DOCID  = "DOCIDs"
COL_FIRST  = "First Name"
COL_LAST   = "Last Name"
COL_MIDDLE = "Middle Name"
COL_SUFFIX = "Suffix"
COL_DOB    = "Full Date of Birth (MM/DD/YYYY)"
COL_SSN    = "Social Security Number"
COL_EMPID    = "Employee Identification Number"
COL_DL       = "Driver's License Number"
COL_PASSPORT = "Passport Number"
COL_PHONE    = "Phone Number"
COL_EMAIL    = "Email Address - Personal"
COL_ADDR     = "Residential Address"
COL_CITY     = "City"
COL_STATE    = "State of Residence (if US)"
COL_PROVINCE = "Province of Residence (if Canada)"
COL_ZIP      = "Zip Code"
COL_COUNTRY  = "Country of Residence"
COL_ROWS_MERGED = "Rows Merged"

ADDRESS_COLS = [COL_ADDR, COL_CITY, COL_STATE, COL_PROVINCE, COL_ZIP, COL_COUNTRY]
TOKEN_ID_COLS = [COL_EMPID, COL_DL, COL_PASSPORT, COL_PHONE, COL_EMAIL]

MERGE_SEP = "; "
EXCEL_MAX_ROWS = 1_048_576
DOCID_CHUNK_SIZE = 20_000   # keep well under Excel's 32,767-char cell limit

MAX_BUCKET_SIZE = 300   # safety valve: skip pairwise test inside a bucket

# Sheets from the first pass to carry through unchanged, if present.
PASSTHROUGH_SHEETS = ["Junk SSN Review", "Large Group Review", "Skipped Bucket Review"]


# ------------------------------------------------------------
# 2) Normalization helpers (same rules as the main script, applied here to
#    the main script's OUTPUT text, which is raw/as-entered again - the
#    first pass's own internal normalization doesn't carry over)
# ------------------------------------------------------------
def norm_text(v) -> str:
    if v is None:
        return ""
    s = str(v).strip().upper()
    s = re.sub(r"\s+", " ", s)
    return "" if s.lower() in ("nan", "none", "null") else s


def norm_name(v) -> str:
    s = norm_text(v)
    return "" if not re.sub(r"[^A-Z0-9]", "", s) else s


def compatible(a: str, b: str) -> bool:
    """True if either side is blank, or both sides are equal."""
    return not a or not b or a == b


def parse_set(v) -> frozenset:
    """Splits an already-merged cell ('12345; 12346' or '12345;12346') into
    a set of individual normalized values."""
    if v is None:
        return frozenset()
    return frozenset(norm_text(p) for p in str(v).split(";") if norm_text(p))


def set_compatible(s1: frozenset, s2: frozenset) -> bool:
    """True if either side is empty, or the two sets share at least one
    value. Two non-empty, COMPLETELY DISJOINT sets is a conflict."""
    return not s1 or not s2 or not s1.isdisjoint(s2)


def zip5(v: str) -> str:
    """First 5 digits of a ZIP code - so '62701' and '62701-1234' compare
    as the same base ZIP instead of a conflict."""
    digits = re.sub(r"[^0-9]", "", v)
    return digits[:5] if len(digits) >= 5 else ""


_STREET_TOKEN_MAP = {
    "NORTH": "N", "N": "N", "SOUTH": "S", "S": "S", "EAST": "E", "E": "E",
    "WEST": "W", "W": "W", "NORTHEAST": "NE", "NE": "NE", "NORTHWEST": "NW", "NW": "NW",
    "SOUTHEAST": "SE", "SE": "SE", "SOUTHWEST": "SW", "SW": "SW",
    "STREET": "ST", "ST": "ST", "AVENUE": "AVE", "AVE": "AVE", "AV": "AVE",
    "BOULEVARD": "BLVD", "BLVD": "BLVD", "ROAD": "RD", "RD": "RD",
    "LANE": "LN", "LN": "LN", "DRIVE": "DR", "DR": "DR", "COURT": "CT", "CT": "CT",
    "CIRCLE": "CIR", "CIR": "CIR", "PLACE": "PL", "PL": "PL",
    "TERRACE": "TER", "TERR": "TER", "TER": "TER", "PARKWAY": "PKWY", "PKWY": "PKWY",
    "HIGHWAY": "HWY", "HWY": "HWY", "SQUARE": "SQ", "SQ": "SQ",
    "TRAIL": "TRL", "TRL": "TRL", "WAY": "WAY", "LOOP": "LOOP", "COVE": "CV", "CV": "CV",
    "POINT": "PT", "PT": "PT", "CROSSING": "XING", "XING": "XING",
    "PLAZA": "PLZ", "PLZ": "PLZ", "EXPRESSWAY": "EXPY", "EXPY": "EXPY",
    "FREEWAY": "FWY", "FWY": "FWY", "ROUTE": "RTE", "RTE": "RTE",
    "JUNCTION": "JCT", "JCT": "JCT", "MOUNT": "MT", "MT": "MT",
    "MOUNTAIN": "MTN", "MTN": "MTN",
    "APARTMENT": "APT", "APT": "APT", "SUITE": "STE", "STE": "STE",
    "BUILDING": "BLDG", "BLDG": "BLDG", "FLOOR": "FL", "FL": "FL", "UNIT": "UNIT",
}
_UNIT_DESIGNATORS = {"APT", "STE", "UNIT", "BLDG", "FL"}
_DIRECTIONAL_TOKENS = {"N", "S", "E", "W", "NE", "NW", "SE", "SW"}
_HOUSE_UNIT_RE = re.compile(r"^(\d+)([A-Z])$")


def norm_street(v) -> str:
    """Same canonicalization as the main script: directionals/street-type
    abbreviations unified, and a unit letter loosely attached to the house
    number ('203A ...', '203 A ...') pulled out to a 'UNIT <letter>'
    suffix, same as an explicit 'Apt A'."""
    s = norm_text(v)
    if not s:
        return ""
    tokens = s.split(" ")
    unit_letter = ""
    m = _HOUSE_UNIT_RE.match(tokens[0])
    if m:
        tokens[0] = m.group(1)
        unit_letter = m.group(2)
    elif (len(tokens) > 1 and tokens[0].isdigit()
          and len(tokens[1]) == 1 and tokens[1] not in _DIRECTIONAL_TOKENS):
        unit_letter = tokens.pop(1)
    out = [_STREET_TOKEN_MAP.get(tok.rstrip("."), tok.rstrip(".")) for tok in tokens]
    if unit_letter and not any(t in _UNIT_DESIGNATORS for t in out):
        out += ["UNIT", unit_letter]
    return " ".join(out)


def split_street_unit(s: str) -> tuple:
    tokens = s.split(" ") if s else []
    for i, tok in enumerate(tokens):
        if tok in _UNIT_DESIGNATORS:
            return " ".join(tokens[:i]), " ".join(tokens[i + 1:])
        if tok.startswith("#") and len(tok) > 1:
            return " ".join(tokens[:i]), " ".join([tok[1:]] + tokens[i + 1:])
    return s, ""


def street_compat(a: str, b: str) -> bool:
    if not a or not b or a == b:
        return True
    base_a, unit_a = split_street_unit(a)
    base_b, unit_b = split_street_unit(b)
    if not base_a or base_a != base_b:
        return False
    return not (unit_a and unit_b and unit_a != unit_b)


def address_conflict(r1, r2) -> bool:
    """True when Residential Address, City, State, Zip, or Province has a
    real, DIFFERING value on both sides (blank on either side is never a
    conflict). Zip uses zip5(), street uses street_compat()."""
    if not street_compat(r1.addr, r2.addr):
        return True
    fields1 = (r1.city, r1.state, zip5(r1.zip), r1.province)
    fields2 = (r2.city, r2.state, zip5(r2.zip), r2.province)
    return any(a and b and a != b for a, b in zip(fields1, fields2))


# ------------------------------------------------------------
# 3) Record type
# ------------------------------------------------------------
class Rec:
    __slots__ = ("idx", "first", "last", "mid", "suf", "ssn", "dob",
                 "addr", "city", "state", "zip", "province",
                 "empids", "dl_ids", "passport_ids", "phones", "emails",
                 "rows_merged")

    def __init__(self, idx):
        self.idx = idx


def build_records(df: pd.DataFrame):
    recs = []
    for i in range(len(df)):
        row = df.iloc[i]
        r = Rec(i)
        r.first = norm_name(row.get(COL_FIRST))
        r.last = norm_name(row.get(COL_LAST))
        r.mid = norm_name(row.get(COL_MIDDLE))
        r.suf = norm_text(row.get(COL_SUFFIX))
        r.ssn = norm_text(row.get(COL_SSN))
        r.dob = parse_set(row.get(COL_DOB))
        r.addr = norm_street(row.get(COL_ADDR))
        r.city = norm_text(row.get(COL_CITY))
        r.state = norm_text(row.get(COL_STATE))
        r.zip = norm_text(row.get(COL_ZIP))
        r.province = norm_text(row.get(COL_PROVINCE))
        r.empids = parse_set(row.get(COL_EMPID))
        r.dl_ids = parse_set(row.get(COL_DL))
        r.passport_ids = parse_set(row.get(COL_PASSPORT))
        r.phones = parse_set(row.get(COL_PHONE))
        r.emails = parse_set(row.get(COL_EMAIL))
        r.rows_merged = int(row.get(COL_ROWS_MERGED) or 1)
        recs.append(r)
    return recs


# ------------------------------------------------------------
# 4) Rule 10 - Full Name Match
# ------------------------------------------------------------
def full_name_match(r1: Rec, r2: Rec) -> bool:
    if not (r1.first and r1.last and r1.first == r2.first and r1.last == r2.last):
        return False
    if not (compatible(r1.mid, r2.mid) and compatible(r1.suf, r2.suf)):
        return False
    if not compatible(r1.ssn, r2.ssn):
        return False
    if not set_compatible(r1.dob, r2.dob):
        return False
    if address_conflict(r1, r2):
        return False
    if not (
        set_compatible(r1.empids, r2.empids)
        and set_compatible(r1.dl_ids, r2.dl_ids)
        and set_compatible(r1.passport_ids, r2.passport_ids)
        and set_compatible(r1.phones, r2.phones)
        and set_compatible(r1.emails, r2.emails)
    ):
        return False

    # Nothing conflicts, but that's not enough on its own - the name match
    # still needs at least one piece of PII to ACTUALLY agree (both sides
    # non-blank and matching/overlapping), not just be blank on one side.
    ssn_match = bool(r1.ssn) and bool(r2.ssn) and r1.ssn == r2.ssn
    dob_match = bool(r1.dob) and bool(r2.dob) and not r1.dob.isdisjoint(r2.dob)
    addr_match = bool(r1.addr) and bool(r2.addr) and street_compat(r1.addr, r2.addr)
    dl_match = bool(r1.dl_ids) and bool(r2.dl_ids) and not r1.dl_ids.isdisjoint(r2.dl_ids)
    phone_match = bool(r1.phones) and bool(r2.phones) and not r1.phones.isdisjoint(r2.phones)
    email_match = bool(r1.emails) and bool(r2.emails) and not r1.emails.isdisjoint(r2.emails)
    return ssn_match or dob_match or addr_match or dl_match or phone_match or email_match


# ------------------------------------------------------------
# 5) Union-Find
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
# 6) Blocking - the only anchor Rule 10 needs is exact First+Last, so one
#    bucket type is enough.
# ------------------------------------------------------------
def bucket_candidate_pairs(recs):
    buckets = defaultdict(list)
    for r in recs:
        if r.first and r.last:
            buckets[(r.last, r.first)].append(r.idx)

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
              f"({skipped_rows:,} rows, likely a very common full name) - review manually.")
    return pairs


# ------------------------------------------------------------
# 7) Merge helpers
# ------------------------------------------------------------
def semicolon_merge(values) -> str:
    """Distinct, non-blank values joined with '; ', splitting every cell on
    ';' first and deduping at the token level (matches the main script)."""
    seen = set()
    out = []
    for v in values:
        if v is None:
            continue
        for tok in str(v).split(";"):
            raw = tok.strip()
            key = norm_text(raw)
            if not key or key in seen:
                continue
            seen.add(key)
            out.append(raw)
    return MERGE_SEP.join(out)


def fullest_value(values) -> str:
    """Longest non-blank/non-null raw value."""
    best, best_len = "", -1
    for v in values:
        raw = "" if v is None else str(v).strip()
        if not raw or raw.lower() in ("nan", "none", "null"):
            continue
        if len(raw) > best_len:
            best, best_len = raw, len(raw)
    return best


def address_key(values) -> tuple:
    """Normalized tuple (blank-insensitive, ZIP+4/plain-ZIP and street-unit
    tolerant, same as the main script) used to tell whether two rows have
    the SAME address - used to find the majority address so City/State/Zip
    stay linked to the Street they actually belong with, instead of being
    flattened independently and losing that pairing."""
    def _key(col, v):
        if col == COL_ZIP:
            z = zip5(norm_text(v))
            return z if z else norm_text(v)
        if col == COL_ADDR:
            return norm_street(v)
        return norm_text(v)
    return tuple(_key(col, v) for col, v in zip(ADDRESS_COLS, values))


def format_full_address(values) -> str:
    parts = []
    for v in values:
        s = "" if v is None else str(v).strip()
        if s and s.lower() not in ("nan", "none", "null"):
            parts.append(s)
    return ", ".join(parts)


def split_addresses(df, group_idxs):
    """Same 'majority address wins the normal columns, everything else
    goes to Other Address' logic as the main script - see that script's
    split_addresses() for the full rationale."""
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


def split_docid_chunks(docid_str, max_chars=DOCID_CHUNK_SIZE):
    if len(docid_str) <= max_chars:
        return [docid_str]
    parts = docid_str.split(MERGE_SEP)
    chunks = []
    current = []
    current_len = 0
    for part in parts:
        added_len = len(part) + (len(MERGE_SEP) if current else 0)
        if current and current_len + added_len > max_chars:
            chunks.append(MERGE_SEP.join(current))
            current = [part]
            current_len = len(part)
        else:
            current.append(part)
            current_len += added_len
    if current:
        chunks.append(MERGE_SEP.join(current))
    return chunks


# ------------------------------------------------------------
# 8) Main
# ------------------------------------------------------------
def main() -> None:
    print(f"Reading {INPUT_XLSX} (sheet '{INPUT_SHEET}') ...")
    df = pd.read_excel(INPUT_XLSX, sheet_name=INPUT_SHEET, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.reset_index(drop=True)

    missing = [c for c in [COL_DOCID, COL_FIRST, COL_LAST] if c not in df.columns]
    if missing:
        raise SystemExit(
            f"These expected columns were not found in {INPUT_XLSX}:\n"
            f"  {missing}\nColumns present:\n  {list(df.columns)}\n"
            "Fix the COL_* names in the CONFIG block."
        )
    print(f"  {len(df):,} rows read.")

    # DOCID may already be split across overflow columns from the first
    # pass ('DOCIDs', 'DOCIDs 2', 'DOCIDs 3', ...) - collect all of them.
    docid_cols = [c for c in df.columns if c == COL_DOCID or c.startswith(f"{COL_DOCID} ")]

    recs = build_records(df)

    print("Clustering (blocked comparison) ...")
    pairs = bucket_candidate_pairs(recs)
    print(f"  {len(pairs):,} candidate pairs to test.")

    uf = UnionFind(len(recs))
    fullname_matches = []
    pairs_list = list(pairs)
    total_pairs = len(pairs_list)
    for tested, (a_idx, b_idx) in enumerate(pairs_list, 1):
        if full_name_match(recs[a_idx], recs[b_idx]):
            uf.union(a_idx, b_idx)
            fullname_matches.append((a_idx, b_idx))
        progress("Clustering", tested, total_pairs)

    groups = defaultdict(list)
    for r in recs:
        groups[uf.find(r.idx)].append(r.idx)
    groups = list(groups.values())

    print(f"  {len(df):,} rows -> {len(groups):,} merged people "
          f"({len(df) - len(groups):,} rows collapsed by this pass).")
    if fullname_matches:
        print(f"  {len(fullname_matches):,} pair(s) merged via Rule 10 (Full Name "
              f"Match) - see the 'Full Name Match Review' sheet.")

    print("Building merged output ...")
    OTHER_ADDR_COL = "Other Address"
    # Address fields are handled specially (majority address + everything
    # else into Other Address), same as the main script - NOT flattened
    # independently, which would scramble e.g. City/Zip pairings across
    # different source addresses in the group.
    special_cols = set(docid_cols) | {COL_ROWS_MERGED, COL_FIRST, COL_LAST,
                                       COL_MIDDLE, COL_SUFFIX, COL_SSN,
                                       OTHER_ADDR_COL} | set(ADDRESS_COLS)
    other_cols = [c for c in df.columns if c not in special_cols]
    total_groups = len(groups)
    out_rows = []
    for n, group_idxs in enumerate(groups, 1):
        progress("Building output", n, total_groups)
        sub = df.iloc[group_idxs]
        sub_recs = [recs[i] for i in group_idxs]

        row = {c: semicolon_merge(sub[c]) for c in other_cols}
        row[COL_FIRST] = fullest_value(sub[COL_FIRST])
        row[COL_LAST] = fullest_value(sub[COL_LAST])
        row[COL_MIDDLE] = fullest_value(sub[COL_MIDDLE])
        row[COL_SUFFIX] = fullest_value(sub[COL_SUFFIX])
        row[COL_SSN] = fullest_value(sub[COL_SSN])

        majority_addr_values, other_address_new = split_addresses(df, group_idxs)
        for c, v in zip(ADDRESS_COLS, majority_addr_values):
            row[c] = v
        # Combine any Other Address text this group's rows ALREADY carried
        # from the first pass with whatever new alternate address this
        # pass's own majority-picking surfaced.
        existing_other = list(sub[OTHER_ADDR_COL]) if OTHER_ADDR_COL in df.columns else []
        row[OTHER_ADDR_COL] = semicolon_merge(existing_other + ([other_address_new] if other_address_new else []))

        docid_all = semicolon_merge(
            v for c in docid_cols for v in sub[c].tolist()
        )
        docid_chunks = split_docid_chunks(docid_all)
        row[COL_DOCID] = docid_chunks[0]
        for extra_i, chunk in enumerate(docid_chunks[1:], start=2):
            row[f"{COL_DOCID} {extra_i}"] = chunk

        row[COL_ROWS_MERGED] = sum(r.rows_merged for r in sub_recs)
        out_rows.append(row)

    df_out = pd.DataFrame(out_rows)

    max_docid_cols = max(
        (int(c.rsplit(" ", 1)[1]) for c in df_out.columns
         if c.startswith(f"{COL_DOCID} ") and c.rsplit(" ", 1)[1].isdigit()),
        default=1,
    )
    docid_extra_cols = [f"{COL_DOCID} {i}" for i in range(2, max_docid_cols + 1)]
    docid_pos = list(df_out.columns).index(COL_DOCID)
    ordered_other = [c for c in df_out.columns if c not in docid_extra_cols]
    new_order = ordered_other[:docid_pos + 1] + docid_extra_cols + ordered_other[docid_pos + 1:]
    df_out = df_out[new_order]
    df_out = df_out.sort_values([COL_ROWS_MERGED], ascending=False).reset_index(drop=True)

    n_multi = (df_out[COL_ROWS_MERGED] > df_out[COL_ROWS_MERGED].min()).sum()
    print(f"  {len(groups):,} output rows ({len(df) - len(groups):,} first-pass rows "
          f"collapsed further in this pass).")

    # 'Full Name Match Review': every pair merged in THIS pass, side-by-side.
    review_rows = []
    for a_idx, b_idx in fullname_matches:
        full_name = " ".join(p for p in (
            df.at[a_idx, COL_FIRST], df.at[a_idx, COL_MIDDLE],
            df.at[a_idx, COL_LAST], df.at[a_idx, COL_SUFFIX],
        ) if p and str(p).strip().lower() not in ("nan", "none", "null"))
        review_rows.append({
            "Full Name": full_name,
            "DOCID A": df.at[a_idx, COL_DOCID],
            "SSN A": df.at[a_idx, COL_SSN], "DOB A": df.at[a_idx, COL_DOB],
            "Address A": df.at[a_idx, COL_ADDR],
            "DOCID B": df.at[b_idx, COL_DOCID],
            "SSN B": df.at[b_idx, COL_SSN], "DOB B": df.at[b_idx, COL_DOB],
            "Address B": df.at[b_idx, COL_ADDR],
        })
    df_review = pd.DataFrame(review_rows)

    sheets = {
        "Merged Notification Data": df_out,
        "Full Name Match Review": df_review,
    }
    try:
        input_xl = pd.ExcelFile(INPUT_XLSX)
        for sheet_name in PASSTHROUGH_SHEETS:
            if sheet_name in input_xl.sheet_names:
                sheets[sheet_name] = pd.read_excel(INPUT_XLSX, sheet_name=sheet_name)
    except Exception:
        pass   # first-pass review sheets are optional - fine if absent

    print(f"Writing {OUTPUT_XLSX} ...")
    _write_workbook(OUTPUT_XLSX, sheets)
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
