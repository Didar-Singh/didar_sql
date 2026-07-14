"""
260711 pd ds notification merge.py

Merge PII/PHI person records from an Excel export into one row per
confirmed person, for the notification report. Nine rules, built step by
step:

  Rule 1 (SSN Exists): rows with the same (non-blank) SSN are merged,
                        UNLESS:
                          - both rows also have a real DOB and it genuinely
                            disagrees (dob_conflict()) - two different real
                            DOBs mean the shared SSN is more likely wrong/
                            reused than proof of the same identity; or
                          - First, Middle, or Last Name genuinely disagrees
                            between the two rows (name_conflict()) - blank
                            on either side is fine, and an incomplete/
                            truncated entry that's a prefix of the fuller
                            one (e.g. "Did" vs "Didar") is NOT treated as a
                            conflict, but a real difference is; or
                          - Suffix is present on both sides and differs
                            (e.g. "Jr" vs "Sr").
                        In any of these cases the two rows are kept separate
                        rather than merged.
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
  Rule 4 (Employee ID, Name): rows sharing at least one common Employee ID
                        (a cell can already contain multiple semicolon-
                        joined IDs) AND the same exact First+Last Name are
                        merged, as long as SSN and DOB are each either
                        matching or blank on both sides (a real, differing
                        value on both sides blocks it).
  Rule 5 (Driver's License, Name): same as Rule 4, but matched on Driver's
                        License Number instead of Employee ID.
  Rule 6 (Passport Number, Name): same as Rule 4, but matched on Passport
                        Number instead of Employee ID.
  Rule 7 (Phone Number, Name): same as Rule 4, but matched on Phone Number
                        instead of Employee ID.
  Rule 8 (Full Address, Name): Residential Address, City, State, Zip, and
                        Province are each checked individually - blank on
                        either side is fine, but a real, differing value in
                        ANY of them blocks the merge, and at least one field
                        must genuinely match on both sides. Unlike SSN,
                        address is NEVER enough to override a name
                        difference on its own - this also always requires
                        an EXACT Name match and SSN/DOB to be
                        matching-or-blank, same as Rules 4-7.
  Rule 9 (Email, Name): same as Rule 4, but matched on Email Address -
                        Personal instead of Employee ID.

NOTE: a "Rule 10 (Full Name Match)" second pass - for the same person
scattered across rows that each carry a DIFFERENT single piece of
corroborating info (one row has only a DOB, another only an SSN, another
only an Address, etc.), with nothing shared to anchor Rules 1-9 on directly
- lives in a SEPARATE script that runs on THIS script's output. See
"260714 pd ds final merge script.py".

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
OUTPUT : a new Excel workbook with four sheets:
         - "Merged Notification Data": ONE ROW PER CONFIRMED PERSON.
           - First Name, Middle Name, Last Name, and SSN: the single
             fullest/most complete value among the merged rows (placeholder
             values like "[Unknown]" are never picked).
           - Suffix: the single fullest non-blank value.
           - DOB: every distinct value seen among the merged rows, joined
             with "; " (Rule 1 doesn't require DOB to match, so a merged
             group can legitimately contain more than one DOB value - all
             of them are kept, not collapsed to a single "winner").
           - Employee ID, Driver's License, and Passport Number: every
             distinct ID TOKEN seen (cells can already contain multiple
             semicolon-joined IDs - see parse_id_tokens()), deduplicated
             and joined with "; ".
           - Every OTHER column (DOCIDs, Gov ID, etc.): every distinct
             value seen, joined with "; ". If a group's merged DOCID list
             would exceed Excel's 32,767-char cell limit, it spills into
             extra "DOCIDs 2", "DOCIDs 3", ... columns, splitting only at
             "; " boundaries (see split_docid_chunks()).
           - Address fields (Residential Address, City, State, Province,
             Zip, Country) are kept TOGETHER as one unit: the most common
             address (by total row count) stays in those columns - with any
             field left blank on one row filled in from another row's fuller
             copy of that SAME address, so e.g. a blank Zip on one row never
             by itself creates a spurious extra address - and every OTHER,
             genuinely different address goes into a new "Other Address"
             column as one combined string per address, semicolon-joined
             (see address_key_conflict()).
         - "Junk SSN Review": every row where a non-blank SSN value was
           ignored as unusable (a masked SSN with too few known digits) - so
           a real-looking SSN that got silently treated as blank doesn't go
           unnoticed (see classify_ssn_issue()).
         - "Large Group Review": every merged group with more than 50 rows -
           usually a sign of a shared junk value (e.g. a fake SSN), worth a
           manual check before trusting the output.
         - "Skipped Bucket Review": a breakdown, BY FIELD TYPE (not raw
           value - bucket keys can contain real PII), of every oversized
           bucket (>MAX_BUCKET_SIZE rows sharing one value) that got
           skipped during blocking instead of being exhaustively compared.

This script does not touch the input file. Save the output only to the
secured/authorized folder for this data (never a desktop) - it contains
SSN, DOB, and other PII/PHI.

Designed for large row counts (uses "blocking" - only compares rows that
already share an exact SSN, an exact First+Last Name+DOB, a Last Name, or
an Employee ID + Name - instead of comparing every row to every other row).

Install once:
    pip install pandas openpyxl pyxlsb

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
INPUT_XLSX  = "Cng Notification_Final_updated.xlsb"
INPUT_SHEET = 0
OUTPUT_XLSX = "260711 re ds notification merge output.xlsx"

COL_DOCID  = "DOCIDs"
COL_FIRST  = "First Name"
COL_LAST   = "Last Name"
COL_MIDDLE = "Middle Name"
COL_SUFFIX = "Suffix"
COL_DOB    = "Full Date of Birth (MM/DD/YYYY)"
COL_SSN    = "Social Security Number"

# COL_EMPID/COL_DL/COL_PASSPORT/COL_PHONE/COL_EMAIL are used for matching
# (Rules 4-7, 9). COL_GOVID isn't used for matching - kept as a plain
# semicolon-merged column in OTHER_MERGE_COLS below.
COL_DL       = "Driver's License Number"
COL_PASSPORT = "Passport Number"
COL_GOVID    = "Government-Issued ID Number"
COL_EMPID    = "Employee Identification Number"
COL_PHONE    = "Phone Number"
COL_EMAIL    = "Email Address - Personal"

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
    COL_EMAIL,
    COL_PHONE,
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


def _numeric_cell_to_str(v) -> str:
    """Coerces a raw cell value to text without letting pandas' float parsing
    corrupt a whole-number ID (SSN, Employee ID, Driver's License, Passport,
    Phone). A purely-numeric column (no dashes, no leading zero) gets read as
    float64 by pandas whenever that column has even one blank cell elsewhere
    - so a valid value like 123456789 silently becomes 123456789.0, and
    str()'ing that adds a spurious extra digit ('123456789.0' -> 10 digits
    after stripping the dot), making an otherwise-identical ID fail to match
    another row's clean text/int version of the same number. A whole-number
    float is converted via int() first to avoid this."""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


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


# Street-address abbreviation map: each variant -> one canonical token, so a
# street written with a full word and one written with the USPS abbreviation
# ("123 West Lane" vs "123 W Ln") normalize to the same value. Covers the
# common directionals and street-type suffixes; anything not listed is left
# unchanged. Only applied to the Residential Address (street) field.
_STREET_TOKEN_MAP = {
    # Directionals
    "NORTH": "N", "N": "N",
    "SOUTH": "S", "S": "S",
    "EAST": "E", "E": "E",
    "WEST": "W", "W": "W",
    "NORTHEAST": "NE", "NE": "NE",
    "NORTHWEST": "NW", "NW": "NW",
    "SOUTHEAST": "SE", "SE": "SE",
    "SOUTHWEST": "SW", "SW": "SW",
    # Street-type suffixes
    "STREET": "ST", "ST": "ST",
    "AVENUE": "AVE", "AVE": "AVE", "AV": "AVE",
    "BOULEVARD": "BLVD", "BLVD": "BLVD",
    "ROAD": "RD", "RD": "RD",
    "LANE": "LN", "LN": "LN",
    "DRIVE": "DR", "DR": "DR",
    "COURT": "CT", "CT": "CT",
    "CIRCLE": "CIR", "CIR": "CIR",
    "PLACE": "PL", "PL": "PL",
    "TERRACE": "TER", "TERR": "TER", "TER": "TER",
    "PARKWAY": "PKWY", "PKWY": "PKWY",
    "HIGHWAY": "HWY", "HWY": "HWY",
    "SQUARE": "SQ", "SQ": "SQ",
    "TRAIL": "TRL", "TRL": "TRL",
    "WAY": "WAY",
    "LOOP": "LOOP",
    "COVE": "CV", "CV": "CV",
    "POINT": "PT", "PT": "PT",
    "CROSSING": "XING", "XING": "XING",
    "PLAZA": "PLZ", "PLZ": "PLZ",
    "EXPRESSWAY": "EXPY", "EXPY": "EXPY",
    "FREEWAY": "FWY", "FWY": "FWY",
    "ROUTE": "RTE", "RTE": "RTE",
    "JUNCTION": "JCT", "JCT": "JCT",
    "MOUNT": "MT", "MT": "MT",
    "MOUNTAIN": "MTN", "MTN": "MTN",
    # Unit / secondary designators
    "APARTMENT": "APT", "APT": "APT",
    "SUITE": "STE", "STE": "STE",
    "BUILDING": "BLDG", "BLDG": "BLDG",
    "FLOOR": "FL", "FL": "FL",
    "UNIT": "UNIT",
}

_UNIT_DESIGNATORS = {"APT", "STE", "UNIT", "BLDG", "FL"}
# Directional abbreviations - used to tell a real directional ("203 W Shore
# Rd") apart from a bare unit letter loosely attached to the house number
# ("203 A West Shore Rd" - 'A' isn't a direction, so it must be a unit).
_DIRECTIONAL_TOKENS = {"N", "S", "E", "W", "NE", "NW", "SE", "SW"}
_HOUSE_UNIT_RE = re.compile(r"^(\d+)([A-Z])$")


def norm_street(v) -> str:
    """Canonicalizes a street address so common formatting/abbreviation
    differences don't look like different addresses:
    - Directionals (West/W, North/N, ...) and street-type suffixes
      (Lane/Ln, Street/St, Avenue/Ave, ...) are each mapped to one standard
      token, and a trailing '.' on any token is dropped ('St.' -> 'ST').
      So '123 West Lane', '123 W Ln', and '123 W Ln.' all normalize to
      '123 W LN'.
    - A unit letter loosely attached to the house number - either jammed
      against it ('203A ...') or as its own token right after it
      ('203 A ...', where 'A' isn't a directional like 'W' or 'NE') - is
      pulled out and moved to the end as a normal 'UNIT <letter>' suffix.
      So '203A West Shore Rd', '203 A West Shore Rd', and '203 W Shore Rd
      Apt A' all normalize to the same base + unit (see split_street_unit(),
      street_compat())."""
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


def norm_name(v) -> str:
    """Upper/trim; placeholder values ('[Unknown]', 'N/A', ...) become ''
    so they never out-compete or conflict with a real name (Rule 3)."""
    s = norm_text(v)
    core = re.sub(r"[^A-Z0-9]", "", s)
    if not core or core in NAME_PLACEHOLDERS:
        return ""
    return s


SSN_MIN_KNOWN_OVERLAP = 4   # min matching KNOWN digits to trust a masked SSN


def norm_ssn(v) -> str:
    """Return a 9-character pattern of digits and 'X' (X = redacted digit),
    or '' if unusable. Mask characters *, #, ? are treated as X, so
        123-45-6789 -> '123456789'
        123-45-XXXX -> '12345XXXX'
        123-45-6XXX -> '123456XXX'
    Rejected (-> ''): not 9 characters, or a masked SSN with fewer than
    SSN_MIN_KNOWN_OVERLAP known digits (too little information to trust).
    A fully-known 9-digit value is always accepted, even a placeholder-
    looking one (e.g. all-same-digit) - this data has no junk SSNs, so no
    such filtering is applied."""
    if v is None:
        return ""
    s = _numeric_cell_to_str(v).upper().replace("*", "X").replace("#", "X").replace("?", "X")
    kept = re.sub(r"[^0-9X]", "", s)
    if len(kept) != 9:
        return ""
    if "X" not in kept:                                    # fully known
        return kept
    known = sum(c != "X" for c in kept)                     # masked
    return kept if known >= SSN_MIN_KNOWN_OVERLAP else ""


def classify_ssn_issue(v) -> str:
    """Returns a short human-readable reason if a NON-BLANK SSN value was
    rejected by norm_ssn() (wrong length, or a masked SSN with too few known
    digits) - '' if the value was already blank/empty, or was accepted as
    usable. Used to build the 'Junk SSN Review' sheet so a real-looking SSN
    that got silently treated as blank doesn't go unnoticed."""
    if v is None:
        return ""
    raw = str(v).strip()
    if not raw or raw.lower() in ("nan", "none", "null"):
        return ""
    s = _numeric_cell_to_str(v).upper().replace("*", "X").replace("#", "X").replace("?", "X")
    kept = re.sub(r"[^0-9X]", "", s)
    if len(kept) != 9:
        return ""   # not 9 digits at all - not SSN-shaped, not flagged
    if "X" not in kept:
        return ""   # fully valid, always accepted
    known = sum(c != "X" for c in kept)
    if known < SSN_MIN_KNOWN_OVERLAP:
        return (f"Masked SSN with only {known} known digit(s) "
                f"(< {SSN_MIN_KNOWN_OVERLAP} required) - too little to trust")
    return ""   # masked but enough known digits - fine


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


def parse_id_tokens(v) -> frozenset:
    """Splits an Employee ID / Driver's License / Passport / Phone cell into
    individual ID tokens. Cells may already contain multiple semicolon-
    joined IDs from an earlier merge (e.g. '12345; 12346' OR '12345;12346'
    - space after the ';' is optional/inconsistent in the source data), so
    this splits on ';' regardless of spacing and normalizes each token.
    Uses _numeric_cell_to_str() (not a plain str()) so a purely-numeric ID
    that pandas read as a float (e.g. 12345.0) doesn't pick up a spurious
    extra digit and silently fail to match another row's clean version of
    the same ID."""
    if v is None:
        return frozenset()
    return frozenset(norm_text(p) for p in _numeric_cell_to_str(v).split(";") if norm_text(p))


# ------------------------------------------------------------
# 3) Record type - __slots__ for fast attribute access at scale
#    (this loop runs millions of times, so dict-key lookups add up)
# ------------------------------------------------------------
class Rec:
    __slots__ = ("idx", "first", "last", "mid", "suffix", "dob", "ssn", "ssn_present",
                 "empids", "dl_ids", "passport_ids", "phones", "emails",
                 "addr", "city", "state", "zip", "province")

    def __init__(self, idx):
        self.idx = idx


def build_records(df: pd.DataFrame):
    """df MUST already have a 0..n-1 RangeIndex (see main()) - idx doubles
    as the record's position, so no separate index->position map is needed.
    Uses positional numpy-array access (df.values) rather than itertuples(),
    since itertuples() mangles column names with spaces/punctuation.

    Also returns ssn_review: every (DOCID, Original SSN, Reason) where a
    non-blank SSN value was rejected by norm_ssn() (a masked SSN with too
    few known digits) - so a real-looking SSN that got silently treated as
    blank doesn't go unnoticed."""
    recs = []
    ssn_review = []
    col_pos = {c: p for p, c in enumerate(df.columns)}
    values = df.values  # numpy object array, fast positional access
    fi, la, mi, sf, do, ss, ei, dl, pp, ph, em, ad, ci, st, zp, pv = (
        col_pos[COL_FIRST], col_pos[COL_LAST], col_pos[COL_MIDDLE], col_pos[COL_SUFFIX],
        col_pos[COL_DOB], col_pos[COL_SSN], col_pos[COL_EMPID], col_pos[COL_DL],
        col_pos[COL_PASSPORT], col_pos[COL_PHONE], col_pos[COL_EMAIL], col_pos[COL_ADDR],
        col_pos[COL_CITY], col_pos[COL_STATE], col_pos[COL_ZIP], col_pos[COL_PROVINCE],
    )
    docid_pos = col_pos[COL_DOCID]
    for i in range(len(df)):
        row = values[i]
        r = Rec(i)
        r.first = norm_name(row[fi])
        r.last = norm_name(row[la])
        r.mid = norm_name(row[mi])
        r.suffix = norm_name(row[sf])
        r.dob = norm_dob(row[do])
        r.ssn = norm_ssn(row[ss])
        # Distinct from r.ssn: TRUE if the raw cell had any real text at
        # all, even if norm_ssn() rejected it (a masked SSN with too few
        # known digits). That kind of SSN is not trusted for MATCHING
        # (r.ssn stays '' so it can't falsely confirm two rows are the same
        # person) - but it's still evidence the row ISN'T a bare, nothing-
        # to-corroborate-with record, so it must not be treated as "no
        # identifiers" by has_no_identifiers() (which would wrongly let
        # Rule 3 pull it into an unrelated group).
        r.ssn_present = bool(classify_ssn_issue(row[ss])) or bool(r.ssn)
        r.empids = parse_id_tokens(row[ei])
        r.dl_ids = parse_id_tokens(row[dl])
        r.passport_ids = parse_id_tokens(row[pp])
        r.phones = parse_id_tokens(row[ph])
        r.emails = parse_id_tokens(row[em])
        r.addr = norm_street(row[ad])
        r.city = norm_text(row[ci])
        r.state = norm_text(row[st])
        r.zip = norm_text(row[zp])
        r.province = norm_text(row[pv])
        recs.append(r)

        reason = classify_ssn_issue(row[ss])
        if reason:
            ssn_review.append({
                "DOCID": row[docid_pos],
                "First Name": row[fi], "Last Name": row[la],
                "Original SSN": row[ss],
                "Remarks": reason,
            })
    return recs, ssn_review


# ------------------------------------------------------------
# 4) Pairwise matching rules - built STEP BY STEP.
#    Add each newly confirmed rule here as its own small function, then
#    call it from is_match() below.
# ------------------------------------------------------------
def dob_conflict(r1: Rec, r2: Rec) -> bool:
    """True when BOTH rows have a usable DOB and it genuinely disagrees.
    Blocks Step 1 (SSN Exists) - a matching SSN is NOT trusted enough to
    override two different real DOBs, regardless of whether the names
    match, are spelling variants, or are unrelated. Two rows with the same
    SSN but different DOBs are kept as separate entities rather than
    merged into one row with multiple DOB values. Blank DOB on either side
    is not a conflict - Rule 1 still merges purely on a shared SSN then, as
    usual."""
    return bool(r1.dob) and bool(r2.dob) and r1.dob != r2.dob


def name_prefix_compat(a: str, b: str) -> bool:
    """True if two (already normalized) name values are compatible: blank on
    either side, exactly equal, or one is a PREFIX of the other - an
    incomplete/truncated entry (e.g. 'DID' vs 'DIDAR') is not treated as a
    real difference. A genuine difference (neither a prefix of the other)
    returns False."""
    if not a or not b or a == b:
        return True
    return a.startswith(b) or b.startswith(a)


def name_conflict(r1: Rec, r2: Rec) -> bool:
    """True when First, Middle, or Last Name genuinely disagrees between two
    rows (see name_prefix_compat() - blank on either side, or one being a
    truncated/incomplete version of the other, is NOT a conflict), or when
    Suffix is present on both sides and differs (e.g. 'Jr' vs 'Sr' - no
    prefix tolerance for Suffix, since it isn't an abbreviation-of-the-same-
    value situation). Blocks Step 1 (SSN Exists) - a matching SSN is not
    trusted enough to override a genuine name difference."""
    if not name_prefix_compat(r1.first, r2.first):
        return True
    if not name_prefix_compat(r1.mid, r2.mid):
        return True
    if not name_prefix_compat(r1.last, r2.last):
        return True
    return bool(r1.suffix) and bool(r2.suffix) and r1.suffix != r2.suffix


def ssn_exists_match(r1: Rec, r2: Rec) -> bool:
    """Step 1 - 'SSN Exists': both rows have a usable (non-blank) SSN and
    it's the same value, AS LONG AS a genuinely differing DOB
    (dob_conflict()) or a genuinely differing Name/Suffix (name_conflict())
    doesn't block it - a shared SSN alone isn't trusted enough to override
    either of those. An incomplete/truncated name (e.g. 'Did' vs 'Didar') is
    NOT treated as a differing name here - see name_prefix_compat()."""
    if dob_conflict(r1, r2):
        return False
    if name_conflict(r1, r2):
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
    to corroborate identity with. Checks ssn_present (not r.ssn) so a row
    whose SSN got blanked for being junk/placeholder still counts as
    HAVING an SSN here - it's just not trusted for matching. Without this
    distinction, a junk SSN would make the row look like a bare record with
    nothing to go on, wrongly letting Rule 3 (Name-Only) pull it into an
    unrelated group that happens to share the same name."""
    return not r.ssn_present and not r.dob


def zip5(v: str) -> str:
    """First 5 digits of a ZIP code, ignoring hyphens/spaces and any ZIP+4
    suffix - so '62701' and '62701-1234' compare as the SAME base ZIP
    instead of a conflict. '' if fewer than 5 digits (not ZIP-shaped)."""
    digits = re.sub(r"[^0-9]", "", v)
    return digits[:5] if len(digits) >= 5 else ""


def split_street_unit(s: str) -> tuple:
    """Splits a normalized street (see norm_street()) into (base, unit):
    'base' is everything before a unit/apartment designator token (APT,
    STE, UNIT, BLDG, FL, or a bare '#123'-style token), 'unit' is just the
    VALUE after that designator - the designator WORD itself is dropped, so
    'APT A' and 'UNIT A' both give unit 'A' (same physical unit, different
    label word - norm_street() already funnels a bare unit letter through
    to a 'UNIT <letter>' suffix, so this keeps that consistent with an
    explicit 'Apt A' in the source data). If no designator token appears,
    base is the whole string and unit is ''."""
    tokens = s.split(" ") if s else []
    for i, tok in enumerate(tokens):
        if tok in _UNIT_DESIGNATORS:
            return " ".join(tokens[:i]), " ".join(tokens[i + 1:])
        if tok.startswith("#") and len(tok) > 1:
            return " ".join(tokens[:i]), " ".join([tok[1:]] + tokens[i + 1:])
    return s, ""


def street_compat(a: str, b: str) -> bool:
    """True if two normalized streets are equal, blank on either side, or
    the SAME base street with a unit/apartment suffix present on only ONE
    side (e.g. '123 ABC LN' vs '123 ABC LN APT 1' - one row is simply
    missing the unit detail, not a different address). If BOTH sides have
    a real unit and it disagrees (e.g. 'APT 1' vs 'APT 2'), that IS treated
    as a genuine conflict - different specific units at the same street,
    not the same address."""
    if not a or not b or a == b:
        return True
    base_a, unit_a = split_street_unit(a)
    base_b, unit_b = split_street_unit(b)
    if not base_a or base_a != base_b:
        return False
    return not (unit_a and unit_b and unit_a != unit_b)


def address_conflict(r1: Rec, r2: Rec) -> bool:
    """True when Residential Address, City, State, Zip, or Province has a
    real, DIFFERING value on both sides (blank on either side is never a
    conflict here). Zip is compared by its 5-digit prefix (zip5()), so a
    plain 5-digit ZIP and its ZIP+4 form are never treated as conflicting.
    Street is compared via street_compat(), so a bare street and the same
    street with an added apartment/unit suffix are never treated as
    conflicting either - only a genuinely different base street, or two
    real but DIFFERENT unit numbers, count as a conflict. Used to block
    Step 3 from merging two same-named people whose addresses actively
    disagree, with no SSN/DOB to corroborate either way - matching name
    alone isn't enough when the address itself contradicts it."""
    if not street_compat(r1.addr, r2.addr):
        return True
    fields1 = (r1.city, r1.state, zip5(r1.zip), r1.province)
    fields2 = (r2.city, r2.state, zip5(r2.zip), r2.province)
    return any(a and b and a != b for a, b in zip(fields1, fields2))


def no_identifier_name_match(r1: Rec, r2: Rec) -> bool:
    """Step 3 - 'Name-Only, No Identifiers': only applies when at least one
    side has no SSN and no DOB at all (see has_no_identifiers()) - otherwise
    Rule 2 already covers rows that both have a DOB. Matches on Last Name
    (exact) plus a compatible First Name (exact match, or blank on either
    side - a blank First Name is never treated as a conflict), AS LONG AS
    the address doesn't actively conflict (see address_conflict()) - two
    same-named people with genuinely different addresses and nothing else
    to go on should stay separate, not merge on name alone. This is
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
    if address_conflict(r1, r2):
        return False
    return True


def compatible(a: str, b: str) -> bool:
    """True if either side is blank, or both sides are equal. Used for
    Rules 4-6's SSN/DOB check - blank never conflicts, matching is fine,
    but a real, differing value on both sides is a conflict."""
    return not a or not b or a == b


def _id_name_match(ids1: frozenset, ids2: frozenset, r1: Rec, r2: Rec) -> bool:
    """Shared logic for Rules 4-6: rows share at least one common ID token
    (from whichever ID field is being checked - see parse_id_tokens(), a
    cell can already contain multiple semicolon-joined IDs) AND have the
    same exact First+Last Name, AND SSN and DOB are each either matching or
    blank on both sides (a real, differing value on either blocks it)."""
    if not (ids1 and ids2 and not ids1.isdisjoint(ids2)):
        return False
    if not (r1.first and r1.last and r1.first == r2.first and r1.last == r2.last):
        return False
    return compatible(r1.ssn, r2.ssn) and compatible(r1.dob, r2.dob)


def empid_name_match(r1: Rec, r2: Rec) -> bool:
    """Step 4 - 'Employee ID, Name' (see _id_name_match())."""
    return _id_name_match(r1.empids, r2.empids, r1, r2)


def dl_name_match(r1: Rec, r2: Rec) -> bool:
    """Step 5 - 'Driver's License, Name' (see _id_name_match())."""
    return _id_name_match(r1.dl_ids, r2.dl_ids, r1, r2)


def passport_name_match(r1: Rec, r2: Rec) -> bool:
    """Step 6 - 'Passport Number, Name' (see _id_name_match())."""
    return _id_name_match(r1.passport_ids, r2.passport_ids, r1, r2)


def phone_name_match(r1: Rec, r2: Rec) -> bool:
    """Step 7 - 'Phone Number, Name' (see _id_name_match())."""
    return _id_name_match(r1.phones, r2.phones, r1, r2)


def email_name_match(r1: Rec, r2: Rec) -> bool:
    """Step 9 - 'Email, Name' (see _id_name_match())."""
    return _id_name_match(r1.emails, r2.emails, r1, r2)


def address_name_match(r1: Rec, r2: Rec) -> bool:
    """Step 8 - 'Full Address, Name': Residential Address, City, State,
    Zip, and Province are each checked individually - blank on either side
    is fine, but a real, differing value in ANY of them is a conflict that
    blocks the merge. At least one of these fields must have a genuine
    matching value on BOTH sides (two addresses that are entirely blank on
    one side never 'match' just because nothing conflicts).

    Unlike SSN (Rule 1), address is NEVER enough on its own to override a
    name difference - an EXACT Name match is always required here too,
    plus SSN/DOB each being matching-or-blank (compatible()), same as
    Rules 4-7. Zip is compared by its 5-digit prefix (zip5()), so a plain
    5-digit ZIP and its ZIP+4 form count as the same value, not a conflict
    and not a missed overlap. Street is compared via street_compat(), so
    '123 Abc Lane' vs '123 Abc Lane Apt 1' (unit detail on only one side)
    counts as a genuine overlap too, not a conflict and not a miss."""
    if address_conflict(r1, r2):
        return False
    # address_conflict() already confirmed the street is compatible (equal,
    # blank on one side, or same base street) - so both sides having ANY
    # real street text is itself genuine overlap evidence, even when the
    # exact text differs (e.g. one side adds "APT 1").
    street_overlap = bool(r1.addr) and bool(r2.addr)
    other_fields1 = (r1.city, r1.state, zip5(r1.zip), r1.province)
    other_fields2 = (r2.city, r2.state, zip5(r2.zip), r2.province)
    other_overlap = any(a and b and a == b for a, b in zip(other_fields1, other_fields2))
    if not (street_overlap or other_overlap):
        return False   # nothing real actually overlaps - not a match
    if not (r1.first and r1.last and r1.first == r2.first and r1.last == r2.last):
        return False
    return compatible(r1.ssn, r2.ssn) and compatible(r1.dob, r2.dob)


def is_match(r1: Rec, r2: Rec) -> bool:
    return (
        ssn_exists_match(r1, r2)
        or exact_name_dob_match(r1, r2)
        or no_identifier_name_match(r1, r2)
        or empid_name_match(r1, r2)
        or dl_name_match(r1, r2)
        or passport_name_match(r1, r2)
        or phone_name_match(r1, r2)
        or address_name_match(r1, r2)
        or email_name_match(r1, r2)
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
        if r.empids and r.first and r.last:           # Rule 4: Employee ID, Name
            for tok in r.empids:
                buckets[("empidname", tok, r.first, r.last)].append(r.idx)
        if r.dl_ids and r.first and r.last:            # Rule 5: Driver's License, Name
            for tok in r.dl_ids:
                buckets[("dlname", tok, r.first, r.last)].append(r.idx)
        if r.passport_ids and r.first and r.last:      # Rule 6: Passport Number, Name
            for tok in r.passport_ids:
                buckets[("ppname", tok, r.first, r.last)].append(r.idx)
        if r.phones and r.first and r.last:            # Rule 7: Phone Number, Name
            for tok in r.phones:
                buckets[("phonename", tok, r.first, r.last)].append(r.idx)
        if r.emails and r.first and r.last:             # Rule 9: Email, Name
            for tok in r.emails:
                buckets[("emailname", tok, r.first, r.last)].append(r.idx)
        # Rule 8: Full Address, Name - bucket by each address FIELD
        # individually (not the whole address as one key), since
        # address_name_match() only needs ONE field to genuinely overlap.
        # is_match() then does the real per-field compatibility check.
        # State/Province deliberately excluded here (though still checked
        # for conflicts by address_name_match() itself) - they're too
        # low-cardinality for large files (e.g. "CA" alone can be shared by
        # tens of thousands of unrelated rows), which blows a bucket way
        # past MAX_BUCKET_SIZE and gets it skipped entirely - silently
        # losing candidate pairs for Street/City/Zip too, not just State.
        # A real matching address almost always also shares City or Zip,
        # so this loses very little real coverage.
        if r.addr:
            buckets[("addrfield", "addr", r.addr)].append(r.idx)
        if r.city:
            buckets[("addrfield", "city", r.city)].append(r.idx)
        if r.zip:
            buckets[("addrfield", "zip", r.zip)].append(r.idx)

    for (kind, last, first), rep_idx in lastfirst_reps.items():
        blank_key = ("lastblank", last)
        if blank_key in buckets:
            buckets[blank_key].append(rep_idx)

    pairs = set()
    skipped_buckets = 0
    skipped_rows = 0
    # Breakdown by bucket TYPE (not raw value - bucket keys can contain real
    # PII like SSN/Name/State, which shouldn't get printed to the console).
    # This still tells you exactly which rule/field is driving the skips.
    skipped_by_type = defaultdict(lambda: [0, 0])   # kind -> [bucket_count, row_count]
    total = len(buckets)
    for n, (key, idxs) in enumerate(buckets.items(), 1):
        progress("Bucketing", n, total, extra=f"skipped={skipped_buckets}")
        if len(idxs) < 2:
            continue
        if len(idxs) > MAX_BUCKET_SIZE:
            skipped_buckets += 1
            skipped_rows += len(idxs)
            kind = f"{key[0]}/{key[1]}" if key[0] == "addrfield" else key[0]
            entry = skipped_by_type[kind]
            entry[0] += 1
            entry[1] += len(idxs)
            continue
        for a, b in itertools.combinations(sorted(idxs), 2):
            pairs.add((a, b))
    skipped_summary = []
    if skipped_buckets:
        print(f"  Skipped {skipped_buckets:,} oversized bucket(s) "
              f"({skipped_rows:,} rows, likely a shared junk value) - review manually.")
        print("  Breakdown by field (bucket type -> buckets skipped, rows involved):")
        for kind, (cnt, rows) in sorted(skipped_by_type.items(), key=lambda kv: -kv[1][1]):
            print(f"    {kind:<14} {cnt:>5,} bucket(s)  {rows:>10,} rows")
            skipped_summary.append({
                "Bucket Type": kind, "Buckets Skipped": cnt, "Rows Involved": rows,
            })
    return pairs, skipped_summary


# ------------------------------------------------------------
# 7) Merge helpers for building the output
# ------------------------------------------------------------
def semicolon_merge(values) -> str:
    """Distinct, non-blank values joined with '; ', first-seen order,
    original casing preserved; dedup key is upper/trimmed. Splits every
    cell on ';' FIRST and dedupes at that token level (not just whole-cell
    strings) - source cells can already contain multiple semicolon-joined
    sub-values themselves (e.g. '12345; 12346', or '12345;12346' with
    inconsistent spacing, or a literal repeat like '123;456;123'), and
    whole-cell-only dedup would leave those repeats sitting in the output
    untouched. Applies uniformly to every semicolon-merged column, not just
    the ID fields, since this repeated-token pattern can turn up in any
    column depending on how the source data was entered."""
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


DOCID_CHUNK_SIZE = 20_000   # keep well under Excel's 32,767-char cell limit


def split_docid_chunks(docid_str, max_chars=DOCID_CHUNK_SIZE):
    """Splits an already-merged DOCID string ('DOC001; DOC002; ...') into
    chunks no longer than max_chars, breaking ONLY at '; ' boundaries (never
    mid-DOCID). A group merging tens of thousands of rows can produce a
    DOCID string longer than Excel's 32,767-character cell limit, which
    Excel silently truncates - splitting across extra 'DOCIDs 2', 'DOCIDs
    3', ... columns instead keeps every DOCID intact and visible."""
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


def has_variation(raw_values) -> bool:
    """True if the group had 2+ distinct real values for this field, even
    though only one (the fullest) was kept in the output - used for the
    'Names Differ' review flag."""
    return len({norm_text(v) for v in raw_values if norm_text(v)}) > 1


def zip_key(v) -> str:
    """Normalized comparison key for a ZIP/postal code: the 5-digit prefix
    for a US-style ZIP (so '12345' and '12345-1234' compare equal), or the
    plain normalized text for anything else (e.g. non-US postal codes that
    don't have 5+ digits to extract a prefix from)."""
    z5 = zip5(norm_text(v))
    return z5 if z5 else norm_text(v)


def address_key(values) -> tuple:
    """Normalized tuple used to tell whether two rows have the SAME address
    (all fields blank-insensitive) - used to find the majority address.
    ADDRESS_COLS order is Residential Address, City, State, Province, Zip,
    Country - Zip is compared via zip_key() (5-digit prefix) and the street
    via norm_street() (abbreviation-canonicalized), so a plain 5-digit ZIP
    vs its ZIP+4 form, and '123 West Lane' vs '123 W Ln', each count as the
    SAME address here - matching how they're already treated during merging
    (see zip5() / norm_street())."""
    def _key(col, v):
        if col == COL_ZIP:
            return zip_key(v)
        if col == COL_ADDR:
            return norm_street(v)
        return norm_text(v)
    return tuple(_key(col, v) for col, v in zip(ADDRESS_COLS, values))


def address_key_conflict(k1: tuple, k2: tuple) -> bool:
    """True if two normalized address keys (see address_key(), each a
    Street/City/State/Province/Zip/Country tuple) genuinely disagree - blank
    on either side is never a conflict, but a real, differing value is.
    Street is compared via street_compat() (a unit/apt suffix present on
    only one side isn't a conflict, matching how Rule 8 treats it); Zip is
    already the zip5-prefix-or-raw key from address_key(), so a plain
    compare is enough; City/State/Province/Country are a plain blank-
    tolerant equality check. Used by split_addresses() so a row that's
    missing just one field (e.g. blank Zip) of an otherwise-identical
    address is folded into that SAME address instead of being listed as a
    separate 'Other Address' - mirrors how address_conflict() already treats
    blanks when deciding whether two rows are the same person in the first
    place, so the output doesn't contradict the matching decision."""
    addr1, city1, state1, prov1, zip1, country1 = k1
    addr2, city2, state2, prov2, zip2, country2 = k2
    if not street_compat(addr1, addr2):
        return True
    pairs = ((city1, city2), (state1, state2), (prov1, prov2), (zip1, zip2), (country1, country2))
    return any(a and b and a != b for a, b in pairs)


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

    Rows are first bucketed by their exact normalized address_key() (blank-
    insensitive per field, abbreviation-canonicalized), then those DISTINCT
    keys are clustered together whenever they don't genuinely conflict (see
    address_key_conflict()) - e.g. the same street/city with Zip blank on
    one row and present on another are one cluster, one address, not two.
    majority_values: the fullest non-blank value per field (see
    fullest_value()) across the winning cluster - the one with the most
    rows in it - so a row missing a field borrows it from another row's
    fuller copy of the same address, instead of leaving it blank or
    spinning off a spurious 'Other Address' entry. These go into the normal
    Residential Address/City/State/Zip/Country columns.
    other_address_string: every OTHER, genuinely different address cluster,
    combined into one string per address and semicolon-joined, for the
    'Other Address' column. A row with no address at all doesn't count as
    a "real" address unless it's the only kind of address in the group."""
    key_order = []       # distinct keys, first-seen order
    key_count = {}       # key -> row count
    key_raw = {}         # key -> first-seen raw ADDRESS_COLS values
    for idx in group_idxs:
        raw = tuple(df.at[idx, c] for c in ADDRESS_COLS)
        key = address_key(raw)
        if key not in key_count:
            key_count[key] = 0
            key_raw[key] = raw
            key_order.append(key)
        key_count[key] += 1

    # Cluster the DISTINCT keys (typically a handful per person, even in a
    # large merged group) that don't conflict with each other - plain
    # pairwise since there are far fewer distinct addresses than rows.
    parent = list(range(len(key_order)))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[max(ra, rb)] = min(ra, rb)

    for i, j in itertools.combinations(range(len(key_order)), 2):
        if not address_key_conflict(key_order[i], key_order[j]):
            union(i, j)

    clusters = defaultdict(list)
    for i in range(len(key_order)):
        clusters[find(i)].append(i)

    def cluster_weight(positions):
        return sum(key_count[key_order[i]] for i in positions)

    def cluster_values(positions):
        """Fullest non-blank raw value per field across every distinct key
        in this cluster - fills a gap like a blank Zip from another row's
        fuller address rather than leaving it blank."""
        out = []
        for col_i in range(len(ADDRESS_COLS)):
            raws = [key_raw[key_order[i]][col_i] for i in positions]
            norms = [key_order[i][col_i] for i in positions]
            out.append(fullest_value(raws, norms))
        return tuple(out)

    non_blank_clusters = [c for c in clusters.values()
                          if any(any(key_order[i]) for i in c)]
    candidates = non_blank_clusters or list(clusters.values())
    majority_cluster = max(candidates, key=lambda c: (cluster_weight(c), -min(c)))

    other_strings = []
    for c in clusters.values():
        if c is majority_cluster or not any(any(key_order[i]) for i in c):
            continue
        other_strings.append(format_full_address(cluster_values(c)))

    return cluster_values(majority_cluster), MERGE_SEP.join(other_strings)


# ------------------------------------------------------------
# 8) Main
# ------------------------------------------------------------
def main() -> None:
    print(f"Reading {INPUT_XLSX} ...")
    # .xlsb (Excel Binary Workbook) needs the pyxlsb engine explicitly -
    # pandas can't infer it the way it does for .xlsx/.xls.
    engine = "pyxlsb" if INPUT_XLSX.lower().endswith(".xlsb") else None
    df = pd.read_excel(INPUT_XLSX, sheet_name=INPUT_SHEET, dtype=str, engine=engine)
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

    recs, ssn_review = build_records(df)   # recs[i].idx == i == df row position
    if ssn_review:
        print(f"  {len(ssn_review):,} SSN value(s) were ignored as junk/unusable "
              f"- see the 'Junk SSN Review' sheet.")

    print("Clustering (blocked comparison) ...")
    pairs, skipped_bucket_summary = bucket_candidate_pairs(recs)
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
    # group_dob[root] = every DISTINCT known DOB currently inside that
    # root's group. Same bridging risk as group_ssn: a blank-DOB row can
    # pairwise-match two OTHER rows (via Rules 1/4-8, which all treat a
    # blank DOB as non-conflicting) that have different real DOBs from each
    # other, transitively merging two genuinely different DOBs into one
    # group even though those two rows would directly fail dob_conflict().
    group_dob = [({r.dob} if r.dob else set()) for r in recs]
    # group_addr[root] = 5 sets (one per address field: Street, City, State,
    # Zip, Province), each holding every DISTINCT known value currently in
    # that root's group. Same bridging risk again: Rule 8 itself can't
    # bridge through a blank address (it requires a genuine value match on
    # both sides), but a DIFFERENT rule (e.g. Rule 1 via SSN, or Rule 3 via
    # name) can still connect a blank-address row to two OTHER rows that
    # have genuinely conflicting addresses with each other - and since
    # those rules don't check address at all, nothing else would catch it.
    group_addr = [
        (({split_street_unit(r.addr)[0]} if r.addr else set()),   # base street, so a unit/apt suffix doesn't falsely conflict
         ({r.city} if r.city else set()),
         ({r.state} if r.state else set()),
         ({zip5(r.zip)} if zip5(r.zip) else set()),   # 5-digit prefix, so ZIP/ZIP+4 don't falsely conflict
         ({r.province} if r.province else set()))
        for r in recs
    ]
    refused_ssn = 0
    refused_name = 0
    refused_dob = 0
    refused_addr = 0

    def try_union(a_idx, b_idx):
        nonlocal refused_ssn, refused_name, refused_dob, refused_addr
        ra, rb = uf.find(a_idx), uf.find(b_idx)
        if ra == rb:
            return
        sa, sb = group_ssn[ra], group_ssn[rb]
        if sa and sb and sa.isdisjoint(sb):
            refused_ssn += 1
            return   # would combine two different real SSNs - refused
        da, db = group_dob[ra], group_dob[rb]
        if da and db and da.isdisjoint(db):
            refused_dob += 1
            return   # would combine two different real DOBs - refused
        aa, ab = group_addr[ra], group_addr[rb]
        if any(fa and fb and fa.isdisjoint(fb) for fa, fb in zip(aa, ab)):
            refused_addr += 1
            return   # would combine two conflicting real addresses - refused
        r1, r2 = recs[a_idx], recs[b_idx]
        # Rules 2, 4-9 all require an EXACT First+Last match between r1/r2
        # directly, so they can never themselves introduce a first-name
        # conflict. Only Rule 3 (blank-First fallback) can actually
        # introduce one, hence the guard below only matters when none of
        # the stronger rules also apply.
        if not (
            ssn_exists_match(r1, r2) or exact_name_dob_match(r1, r2)
            or empid_name_match(r1, r2) or dl_name_match(r1, r2)
            or passport_name_match(r1, r2) or phone_name_match(r1, r2)
            or address_name_match(r1, r2) or email_name_match(r1, r2)
        ):
            fa, fb = group_first[ra], group_first[rb]
            if fa and fb and fa.isdisjoint(fb):
                refused_name += 1
                return   # weak Rule 3 evidence only - would combine two different real First Names
        uf.union(a_idx, b_idx)
        merged_root = min(ra, rb)
        group_ssn[merged_root] = sa | sb
        group_dob[merged_root] = da | db
        group_addr[merged_root] = tuple(fa | fb for fa, fb in zip(aa, ab))
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
    if refused_dob:
        print(f"  {refused_dob:,} candidate merge(s) were refused - would have "
              f"combined 2+ different real DOBs into one group.")
    if refused_addr:
        print(f"  {refused_addr:,} candidate merge(s) were refused - would have "
              f"combined 2+ conflicting real addresses into one group.")

    print("Building merged output ...")
    SEMICOLON_COLS = [COL_DOCID, COL_DOB] + OTHER_MERGE_COLS
    total_groups = len(groups)
    out_rows = []
    docid_overflow_groups = 0
    max_docid_cols = 1
    for n, group_idxs in enumerate(groups, 1):
        progress("Building output", n, total_groups)
        sub = df.iloc[group_idxs]           # O(group size), not O(n)
        sub_recs = [recs[i] for i in group_idxs]

        row = {c: semicolon_merge(sub[c]) for c in SEMICOLON_COLS}

        # A group merging enough rows can produce a DOCID string longer than
        # Excel's 32,767-char cell limit (silently truncated otherwise) -
        # split it across 'DOCIDs', 'DOCIDs 2', 'DOCIDs 3', ... columns,
        # breaking only at '; ' boundaries so no DOCID is ever cut in half.
        docid_chunks = split_docid_chunks(row[COL_DOCID])
        row[COL_DOCID] = docid_chunks[0]
        for extra_i, chunk in enumerate(docid_chunks[1:], start=2):
            row[f"{COL_DOCID} {extra_i}"] = chunk
        if len(docid_chunks) > 1:
            docid_overflow_groups += 1
            max_docid_cols = max(max_docid_cols, len(docid_chunks))

        row[COL_FIRST] = fullest_value(sub[COL_FIRST], [r.first for r in sub_recs])
        row[COL_LAST] = fullest_value(sub[COL_LAST], [r.last for r in sub_recs])
        row[COL_MIDDLE] = fullest_value(sub[COL_MIDDLE], [r.mid for r in sub_recs])
        row[COL_SUFFIX] = fullest_value(sub[COL_SUFFIX], [norm_text(v) for v in sub[COL_SUFFIX]])
        row[COL_SSN] = fullest_value(sub[COL_SSN], [r.ssn for r in sub_recs])

        majority_addr_values, other_address = split_addresses(df, group_idxs)
        for c, v in zip(ADDRESS_COLS, majority_addr_values):
            row[c] = v
        row["Other Address"] = other_address

        row["Rows Merged"] = len(group_idxs)
        row["Names Differ"] = has_variation(sub[COL_FIRST]) or has_variation(sub[COL_LAST])
        out_rows.append(row)

    df_out = pd.DataFrame(out_rows)

    # Column order: follow the INPUT file's own header sequence (not a
    # fixed order), so the output layout matches the source workbook the
    # user already knows - any 'DOCIDs 2', 'DOCIDs 3', ... overflow columns
    # go right after 'DOCIDs', and columns with no input counterpart
    # ('Other Address', 'Rows Merged', 'Names Differ') go at the end.
    docid_extra_cols = [f"{COL_DOCID} {i}" for i in range(2, max_docid_cols + 1)]
    input_order = list(df.columns)
    extra_cols = [c for c in df_out.columns if c not in input_order and c not in docid_extra_cols]
    new_order = []
    for c in input_order:
        if c not in df_out.columns:
            continue
        new_order.append(c)
        if c == COL_DOCID:
            new_order.extend(docid_extra_cols)
    new_order.extend(extra_cols)
    df_out = df_out[new_order]

    df_out = df_out.sort_values(["Rows Merged"], ascending=False).reset_index(drop=True)

    if docid_overflow_groups:
        print(f"  {docid_overflow_groups:,} group(s) had a DOCID list too long for one "
              f"cell (> {DOCID_CHUNK_SIZE:,} chars) - split across up to "
              f"{max_docid_cols} '{COL_DOCID}' columns.")

    n_multi = (df_out["Rows Merged"] > 1).sum()
    print(f"  {n_multi:,} merged groups combine 2+ original rows.")
    biggest = df_out["Rows Merged"].max()
    print(f"  Largest merged group: {biggest:,} rows.")
    df_large_groups = df_out[df_out["Rows Merged"] > 50].sort_values("Rows Merged", ascending=False)
    if len(df_large_groups):
        print(f"  WARNING: {len(df_large_groups):,} group(s) merged >50 rows - "
              f"usually a shared junk value (e.g. a fake SSN). See the "
              f"'Large Group Review' sheet before trusting the output.")

    df_ssn_review = pd.DataFrame(ssn_review)
    df_skipped_buckets = pd.DataFrame(skipped_bucket_summary)

    print(f"Writing {OUTPUT_XLSX} ...")
    _write_workbook(OUTPUT_XLSX, {
        "Merged Notification Data": df_out,
        "Junk SSN Review": df_ssn_review,
        "Large Group Review": df_large_groups,
        "Skipped Bucket Review": df_skipped_buckets,
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

