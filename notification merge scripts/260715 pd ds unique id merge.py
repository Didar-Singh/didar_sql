"""
260715 pd ds unique id merge.py

Merge PII/PHI person records from an Excel export into ONE ROW PER confirmed
identity, using "Unique ID" as the primary grouping key - but NOT trusting it
blindly. Unlike "260711 pd ds notification merge.py" (which clusters rows
via fuzzy SSN/Name/DOB/ID/Address matching across the WHOLE dataset), here
Unique ID does the heavy lifting, with two safety rules layered on top:

  1. A blank/missing Unique ID never merges with another blank - each such
     row gets its own group, since treating all blanks as "the same ID"
     would incorrectly combine unrelated rows.
  2. A row whose First Name AND Last Name are BOTH entirely blank (or a
     placeholder like "[Unknown]"/"N/A" - see NAME_PLACEHOLDERS/norm_name())
     is NEVER merged with any other row, even if it shares a Unique ID with
     them - there isn't enough identifying information to safely combine it.
  3. Within one Unique ID's rows, two rows are split apart (kept as separate
     people) whenever their First Name AND Last Name BOTH genuinely differ
     (a real, non-blank value on both sides that disagrees) UNLESS a
     matching, real (non-blank on both sides) SSN or DOB confirms they're
     actually the same person despite the name difference - a shared Unique
     ID is not trusted enough to override a genuine identity conflict on its
     own (see name_conflict_blocks_merge()/split_bucket_by_identity()). This
     split uses the same transitive-safe "refuse if it would combine two
     conflicting known identities" pattern as the notification-merge
     script's SSN/DOB/address guards, so a chain of rows can't quietly
     bridge two genuinely different people together.

Per-column merge rule, as specified for this report:

  - First Name, Last Name: the single fullest (longest) non-blank value
    among the group's rows. A placeholder value ("[Unknown]", "N/A", ...)
    never outcompetes a real name (see NAME_PLACEHOLDERS/norm_name()).
  - Middle Name: Max Length value - the single longest non-blank raw value
    (no placeholder filtering).
  - Suffix, Data Subject Type: Max Count value - the most frequently
    occurring non-blank value in the group (ties broken by first-seen
    order).
  - Social Security Number, Birth Information: the single fullest (longest)
    non-blank value.
  - Full Date of Birth (MM/DD/YYYY): the most frequently occurring real date
    among the group (compared by the NORMALIZED date - see norm_dob() - so
    the same date typed in different formats across rows is treated as one
    date, not two), always displayed as a clean "MM/DD/YYYY" string. If no
    row has a parseable date, falls back to the fullest raw text.
  - Employee Identification Number, Address Comments, Email Address -
    Personal, Phone Number, Contact Information, Driver's License Number,
    DL Issuing Country/Province/State, Passport Country/Number, Government
    ID Issuing Country, Government- Issued Identification, Government-Issued
    ID Number, Health/Work/Family/Financial/Student/Demographic/Biometric
    Information, PI Notes, Access Credentials: every distinct value across
    the group, deduplicated and joined with "; " (see semicolon_merge() -
    a cell that already contains multiple semicolon-joined values is split
    and deduped at that token level too).
  - Residential Address, City, State, Province, Zip Code, Country of
    Residence: kept TOGETHER as one unit - the MOST COMMON (max count)
    complete address among the group's rows stays in these columns (with any
    field left blank on one row filled in from another row's fuller copy of
    that SAME address), and every OTHER, genuinely different address goes
    into a new "Other Address" column as one combined string per address,
    semicolon-joined if there's more than one.
  - DOCIDs: every distinct DOCID, deduplicated and joined with "; ". If the
    INPUT already has "DOCIDs 2"/"DOCIDs 3"/"DOCIDs 4" (e.g. it's the output
    of a prior merge that itself overflowed), whichever of those are present
    are read and merged in too - not just "DOCIDs" - so a value that only
    lives in one row's "DOCIDs 3" isn't silently dropped. The combined,
    deduplicated result is then re-chunked for output the same way: if it
    would exceed Excel's 32,767-char cell limit, it spills into "DOCIDs 2",
    "DOCIDs 3", "DOCIDs 4" (splitting only at "; " boundaries) - capped at
    these 4 columns; a group with more DOCIDs than fits across all 4 gets a
    console warning instead of a 5th column.
  - Rows Merged (output-only): count of original rows merged into this
    identity group.
  - Other Address (output-only): see above.

INPUT  : an Excel workbook with the columns listed in EXPECTED_COLS below.
OUTPUT : a new Excel workbook, one sheet, "Merged Data" - one row per
         confirmed identity (a Unique ID can produce MORE than one output
         row if its rows split apart per rule 3 above).

This script does not touch the input file. Save the output only to the
secured/authorized folder for this data (never a desktop) - it contains
SSN, DOB, and other PII/PHI.

Install once:
    pip install pandas openpyxl pyxlsb

Run (uses INPUT_XLSX/OUTPUT_XLSX from the CONFIG block below as-is):
    python "260715 pd ds unique id merge.py"

Or override the input/output file on the command line instead of editing
the CONFIG block:
    python "260715 pd ds unique id merge.py" input.xlsx
    python "260715 pd ds unique id merge.py" input.xlsx output.xlsx
"""

import sys
import re
import itertools
import unicodedata
from collections import defaultdict

import pandas as pd

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
INPUT_XLSX  = "INPUT_FILE.xlsx"   # edit this, or pass the path on the command line
INPUT_SHEET = 0
OUTPUT_XLSX = "260715 re ds unique id merge output.xlsx"

COL_UNIQUEID = "Unique ID"
COL_DOCID    = "DOCIDs"
COL_FIRST    = "First Name"
COL_MIDDLE   = "Middle Name"
COL_LAST     = "Last Name"
COL_SUFFIX   = "Suffix"
COL_SSN      = "Social Security Number"
COL_DOB      = "Full Date of Birth (MM/DD/YYYY)"
COL_DATASUBJECT = "Data Subject Type"
COL_BIRTHINFO   = "Birth Information"

# Address fields are handled specially (see ADDRESS_COLS below), not as
# plain semicolon-merged columns - they need to stay together as one unit
# per address, not be shuffled independently per field.
COL_ADDR     = "Residential Address"
COL_CITY     = "City"
COL_STATE    = "State of Residence (if US)"
COL_PROVINCE = "Province of Residence (if Canada)"
COL_ZIP      = "Zip Code"
COL_COUNTRY  = "Country of Residence"
ADDRESS_COLS = [COL_ADDR, COL_CITY, COL_STATE, COL_PROVINCE, COL_ZIP, COL_COUNTRY]

# Output-only column holding every non-majority address (see split_addresses()
# below). Not part of ADDRESS_COLS/EXPECTED_COLS - but if the INPUT already
# has it (e.g. it's the output of a prior merge), its content is carried
# forward and merged in too, rather than being silently overwritten by this
# run's own freshly-computed "Other Address" value.
COL_OTHER_ADDR = "Other Address"

# "Merge With Semicolon" columns - every distinct value joined with "; ".
SEMICOLON_COLS = [
    "Employee Identification Number",
    "Address Comments",
    "Email Address - Personal",
    "Phone Number",
    "Contact Information",
    "Driver's License Number",
    "DL Issuing Country",
    "DL Issuing Province (if Canada)",
    "DL Issuing State (if US)",
    "Passport Country",
    "Passport Number",
    "Government ID Issuing Country",
    "Government- Issued Identification",
    "Government-Issued ID Number",
    "Health Related Information",
    "Work-Related Information",
    "Family Information",
    "Financial Account Information",
    "Student-Related Information",
    "Demographic Information",
    "Biometric Data",
    "PI Notes",
    "Access Credentials (Non-Financial Account)",
]

# "Max Length" / fullest-value columns (no placeholder filtering).
FULLEST_COLS = [COL_MIDDLE, COL_SSN, COL_BIRTHINFO]

# Fullest-value columns that additionally skip placeholder values like
# "[Unknown]"/"N/A" (see NAME_PLACEHOLDERS/norm_name()).
NAME_FULLEST_COLS = [COL_FIRST, COL_LAST]

# "Max Count" / mode columns - most frequently occurring non-blank value.
MODE_COLS = [COL_SUFFIX, COL_DATASUBJECT]

EXPECTED_COLS = (
    [COL_UNIQUEID, COL_DOCID] + NAME_FULLEST_COLS + FULLEST_COLS
    + [COL_DOB] + MODE_COLS + ADDRESS_COLS + SEMICOLON_COLS
)

MERGE_SEP = "; "
EXCEL_MAX_ROWS = 1_048_576
DOCID_CHUNK_SIZE = 20_000   # keep well under Excel's 32,767-char cell limit
MAX_DOCID_COLS = 4          # fixed: DOCIDs, DOCIDs 2, DOCIDs 3, DOCIDs 4
OTHER_ADDR_CHUNK_SIZE = 25_000   # keep well under Excel's 32,767-char cell limit
MAX_OTHER_ADDR_COLS = 4          # fixed: Other Address, Other Address 2, 3, 4

# Placeholder name values treated as blank (never win a fullest-value pick;
# a real name always supersedes these). Checked after stripping brackets/
# parens/periods, so "[Unknown]", "(unknown)", "N/A" all match.
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


def norm_text(v) -> str:
    """Used as the DEDUP KEY everywhere (grouping, semicolon_merge, mode_value,
    address_key). Normalizes unicode (NFKC), strips invisible characters,
    collapses whitespace, and upper-cases."""
    if v is None:
        return ""
    s = unicodedata.normalize("NFKC", str(v))
    s = _INVISIBLE_CHARS.sub(" ", s)
    s = re.sub(r"\s+", " ", s).strip().upper()
    return "" if s.lower() in ("nan", "none", "null") else s


def norm_name(v) -> str:
    """Like norm_text(), but placeholder values ('[Unknown]', 'N/A', ...)
    become '' so they never win a fullest-value pick over a real name."""
    s = norm_text(v)
    core = re.sub(r"[^A-Z0-9]", "", s)
    if not core or core in NAME_PLACEHOLDERS:
        return ""
    return s


# Excel's date epoch: day 1 = 1900-01-01, but Excel treats 1900 as a leap
# year (it wasn't) - using 1899-12-30 as day 0 reproduces that quirk, so a
# serial number converts to the SAME date Excel itself displays.
_EXCEL_SERIAL_EPOCH = pd.Timestamp("1899-12-30")
_EXCEL_SERIAL_RE = re.compile(r"\d{1,6}(\.\d+)?")


def norm_dob(v) -> str:
    """Parses a DOB cell into 'YYYYMMDD', or '' if unparseable/blank.
    Also handles a raw Excel SERIAL date number (e.g. '20037') showing up
    as the cell's text instead of an actual date (see the .xlsb/pyxlsb note
    in the main notification-merge script for why this happens)."""
    if v is None:
        return ""
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return ""
    if _EXCEL_SERIAL_RE.fullmatch(s):
        serial = int(float(s))
        if 1 <= serial <= 60000:   # sane range: years ~1900-2064
            ts = _EXCEL_SERIAL_EPOCH + pd.Timedelta(days=serial)
            return ts.strftime("%Y%m%d")
    ts = pd.to_datetime(s, errors="coerce")
    return "" if pd.isna(ts) else ts.strftime("%Y%m%d")


def ssn_digits(v) -> str:
    """Digits-only SSN, or '' if not exactly 9 digits. Used ONLY for the
    identity-conflict guard (see split_bucket_by_identity()) - a real,
    matching SSN here is trusted enough to override a genuine Name conflict.
    This is separate from the 'Social Security Number' OUTPUT column itself
    (see fullest_value()), which keeps the value's original formatting."""
    if v is None:
        return ""
    digits = re.sub(r"[^0-9]", "", str(v))
    return digits if len(digits) == 9 else ""


def zip5(v: str) -> str:
    """First 5 digits of a ZIP code, ignoring hyphens/spaces and any ZIP+4
    suffix - so '62701' and '62701-1234' compare as the SAME base ZIP
    instead of a conflict. '' if fewer than 5 digits (not ZIP-shaped)."""
    digits = re.sub(r"[^0-9]", "", v)
    return digits[:5] if len(digits) >= 5 else ""


def zip_key(v) -> str:
    """Normalized comparison key for a ZIP/postal code: the 5-digit prefix
    for a US-style ZIP, or the plain normalized text otherwise (e.g. non-US
    postal codes without 5+ digits to extract a prefix from)."""
    z5 = zip5(norm_text(v))
    return z5 if z5 else norm_text(v)


# Street-address abbreviation map: each variant -> one canonical token, so a
# street written with a full word and one written with the USPS abbreviation
# ("123 West Lane" vs "123 W Ln") normalize to the same value.
_STREET_TOKEN_MAP = {
    "NORTH": "N", "N": "N", "SOUTH": "S", "S": "S", "EAST": "E", "E": "E",
    "WEST": "W", "W": "W", "NORTHEAST": "NE", "NE": "NE",
    "NORTHWEST": "NW", "NW": "NW", "SOUTHEAST": "SE", "SE": "SE",
    "SOUTHWEST": "SW", "SW": "SW",
    "STREET": "ST", "ST": "ST", "AVENUE": "AVE", "AVE": "AVE", "AV": "AVE",
    "BOULEVARD": "BLVD", "BLVD": "BLVD", "ROAD": "RD", "RD": "RD",
    "LANE": "LN", "LN": "LN", "DRIVE": "DR", "DR": "DR",
    "COURT": "CT", "CT": "CT", "CIRCLE": "CIR", "CIR": "CIR",
    "PLACE": "PL", "PL": "PL", "TERRACE": "TER", "TERR": "TER", "TER": "TER",
    "PARKWAY": "PKWY", "PKWY": "PKWY", "HIGHWAY": "HWY", "HWY": "HWY",
    "SQUARE": "SQ", "SQ": "SQ", "TRAIL": "TRL", "TRL": "TRL",
    "WAY": "WAY", "LOOP": "LOOP", "COVE": "CV", "CV": "CV",
    "POINT": "PT", "PT": "PT", "CROSSING": "XING", "XING": "XING",
    "PLAZA": "PLZ", "PLZ": "PLZ", "EXPRESSWAY": "EXPY", "EXPY": "EXPY",
    "FREEWAY": "FWY", "FWY": "FWY", "ROUTE": "RTE", "RTE": "RTE",
    "JUNCTION": "JCT", "JCT": "JCT", "MOUNT": "MT", "MT": "MT",
    "MOUNTAIN": "MTN", "MTN": "MTN",
    "APARTMENT": "APT", "APT": "APT", "SUITE": "STE", "STE": "STE",
    "BUILDING": "BLDG", "BLDG": "BLDG", "FLOOR": "FL", "FL": "FL",
    "UNIT": "UNIT",
}
_UNIT_DESIGNATORS = {"APT", "STE", "UNIT", "BLDG", "FL"}
_DIRECTIONAL_TOKENS = {"N", "S", "E", "W", "NE", "NW", "SE", "SW"}
_HOUSE_UNIT_RE = re.compile(r"^(\d+)([A-Z])$")


def norm_street(v) -> str:
    """Canonicalizes a street address so common formatting/abbreviation
    differences don't look like different addresses (directionals and
    street-type suffixes mapped to one standard token; a unit letter loosely
    attached to the house number is pulled out to a trailing 'UNIT <letter>')."""
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
    """Splits a normalized street into (base, unit): 'unit' is just the VALUE
    after a unit/apartment designator token (APT, STE, UNIT, BLDG, FL, or a
    bare '#123'-style token); base is everything before it."""
    tokens = s.split(" ") if s else []
    for i, tok in enumerate(tokens):
        if tok in _UNIT_DESIGNATORS:
            return " ".join(tokens[:i]), " ".join(tokens[i + 1:])
        if tok.startswith("#") and len(tok) > 1:
            return " ".join(tokens[:i]), " ".join([tok[1:]] + tokens[i + 1:])
    return s, ""


def street_compat(a: str, b: str) -> bool:
    """True if two normalized streets are equal, blank on either side, or the
    SAME base street with a unit/apartment suffix present on only ONE side."""
    if not a or not b or a == b:
        return True
    base_a, unit_a = split_street_unit(a)
    base_b, unit_b = split_street_unit(b)
    if not base_a or base_a != base_b:
        return False
    return not (unit_a and unit_b and unit_a != unit_b)


# ------------------------------------------------------------
# 3) Column-level merge helpers
# ------------------------------------------------------------
def fullest_value(values, skip_placeholders: bool = False) -> str:
    """Longest non-blank raw value across the group ('Max Length' / fullest-
    value rule). If skip_placeholders is True, a placeholder value like
    '[Unknown]' or 'N/A' is treated as blank (see norm_name()) so a real
    name always wins over it."""
    best, best_len = "", -1
    for v in values:
        key = norm_name(v) if skip_placeholders else norm_text(v)
        if not key:
            continue
        raw = "" if v is None else str(v).strip()
        if len(raw) > best_len:
            best, best_len = raw, len(raw)
    return best


def mode_value(values) -> str:
    """Most frequently occurring non-blank value across the group ('Max
    Count' rule); ties broken by first-seen order. Returns the first-seen
    raw representation of the winning value (preserves original casing)."""
    counts = {}
    order = []
    first_raw = {}
    for v in values:
        key = norm_text(v)
        if not key:
            continue
        if key not in counts:
            counts[key] = 0
            order.append(key)
            first_raw[key] = "" if v is None else str(v).strip()
        counts[key] += 1
    if not order:
        return ""
    best_key = max(order, key=lambda k: counts[k])   # max() returns the first max on a tie
    return first_raw[best_key]


def dob_value(values) -> str:
    """Most frequently occurring real DOB across the group (compared by the
    NORMALIZED date - see norm_dob() - so the same date typed in different
    formats collapses into one), displayed as a clean 'MM/DD/YYYY' string.
    Falls back to fullest_value() if no row has a parseable date, so a
    non-blank-but-unparseable entry isn't silently dropped."""
    counts = {}
    order = []
    for v in values:
        key = norm_dob(v)
        if not key:
            continue
        if key not in counts:
            counts[key] = 0
            order.append(key)
        counts[key] += 1
    if not order:
        return fullest_value(values)
    best_key = max(order, key=lambda k: counts[k])
    return f"{best_key[4:6]}/{best_key[6:8]}/{best_key[0:4]}"


def semicolon_merge(values) -> str:
    """Distinct, non-blank values joined with '; ', first-seen order,
    original casing preserved. Splits every cell on ';' FIRST and dedupes at
    that token level (not just whole-cell strings) - source cells can already
    contain multiple semicolon-joined sub-values themselves."""
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


def split_docid_chunks(docid_str, max_chars=DOCID_CHUNK_SIZE, max_cols=MAX_DOCID_COLS):
    """Splits an already-merged DOCID string into chunks no longer than
    max_chars, breaking ONLY at '; ' boundaries (never mid-DOCID) - capped at
    max_cols total chunks. If content remains after the (max_cols - 1)th
    split, everything left is packed into the final chunk (even if it then
    exceeds max_chars) rather than spilling into a 5th column."""
    if len(docid_str) <= max_chars:
        return [docid_str]
    parts = docid_str.split(MERGE_SEP)
    chunks = []
    current = []
    current_len = 0
    for part in parts:
        added_len = len(part) + (len(MERGE_SEP) if current else 0)
        if (current and current_len + added_len > max_chars
                and len(chunks) < max_cols - 1):
            chunks.append(MERGE_SEP.join(current))
            current = [part]
            current_len = len(part)
        else:
            current.append(part)
            current_len += added_len
    if current:
        chunks.append(MERGE_SEP.join(current))
    return chunks


def format_full_address(values) -> str:
    """One address, all its parts combined into a single readable string,
    e.g. '123 ABC Ln, Springfield, IL, 62701, USA'. Blank parts are skipped."""
    parts = []
    for v in values:
        s = "" if v is None else str(v).strip()
        if s and s.lower() not in ("nan", "none", "null"):
            parts.append(s)
    return ", ".join(parts)


def address_key(values) -> tuple:
    """Normalized tuple used to tell whether two rows have the SAME address
    (all fields blank-insensitive). ADDRESS_COLS order is Residential
    Address, City, State, Province, Zip, Country."""
    def _key(col, v):
        if col == COL_ZIP:
            return zip_key(v)
        if col == COL_ADDR:
            return norm_street(v)
        return norm_text(v)
    return tuple(_key(col, v) for col, v in zip(ADDRESS_COLS, values))


def address_key_conflict(k1: tuple, k2: tuple) -> bool:
    """True if two normalized address keys genuinely disagree - blank on
    either side is never a conflict, but a real, differing value is."""
    addr1, city1, state1, prov1, zip1, country1 = k1
    addr2, city2, state2, prov2, zip2, country2 = k2
    if not street_compat(addr1, addr2):
        return True
    pairs = ((city1, city2), (state1, state2), (prov1, prov2), (zip1, zip2), (country1, country2))
    return any(a and b and a != b for a, b in pairs)


def split_addresses(df, group_idxs):
    """Returns (majority_values, other_address_string) for one merged group.

    Rows are first bucketed by their exact normalized address_key(), then
    those DISTINCT keys are clustered together whenever they don't genuinely
    conflict (see address_key_conflict()) - e.g. the same street/city with
    Zip blank on one row and present on another are one cluster, one
    address, not two. majority_values: the fullest non-blank value per field
    across the winning cluster - the one with the most rows in it ('Max
    Count' rule). other_address_string: every OTHER, genuinely different
    address cluster, combined into one string per address and semicolon-
    joined, for the 'Other Address' column.

    Uses the same transitive-safe Union-Find pattern as
    split_bucket_by_identity(): a pairwise union is refused (not just
    skipped) whenever it would combine two clusters that each already
    contain a key conflicting with a key on the other side - checking every
    key already accumulated in each cluster, not just the two specific keys
    being compared. This stops a fully-blank address row from acting as a
    'bridge' that transitively merges two otherwise-conflicting addresses
    into one (which would silently drop one of them, since a single merged
    cluster leaves nothing left over for 'Other Address')."""
    key_order = []
    key_count = {}
    key_raw = {}
    for idx in group_idxs:
        raw = tuple(df.at[idx, c] for c in ADDRESS_COLS)
        key = address_key(raw)
        if key not in key_count:
            key_count[key] = 0
            key_raw[key] = raw
            key_order.append(key)
        key_count[key] += 1

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

    key_sets = [{key_order[i]} for i in range(len(key_order))]

    def keys_conflict(set1, set2):
        return any(address_key_conflict(k1, k2) for k1 in set1 for k2 in set2)

    for i, j in itertools.combinations(range(len(key_order)), 2):
        ri, rj = find(i), find(j)
        if ri == rj:
            continue
        if keys_conflict(key_sets[ri], key_sets[rj]):
            continue   # refused - would bridge two conflicting addresses
        union(i, j)
        new_root, old_root = min(ri, rj), max(ri, rj)
        key_sets[new_root] |= key_sets[old_root]

    clusters = defaultdict(list)
    for i in range(len(key_order)):
        clusters[find(i)].append(i)

    def cluster_weight(positions):
        return sum(key_count[key_order[i]] for i in positions)

    def cluster_values(positions):
        out = []
        for col_i in range(len(ADDRESS_COLS)):
            raws = [key_raw[key_order[i]][col_i] for i in positions]
            out.append(fullest_value(raws))
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
# 3b) Identity-conflict guard - don't blindly trust a shared Unique ID.
# ------------------------------------------------------------
def name_conflict_blocks_merge(g1: tuple, g2: tuple) -> bool:
    """True if two identity-groups (each a (first_names, last_names, ssns,
    dobs) tuple of sets - see split_bucket_by_identity()) must NOT be merged:
    both have a real, known First Name AND a real, known Last Name, and
    BOTH disagree with the other group's - UNLESS a real SSN or DOB is
    known on both sides AND matches (which is trusted enough to confirm
    they're the same person despite the name difference)."""
    first1, last1, ssn1, dob1 = g1
    first2, last2, ssn2, dob2 = g2
    names_conflict = (
        bool(first1) and bool(first2) and first1.isdisjoint(first2)
        and bool(last1) and bool(last2) and last1.isdisjoint(last2)
    )
    if not names_conflict:
        return False
    ssn_confirms = bool(ssn1) and bool(ssn2) and not ssn1.isdisjoint(ssn2)
    dob_confirms = bool(dob1) and bool(dob2) and not dob1.isdisjoint(dob2)
    return not (ssn_confirms or dob_confirms)


def split_bucket_by_identity(df, idxs):
    """Given a list of row indices that already share one (normalized)
    Unique ID, further splits them into sub-groups whenever two rows have a
    genuine First+Last Name conflict not confirmed-away by a matching SSN or
    DOB (see name_conflict_blocks_merge()) - a shared Unique ID alone is not
    trusted enough to override that.

    Uses the same transitive-safe Union-Find pattern as split_addresses():
    each row starts in its own singleton, and a pairwise union is refused
    (not just skipped) whenever it would combine two GROUP-LEVEL conflicting
    identities - checking the accumulated set of every name/SSN/DOB already
    inside each side's cluster, not just the two specific rows being
    compared. This stops a 'bridge' row (e.g. blank name) from transitively
    linking two otherwise-conflicting people together. Returns a list of
    index-lists (sub-groups); a single-row input returns unchanged."""
    n = len(idxs)
    if n == 1:
        return [idxs]

    parent = list(range(n))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[max(ra, rb)] = min(ra, rb)

    first_sets, last_sets, ssn_sets, dob_sets = [], [], [], []
    for idx in idxs:
        f = norm_name(df.at[idx, COL_FIRST])
        l = norm_name(df.at[idx, COL_LAST])
        s = ssn_digits(df.at[idx, COL_SSN])
        d = norm_dob(df.at[idx, COL_DOB])
        first_sets.append({f} if f else set())
        last_sets.append({l} if l else set())
        ssn_sets.append({s} if s else set())
        dob_sets.append({d} if d else set())

    for i, j in itertools.combinations(range(n), 2):
        ri, rj = find(i), find(j)
        if ri == rj:
            continue
        g1 = (first_sets[ri], last_sets[ri], ssn_sets[ri], dob_sets[ri])
        g2 = (first_sets[rj], last_sets[rj], ssn_sets[rj], dob_sets[rj])
        if name_conflict_blocks_merge(g1, g2):
            continue   # refused - keep these apart despite the shared Unique ID
        union(i, j)
        new_root, old_root = min(ri, rj), max(ri, rj)
        first_sets[new_root] |= first_sets[old_root]
        last_sets[new_root] |= last_sets[old_root]
        ssn_sets[new_root] |= ssn_sets[old_root]
        dob_sets[new_root] |= dob_sets[old_root]

    clusters = defaultdict(list)
    for i in range(n):
        clusters[find(i)].append(idxs[i])
    return list(clusters.values())


# ------------------------------------------------------------
# 4) Main
# ------------------------------------------------------------
def main() -> None:
    print(f"Reading {INPUT_XLSX} ...")
    # .xlsb (Excel Binary Workbook) needs the pyxlsb engine explicitly -
    # pandas can't infer it the way it does for .xlsx/.xls.
    engine = "pyxlsb" if INPUT_XLSX.lower().endswith(".xlsb") else None
    df = pd.read_excel(INPUT_XLSX, sheet_name=INPUT_SHEET, dtype=str, engine=engine)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.reset_index(drop=True)   # guarantees row position == iloc position

    missing = [c for c in EXPECTED_COLS if c not in df.columns]
    if missing:
        raise SystemExit(
            f"These expected columns were not found in {INPUT_XLSX}:\n"
            f"  {missing}\nColumns present:\n  {list(df.columns)}\n"
            "Fix the COL_*/SEMICOLON_COLS names in the CONFIG block."
        )
    print(f"  {len(df):,} rows read.")

    # "DOCIDs" is always required (see EXPECTED_COLS above), but the input
    # may ALSO already have "DOCIDs 2"/"DOCIDs 3"/"DOCIDs 4" - e.g. it's the
    # output of a prior merge that itself overflowed into those columns.
    # Whichever of those are present get read and merged in too, so a value
    # that only lives in one row's "DOCIDs 3" isn't silently dropped.
    docid_input_cols = [COL_DOCID] + [
        f"{COL_DOCID} {i}" for i in range(2, MAX_DOCID_COLS + 1)
        if f"{COL_DOCID} {i}" in df.columns
    ]
    if len(docid_input_cols) > 1:
        print(f"  Input already has overflow DOCID columns "
              f"({', '.join(docid_input_cols[1:])}) - merging them in too.")

    # "Other Address" is optional (not in EXPECTED_COLS - a fresh source file
    # won't have it), but if the input already has it - plus any
    # "Other Address 2"/"Other Address 3"/"Other Address 4" from a prior
    # merge that itself overflowed - all of those are read and merged in too.
    other_addr_input_cols = [c for c in (
        [COL_OTHER_ADDR] + [f"{COL_OTHER_ADDR} {i}" for i in range(2, MAX_OTHER_ADDR_COLS + 1)]
    ) if c in df.columns]
    if len(other_addr_input_cols) > 1:
        print(f"  Input already has overflow Other Address columns "
              f"({', '.join(other_addr_input_cols[1:])}) - merging them in too.")

    # Bucket by normalized Unique ID. A row is instead isolated into its own
    # singleton bucket (never merged with anything) when either:
    #   - its Unique ID is blank/missing (a blank never merges with another
    #     blank - treating all blanks as "the same ID" would incorrectly
    #     combine unrelated rows), or
    #   - First Name AND Last Name are BOTH entirely blank/placeholder (see
    #     norm_name()) - not enough identifying information to safely merge.
    blank_counter = itertools.count()
    groups_map = defaultdict(list)
    for i in range(len(df)):
        first_blank = not norm_name(df.at[i, COL_FIRST])
        last_blank = not norm_name(df.at[i, COL_LAST])
        if first_blank and last_blank:
            groups_map[f"__noname_{next(blank_counter)}__"].append(i)
            continue
        uid = norm_text(df.at[i, COL_UNIQUEID])
        key = uid or f"__blankuid_{next(blank_counter)}__"
        groups_map[key].append(i)

    # Within each Unique ID's bucket, split apart any rows whose First+Last
    # Name genuinely conflict and aren't confirmed-same by a matching SSN or
    # DOB (see split_bucket_by_identity()) - a shared Unique ID is a strong
    # signal, not an absolute one.
    groups = []
    split_bucket_count = 0
    for idxs in groups_map.values():
        sub_groups = split_bucket_by_identity(df, idxs)
        if len(sub_groups) > 1:
            split_bucket_count += 1
        groups.extend(sub_groups)

    print(f"  {len(df):,} rows -> {len(groups):,} identity groups "
          f"({len(df) - len(groups):,} rows collapsed by merging).")
    if split_bucket_count:
        print(f"  {split_bucket_count:,} Unique ID(s) were split into 2+ separate "
              f"identity groups due to a Name conflict not confirmed by a "
              f"matching SSN/DOB.")

    out_rows = []
    docid_overflow_groups = 0
    max_docid_cols_used = 1
    other_addr_overflow_groups = 0
    max_other_addr_cols_used = 1
    total_groups = len(groups)
    for n, group_idxs in enumerate(groups, 1):
        progress("Building output", n, total_groups)
        sub = df.iloc[group_idxs]

        row = {COL_UNIQUEID: fullest_value(sub[COL_UNIQUEID])}

        row[COL_FIRST] = fullest_value(sub[COL_FIRST], skip_placeholders=True)
        row[COL_LAST] = fullest_value(sub[COL_LAST], skip_placeholders=True)
        for c in FULLEST_COLS:
            row[c] = fullest_value(sub[c])
        for c in MODE_COLS:
            row[c] = mode_value(sub[c])
        row[COL_DOB] = dob_value(sub[COL_DOB])

        for c in SEMICOLON_COLS:
            row[c] = semicolon_merge(sub[c])

        docid_merged = semicolon_merge(
            itertools.chain.from_iterable(sub[c] for c in docid_input_cols)
        )
        docid_chunks = split_docid_chunks(docid_merged)
        row[COL_DOCID] = docid_chunks[0]
        for extra_i, chunk in enumerate(docid_chunks[1:], start=2):
            row[f"{COL_DOCID} {extra_i}"] = chunk
        if len(docid_chunks) > 1:
            docid_overflow_groups += 1
            max_docid_cols_used = max(max_docid_cols_used, len(docid_chunks))
        if len(docid_chunks) == MAX_DOCID_COLS and len(docid_chunks[-1]) > DOCID_CHUNK_SIZE:
            print(f"\n  WARNING: Unique ID {row[COL_UNIQUEID]!r} has more DOCIDs than "
                  f"fit across {MAX_DOCID_COLS} columns - '{COL_DOCID} {MAX_DOCID_COLS}' "
                  f"exceeds Excel's comfortable cell size and may display truncated.")

        majority_addr_values, other_address = split_addresses(df, group_idxs)
        for c, v in zip(ADDRESS_COLS, majority_addr_values):
            row[c] = v
        if other_addr_input_cols:
            # Preserve whatever the input already had here (e.g. from a
            # prior merge pass) instead of overwriting it with only this
            # run's own newly-computed value.
            other_address = semicolon_merge(itertools.chain(
                itertools.chain.from_iterable(sub[c] for c in other_addr_input_cols),
                [other_address],
            ))
        other_addr_chunks = split_docid_chunks(
            other_address, max_chars=OTHER_ADDR_CHUNK_SIZE, max_cols=MAX_OTHER_ADDR_COLS
        )
        row[COL_OTHER_ADDR] = other_addr_chunks[0]
        for extra_i, chunk in enumerate(other_addr_chunks[1:], start=2):
            row[f"{COL_OTHER_ADDR} {extra_i}"] = chunk
        if len(other_addr_chunks) > 1:
            other_addr_overflow_groups += 1
            max_other_addr_cols_used = max(max_other_addr_cols_used, len(other_addr_chunks))
        if (len(other_addr_chunks) == MAX_OTHER_ADDR_COLS
                and len(other_addr_chunks[-1]) > OTHER_ADDR_CHUNK_SIZE):
            print(f"\n  WARNING: Unique ID {row[COL_UNIQUEID]!r} has more Other "
                  f"Address entries than fit across {MAX_OTHER_ADDR_COLS} columns - "
                  f"'{COL_OTHER_ADDR} {MAX_OTHER_ADDR_COLS}' exceeds Excel's "
                  f"comfortable cell size and may display truncated.")

        row["Rows Merged"] = len(group_idxs)
        out_rows.append(row)

    df_out = pd.DataFrame(out_rows)

    # Column order: follow the INPUT file's own header sequence, with any
    # 'DOCIDs 2/3/4' overflow columns right after 'DOCIDs' and any
    # 'Other Address 2/3/4' overflow columns right after 'Other Address'
    # (regardless of whether the input already had those overflow columns
    # itself - they're always placed via this fixed logic, never via their
    # own raw input position too, so a chained run's input never ends up
    # with the same overflow column listed twice), and columns with no
    # input counterpart ('Other Address' on a fresh source file, 'Rows
    # Merged') at the end.
    docid_extra_cols = [f"{COL_DOCID} {i}" for i in range(2, MAX_DOCID_COLS + 1)
                        if f"{COL_DOCID} {i}" in df_out.columns]
    other_addr_extra_cols = [f"{COL_OTHER_ADDR} {i}" for i in range(2, MAX_OTHER_ADDR_COLS + 1)
                             if f"{COL_OTHER_ADDR} {i}" in df_out.columns]
    docid_extra_names = {f"{COL_DOCID} {i}" for i in range(2, MAX_DOCID_COLS + 1)}
    other_addr_extra_names = {f"{COL_OTHER_ADDR} {i}" for i in range(2, MAX_OTHER_ADDR_COLS + 1)}
    input_order = list(df.columns)
    extra_cols = [c for c in df_out.columns
                  if c not in input_order
                  and c not in docid_extra_names
                  and c not in other_addr_extra_names]
    new_order = []
    for c in input_order:
        if c in docid_extra_names or c in other_addr_extra_names:
            continue
        if c not in df_out.columns:
            continue
        new_order.append(c)
        if c == COL_DOCID:
            new_order.extend(docid_extra_cols)
        if c == COL_OTHER_ADDR:
            new_order.extend(other_addr_extra_cols)
    for c in extra_cols:
        new_order.append(c)
        if c == COL_OTHER_ADDR:
            new_order.extend(other_addr_extra_cols)
    df_out = df_out[new_order]

    df_out = df_out.sort_values(["Rows Merged"], ascending=False).reset_index(drop=True)

    if docid_overflow_groups:
        print(f"  {docid_overflow_groups:,} identity group(s) had a DOCID list too long "
              f"for one cell (> {DOCID_CHUNK_SIZE:,} chars) - split across up to "
              f"{max_docid_cols_used} '{COL_DOCID}' columns.")
    if other_addr_overflow_groups:
        print(f"  {other_addr_overflow_groups:,} identity group(s) had an Other Address "
              f"list too long for one cell (> {OTHER_ADDR_CHUNK_SIZE:,} chars) - split "
              f"across up to {max_other_addr_cols_used} '{COL_OTHER_ADDR}' columns.")

    n_multi = (df_out["Rows Merged"] > 1).sum()
    print(f"  {n_multi:,} identity groups combine 2+ original rows.")
    biggest = df_out["Rows Merged"].max()
    print(f"  Largest merged group: {biggest:,} rows.")

    print(f"Writing {OUTPUT_XLSX} ...")
    _write_workbook(OUTPUT_XLSX, {"Merged Data": df_out})
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
    # Optional command-line overrides, so the input/output file doesn't
    # require editing the CONFIG block every run:
    #   python "260715 pd ds unique id merge.py"                       (uses INPUT_XLSX/OUTPUT_XLSX above as-is)
    #   python "260715 pd ds unique id merge.py" input.xlsx            (overrides INPUT_XLSX only)
    #   python "260715 pd ds unique id merge.py" input.xlsx out.xlsx   (overrides both)
    if len(sys.argv) > 3:
        print(f"Usage: python {sys.argv[0]!r} [input_file] [output_file]", file=sys.stderr)
        sys.exit(1)
    if len(sys.argv) > 1:
        INPUT_XLSX = sys.argv[1]
    if len(sys.argv) > 2:
        OUTPUT_XLSX = sys.argv[2]

    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
