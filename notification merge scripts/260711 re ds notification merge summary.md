# Notification Report — Merge Logic Summary

This document explains, in plain language, the rules for merging person records
to build the notification report. No code yet — this is meant to be reviewed
and confirmed first.

## All matching rules at a glance

| # | Condition | Result |
|---|---|---|
| Base | SSN matches **and** DOB matches (name not considered) | Merge |
| 1 | Same SSN, matching name (exact or typo), DOB differs | Merge; use the majority (most frequent) DOB |
| 2 | Same name, one row has DOB only / other has SSN only, no conflict | Merge (complementary data) |
| 3 | Matching PII (e.g. Driver's License/Passport/other gov ID), one row's name is "[Unknown]" | Merge; real name supersedes "[Unknown]" |
| 4 | Same last name, first name is an initial vs. full (e.g. "H" vs. "Harish"), SSN/PII confirms | Merge; full name kept |
| 5 | Same last name, first name is a partial spelling (e.g. "Did" vs. "Didar"), SSN/DOB match | Merge; full name kept |
| 6 | Same first name, last name is a partial spelling/typo (e.g. "Sin"/"Sing" vs. "Singh"), SSN/DOB match | Merge; other differing PII joined with `;` |
| 7 | Middle name blank, an initial, or misspelled on one row; first/last/SSN/DOB otherwise match | Merge; fullest middle name spelling kept |
| 8 | Suffix conflicts (e.g. Jr vs. Sr, II vs. III), even if SSN/DOB/PII otherwise match | **Do not merge** — flag for manual review |
| 9 | Any suffix (Jr, Sr, II, III, IV, etc.) present on one row, blank on the other, everything else matches | Merge; non-blank suffix kept |
| — | SSN or DOB differs (even if name is identical) | **Do not merge** — different person |
| — | No matching SSN+DOB and none of rules 1–7 apply | **Do not merge** — stays a separate row |

## What counts as "the same person"

A row is treated as the **same person** as another row only when **both**:

- Social Security Number (SSN) matches, **and**
- Date of Birth (DOB) matches

Name is **not** used to decide if it's the same person. So if the SSN and DOB
match but the name is spelled differently (typo, maiden name, nickname), it's
still treated as the same person.

**Example — merged (same person):**

| Name | SSN | DOB |
|---|---|---|
| Jon Smith | 123-45-6789 | 01/02/1980 |
| Jonathan Smith | 123-45-6789 | 01/02/1980 |

Both rows share the same SSN and DOB → treated as one person.

## Additional matching rules

Besides the main SSN + DOB rule above, these situations are also treated as a
match:

**1. Same SSN, matching name (exact or a typo), but DOB is different.**
These are still merged into one record. Since the DOBs disagree, the report
uses whichever DOB value shows up **most often** among the matching rows (the
"majority" DOB) as the merged record's DOB.

| Name | SSN | DOB |
|---|---|---|
| Jon Smith | 123-45-6789 | 01/02/1980 |
| Jonathan Smith | 123-45-6789 | 01/02/1980 |
| Jonathan Smith | 123-45-6789 | 05/09/1979 |

→ Merged into one person, using `01/02/1980` as the DOB (it appears twice,
more than `05/09/1979`).

**2. Same name, but one row only has a DOB and the other row only has an SSN**
(no value in the row to disagree with), and nothing else conflicts. These are
merged, since the two rows simply have complementary information about the
same person rather than contradicting each other.

| Name | SSN | DOB |
|---|---|---|
| Jonathan Smith | 123-45-6789 | *(blank)* |
| Jonathan Smith | *(blank)* | 01/02/1980 |

→ Merged into one person: `Jonathan Smith, SSN 123-45-6789, DOB 01/02/1980`.

**3. Matching PII (e.g. same Driver's License, Passport, or other government
ID), where one row's name is captured as "[Unknown]" and the other row has an
actual name.** These are merged, and the real name takes precedence over
"[Unknown]" (the "[Unknown]" value is dropped, not kept alongside the real
name).

| Name | Driver's License |
|---|---|
| [Unknown] | D1234567 |
| Jonathan Smith | D1234567 |

→ Merged into one person: `Jonathan Smith, Driver's License D1234567`.

**4. First name is an initial on one row and the full first name on the other
(e.g. "H" vs. "Harish"), with the same last name, and SSN or PII confirms it's
the same person.** These are merged — an initial is treated as shorthand for
the full name, not a conflict, as long as the SSN/PII match backs it up. The
full name is kept in the merged record (the initial is dropped).

| First Name | Last Name | SSN |
|---|---|---|
| H | Singh | 123-45-6789 |
| Harish | Singh | 123-45-6789 |

→ Merged into one person: `Harish Singh, SSN 123-45-6789`.

**5. First name is a partial spelling of the full first name (e.g. "Did" vs.
"Didar"), with the same last name and matching SSN/DOB.** These are merged,
same as the initial-vs-full-name rule above — a partial first name is treated
as shorthand, not a conflict, when SSN/DOB backs it up.

| First Name | Last Name | SSN | DOB |
|---|---|---|---|
| Did | Singh | 123-45-6789 | 01/02/1980 |
| Didar | Singh | 123-45-6789 | 01/02/1980 |

→ Merged into one person: `Didar Singh, SSN 123-45-6789, DOB 01/02/1980`.

**6. Last name is a partial spelling or a typo of the full last name (e.g.
"Sin" or "Sing" vs. "Singh"), with the same first name and matching SSN/DOB.**
These are also merged. If any other PII field differs between the two rows
(e.g. a different Driver's License number), that field is combined with a
semicolon as usual, rather than blocking the merge.

| First Name | Last Name | SSN | DOB | Driver's License |
|---|---|---|---|---|
| Didar | Sing | 123-45-6789 | 01/02/1980 | D1234567 |
| Didar | Singh | 123-45-6789 | 01/02/1980 | D7654321 |

→ Merged into one person: `Didar Singh, SSN 123-45-6789, DOB 01/02/1980,
Driver's License D1234567; D7654321`.

**7. Middle name differences never block a merge**, as long as first name,
last name, and SSN/DOB otherwise match (using the shorthand/partial rules
above where needed). This covers all of the following, and any combination of
them:

- Middle name is blank/missing on one row (e.g. "Didar Kumar Singh" vs. "Didar
  Singh").
- Middle name is an initial on one row (e.g. "Kumar" vs. "K").
- Middle name is spelled differently or misspelled on one row (e.g. "Kumar"
  vs. "Kumaar").

| First Name | Middle Name | Last Name | SSN |
|---|---|---|---|
| Didar | Kumar | Singh | 123-45-6789 |
| Didar | *(blank)* | Singh | 123-45-6789 |
| Didar | K | Singh | 123-45-6789 |
| Didar | Kumaar | Singh | 123-45-6789 |

→ All merged into one person: `Didar Kumar Singh, SSN 123-45-6789` (the
fullest/most complete middle name spelling is kept).

**8. Suffix conflicts (e.g. "Jr" vs. "Sr", "II" vs. "III") block the merge —
even if SSN, DOB, and PII otherwise match.** Unlike a middle name difference,
a suffix like Jr/Sr usually identifies two distinct, related people (e.g. a
father and son), not a spelling variation of the same person. So these rows
are kept separate and flagged for manual review rather than merged.

| First Name | Last Name | Suffix | SSN | DOB |
|---|---|---|---|---|
| Didar | Singh | Jr | 123-45-6789 | 01/02/1980 |
| Didar | Singh | Sr | 123-45-6789 | 01/02/1980 |

→ **Not merged** — kept as two separate rows and flagged for review, since a
matching SSN/DOB alongside a Jr/Sr conflict is itself a sign of a
data-quality issue worth investigating.

**9. Suffix is present on one row (any suffix — "Jr", "Sr", "II", "III", "IV",
etc.) and blank/missing on the other row, with everything else matching.**
This is **not** a conflict, regardless of which suffix it is — a blank suffix
just means it wasn't captured, not that it disagrees. These rows are merged,
same as the middle-name-blank case, and the non-blank suffix is kept.

| First Name | Last Name | Suffix | SSN | DOB |
|---|---|---|---|---|
| Didar | Singh | Jr | 123-45-6789 | 01/02/1980 |
| Didar | Singh | *(blank)* | 123-45-6789 | 01/02/1980 |

→ Merged into one person: `Didar Singh Jr, SSN 123-45-6789, DOB 01/02/1980`.

| First Name | Last Name | Suffix | SSN | DOB |
|---|---|---|---|---|
| Didar | Singh | III | 123-45-6789 | 01/02/1980 |
| Didar | Singh | *(blank)* | 123-45-6789 | 01/02/1980 |

→ Merged into one person: `Didar Singh III, SSN 123-45-6789, DOB 01/02/1980`.

Note rule 8 still applies whenever **both** rows have a real, differing
suffix (Jr vs. Sr, II vs. III, Sr vs. IV, etc.) — that combination blocks the
merge. Rule 9 only covers the case where one side is blank.

## What counts as a "different person"

If SSN or DOB is different, it's a different person — **even if the name is
identical**.

**Example — NOT merged (different people):**

| Name | SSN | DOB |
|---|---|---|
| John Smith | 123-45-6789 | 01/02/1980 |
| John Smith | 987-65-4321 | 03/04/1975 |

Same name, but different SSN and DOB → these stay as two separate records.

## What happens when a match is found

When two or more rows are confirmed to be the same person, their data fields
are combined into a single row:

- All of that person's **PII** fields (Driver's License, Passport, other
  government ID) are combined.
- All of that person's **PHI** fields are combined the same way.
- For each field, only the **different** values are kept, joined together with
  a semicolon (`;`). If all rows already have the same value, it just appears
  once — nothing is duplicated.
- If the name is spelled differently across the matched rows, those different
  name values are also kept and joined with semicolons, so a reviewer can see
  the discrepancy rather than the report silently picking one.

**Example — before merge:**

| Name | SSN | DOB | Driver's License |
|---|---|---|---|
| Jon Smith | 123-45-6789 | 01/02/1980 | D1234567 |
| Jonathan Smith | 123-45-6789 | 01/02/1980 | D1234567 |
| Jonathan Smith | 123-45-6789 | 01/02/1980 | D7654321 |

**Example — after merge (one row):**

| Name | SSN | DOB | Driver's License |
|---|---|---|---|
| Jon Smith; Jonathan Smith | 123-45-6789 | 01/02/1980 | D1234567; D7654321 |

Note the Driver's License value `D1234567` only appears once, since it was
repeated.

## What stays separate

Any row that does not have a matching SSN + DOB with another row stays exactly
as it is — it is not combined with anything.

## Handling reminder (PII/PHI)

This process works with SSN, DOB, and other PII/PHI. Per organizational
compliance obligations (ISO 27001, SOC 2, HIPAA):

- Real PII/PHI data should never be pasted into chat tools — only the logic
  and structure should be discussed here.
- The merged output must be saved only to the authorized/secured Global
  Insider folder — never to a desktop or personal drive.
- Access to the merged report should be limited to those authorized to view
  PII/PHI.

## Next steps

Once this logic is confirmed as correct, the next step is to build the actual
merge — either as a SQL query against the database or a script against the
Excel export — using exactly the rules described above.
