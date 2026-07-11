/* ============================================================
   260711 pd ds notification merge.sql

   Merges PII/PHI person records in [cng_db].[dbo].[cng_dedup] into
   one row per confirmed person, for the notification report, per the
   rules documented in:
       260711 re ds notification merge summary.md

   Match key: Social Security Number + Full Date of Birth (Base rule),
   plus 9 additional corroborating rules (same-SSN+name/DOB-differs,
   complementary SSN/DOB, [Unknown]-name override, name shorthand/typo
   variants, middle-name flexibility, suffix-conflict guard, suffix-
   blank exception). See the summary .md for the plain-language version
   of every rule below.

   THIS SCRIPT ONLY READS DATA (SELECT). It does not UPDATE, DELETE, or
   modify [cng_db].[dbo].[cng_dedup] in any way.

   IMPORTANT before trusting the output:
     - Run this against a STAGING/DEV copy of the table first and spot
       check the merged groups before using it for anything official.
     - This output contains SSN, DOB, and other PII/PHI. Save any
       exported result only to the secured/authorized folder for this
       data - never to a desktop or personal drive.
     - Requires SQL Server 2017+ (STRING_AGG ... WITHIN GROUP).

   Known simplifications (call these out to the business owner before
   sign-off):
     - SSN is treated as valid only when it is exactly 9 digits after
       stripping non-digit characters, and not a common junk value
       (all-same-digit, or one of a short placeholder list below).
       Masked/partial SSNs (with X's, asterisks, etc.) are NOT handled
       here and will simply fail to match on SSN.
     - DOB is parsed assuming MM/DD/YYYY; a value that doesn't parse is
       treated as "no DOB" (can't be used to match).
     - "Fullest name kept" and "majority DOB" tie-breaks are heuristic
       (longest string / most frequent value) - review groups with
       more than 2-3 rows before trusting the chosen value blindly.
   ============================================================ */

SET NOCOUNT ON;

-------------------------------------------------------------------
-- 0) Tunable placeholder lists (edit these if you find more junk
--    values in the real data)
-------------------------------------------------------------------
IF OBJECT_ID('tempdb..#name_placeholders') IS NOT NULL DROP TABLE #name_placeholders;
CREATE TABLE #name_placeholders (v VARCHAR(30) PRIMARY KEY);
INSERT INTO #name_placeholders (v) VALUES
    ('UNKNOWN'),('UNK'),('UNKN'),('NA'),('NONE'),('NULL'),('NIL'),
    ('TEST'),('XXX'),('XX'),('X'),('NMN'),('NONAME'),('NOTGIVEN'),('NOTPROVIDED');

IF OBJECT_ID('tempdb..#ssn_placeholders') IS NOT NULL DROP TABLE #ssn_placeholders;
CREATE TABLE #ssn_placeholders (v CHAR(9) PRIMARY KEY);
INSERT INTO #ssn_placeholders (v) VALUES
    ('123456789'),('987654321'),('111223333'),('123121234'),
    ('456789123'),('078051120'),('219099999'),('457555462');

-------------------------------------------------------------------
-- 1) Load + normalize source rows into a working copy
-------------------------------------------------------------------
IF OBJECT_ID('tempdb..#base') IS NOT NULL DROP TABLE #base;

SELECT
    s.[DOCIDs],
    s.[First Name], s.[Last Name], s.[Middle Name], s.[Suffix],
    s.[Data Subject Type], s.[Birth Information],
    s.[Full Date of Birth (MM/DD/YYYY)],
    s.[Residential Address], s.[City],
    s.[State of Residence (if US)], s.[Province of Residence (if Canada)],
    s.[Zip Code], s.[Country of Residence], s.[Address Comments],
    s.[Email Address - Personal], s.[Phone Number], s.[Contact Information],
    s.[Social Security Number],
    s.[Driver's License Number], s.[DL Issuing Country],
    s.[DL Issuing Province (if Canada)], s.[DL Issuing State (if US)],
    s.[Passport Country], s.[Passport Number],
    s.[Government ID Issuing Country], s.[Government- Issued Identification],
    s.[Government-Issued ID Number], s.[Health Related Information],
    s.[Employee Identification Number], s.[Work-Related Information],
    s.[Family Information], s.[Financial Account Information],
    s.[Student-Related Information], s.[Demographic Information],
    s.[Biometric Data], s.[PI Notes],
    s.[Access Credentials (Non-Financial Account)],

    -- normalized fields used only for matching / clustering, filled in below
    CAST(NULL AS CHAR(9))     AS n_ssn,
    CAST(NULL AS DATE)        AS n_dob,
    CAST('' AS NVARCHAR(100)) AS n_first,
    CAST('' AS NVARCHAR(100)) AS n_last,
    CAST('' AS NVARCHAR(100)) AS n_middle,
    CAST('' AS NVARCHAR(20))  AS n_suffix
INTO #base
FROM [cng_db].[dbo].[cng_dedup] s;

ALTER TABLE #base ADD row_id INT IDENTITY(1,1);
ALTER TABLE #base ADD grp INT;
UPDATE #base SET grp = row_id;   -- every row starts as its own group

-- SSN: keep only if exactly 9 digits after stripping punctuation, and not a
-- known junk/placeholder value or an all-same-digit fake (000000000, ...)
UPDATE b
SET n_ssn = CASE
    WHEN LEN(d.digits_only) = 9
         AND NOT EXISTS (SELECT 1 FROM #ssn_placeholders p WHERE p.v = d.digits_only)
         AND d.digits_only <> REPLICATE(LEFT(d.digits_only, 1), 9)
    THEN d.digits_only ELSE NULL END
FROM #base b
CROSS APPLY (
    SELECT REPLACE(REPLACE(REPLACE(LTRIM(RTRIM(ISNULL(b.[Social Security Number], ''))), '-', ''), ' ', ''), '.', '') AS digits_only
) d;

-- DOB: parse assuming MM/DD/YYYY (style 101); NULL if it doesn't parse
UPDATE #base
SET n_dob = COALESCE(
        TRY_CONVERT(date, [Full Date of Birth (MM/DD/YYYY)], 101),
        TRY_CONVERT(date, [Full Date of Birth (MM/DD/YYYY)])
    );

-- Names: upper/trim; placeholder values ("[Unknown]", "N/A", ...) become ''
-- after stripping brackets/parens/periods, so a placeholder name never wins
-- against - or falsely conflicts with - a real name (Rule 3).
UPDATE b
SET n_first = CASE WHEN EXISTS (SELECT 1 FROM #name_placeholders p WHERE p.v = fc.core) THEN '' ELSE f.norm END,
    n_last  = CASE WHEN EXISTS (SELECT 1 FROM #name_placeholders p WHERE p.v = lc.core) THEN '' ELSE l.norm END,
    n_middle= CASE WHEN EXISTS (SELECT 1 FROM #name_placeholders p WHERE p.v = mc.core) THEN '' ELSE m.norm END,
    n_suffix= UPPER(REPLACE(LTRIM(RTRIM(ISNULL(b.[Suffix], ''))), '.', ''))
FROM #base b
CROSS APPLY (SELECT UPPER(LTRIM(RTRIM(ISNULL(b.[First Name], ''))))  AS norm) f
CROSS APPLY (SELECT REPLACE(REPLACE(f.norm, '[', ''), ']', '')       AS core) fc
CROSS APPLY (SELECT UPPER(LTRIM(RTRIM(ISNULL(b.[Last Name], ''))))   AS norm) l
CROSS APPLY (SELECT REPLACE(REPLACE(l.norm, '[', ''), ']', '')       AS core) lc
CROSS APPLY (SELECT UPPER(LTRIM(RTRIM(ISNULL(b.[Middle Name], '')))) AS norm) m
CROSS APPLY (SELECT REPLACE(REPLACE(m.norm, '[', ''), ']', '')       AS core) mc;

-------------------------------------------------------------------
-- 2) Build a symmetric edge list: which rows are the same person?
--    Every branch below is guarded by the Rule 8 suffix-conflict check
--    up front, since a real conflicting suffix (Jr vs Sr, II vs III)
--    blocks EVERY other rule, even the base SSN+DOB match.
-------------------------------------------------------------------
IF OBJECT_ID('tempdb..#edges') IS NOT NULL DROP TABLE #edges;

SELECT a.row_id AS row_a, b.row_id AS row_b
INTO #edges
FROM #base a
JOIN #base b ON b.row_id > a.row_id
WHERE
    -- Rule 8: a real, differing suffix on both sides blocks any merge
    NOT (a.n_suffix <> '' AND b.n_suffix <> '' AND a.n_suffix <> b.n_suffix)
    -- Two DIFFERENT real SSNs must never be merged, even when a fuzzy name +
    -- matching DOB (Rules 4-7) would otherwise corroborate the pair. A blank
    -- SSN on either side is not a conflict (Rule 2 still applies).
    AND NOT (a.n_ssn IS NOT NULL AND b.n_ssn IS NOT NULL AND a.n_ssn <> b.n_ssn)
    -- Two genuinely different names (both sides have a real first+last name,
    -- and they don't even loosely match - not a typo/prefix/initial) block
    -- EVERY rule, including the Base rule. A matching SSN+DOB alongside a
    -- real name conflict usually means a fake/reused/incorrect SSN, not
    -- confirmation of the same person.
    AND NOT (
        a.n_first <> '' AND a.n_last <> '' AND b.n_first <> '' AND b.n_last <> ''
        AND NOT (
            (a.n_last = b.n_last
                 OR (LEN(a.n_last) >= 3 AND b.n_last LIKE a.n_last + '%')
                 OR (LEN(b.n_last) >= 3 AND a.n_last LIKE b.n_last + '%'))
            AND (a.n_first = b.n_first
                 OR (LEN(a.n_first) >= 1 AND b.n_first LIKE a.n_first + '%')
                 OR (LEN(b.n_first) >= 1 AND a.n_first LIKE b.n_first + '%'))
            AND (a.n_middle = b.n_middle OR a.n_middle = '' OR b.n_middle = ''
                 OR b.n_middle LIKE a.n_middle + '%' OR a.n_middle LIKE b.n_middle + '%')
        )
    )
AND (
    -- Base rule: SSN + DOB both present and match exactly (name irrelevant)
    (a.n_ssn IS NOT NULL AND a.n_ssn = b.n_ssn AND a.n_dob IS NOT NULL AND a.n_dob = b.n_dob)

    -- Rule 1: same SSN + matching/typo name, DOB differs -> majority DOB wins later
    OR (a.n_ssn IS NOT NULL AND a.n_ssn = b.n_ssn
        AND a.n_last <> '' AND b.n_last <> ''
        AND (a.n_last = b.n_last
             OR (LEN(a.n_last) >= 3 AND b.n_last LIKE a.n_last + '%')
             OR (LEN(b.n_last) >= 3 AND a.n_last LIKE b.n_last + '%'))
        AND (a.n_first = b.n_first
             OR (LEN(a.n_first) >= 1 AND b.n_first LIKE a.n_first + '%')
             OR (LEN(b.n_first) >= 1 AND a.n_first LIKE b.n_first + '%'))
       )

    -- Rule 2: same name, complementary data (one row SSN-only, other DOB-only)
    OR (a.n_last <> '' AND a.n_last = b.n_last AND a.n_first = b.n_first
        AND ((a.n_ssn IS NOT NULL AND b.n_ssn IS NULL) OR (a.n_ssn IS NULL AND b.n_ssn IS NOT NULL))
        AND ((a.n_dob IS NOT NULL AND b.n_dob IS NULL) OR (a.n_dob IS NULL AND b.n_dob IS NOT NULL))
       )

    -- Rule 3: matching PII (Driver's License / Passport / Government ID
    -- Number), and one side's name is entirely blank/"[Unknown]"
    OR (
        (
            (NULLIF(LTRIM(RTRIM(a.[Driver's License Number])), '') IS NOT NULL
             AND UPPER(LTRIM(RTRIM(a.[Driver's License Number]))) = UPPER(LTRIM(RTRIM(b.[Driver's License Number]))))
         OR (NULLIF(LTRIM(RTRIM(a.[Passport Number])), '') IS NOT NULL
             AND UPPER(LTRIM(RTRIM(a.[Passport Number]))) = UPPER(LTRIM(RTRIM(b.[Passport Number]))))
         OR (NULLIF(LTRIM(RTRIM(a.[Government-Issued ID Number])), '') IS NOT NULL
             AND UPPER(LTRIM(RTRIM(a.[Government-Issued ID Number]))) = UPPER(LTRIM(RTRIM(b.[Government-Issued ID Number]))))
        )
        AND (
            (a.n_first = '' AND a.n_last = '' AND (b.n_first <> '' OR b.n_last <> ''))
            OR (b.n_first = '' AND b.n_last = '' AND (a.n_first <> '' OR a.n_last <> ''))
        )
       )

    -- Rules 4-7: fuzzy name (initial / partial-spelling first name,
    -- partial-spelling/typo last name, flexible middle name), corroborated
    -- by a matching SSN or a matching DOB
    OR (
        a.n_last <> '' AND b.n_last <> ''
        AND (a.n_last = b.n_last
             OR (LEN(a.n_last) >= 3 AND b.n_last LIKE a.n_last + '%')
             OR (LEN(b.n_last) >= 3 AND a.n_last LIKE b.n_last + '%'))
        AND (a.n_first = b.n_first
             OR (LEN(a.n_first) >= 1 AND b.n_first LIKE a.n_first + '%')
             OR (LEN(b.n_first) >= 1 AND a.n_first LIKE b.n_first + '%'))
        AND (a.n_middle = b.n_middle OR a.n_middle = '' OR b.n_middle = ''
             OR b.n_middle LIKE a.n_middle + '%' OR a.n_middle LIKE b.n_middle + '%')
        AND ((a.n_ssn IS NOT NULL AND a.n_ssn = b.n_ssn) OR (a.n_dob IS NOT NULL AND a.n_dob = b.n_dob))
       )

    -- Rule 9 is implicit: a blank suffix never triggers the Rule 8 guard
    -- above, so it never blocks any of the rules it would otherwise satisfy
);

-- make the edge list symmetric so the propagation step below can walk it
-- in either direction
INSERT INTO #edges (row_a, row_b)
SELECT row_b, row_a FROM #edges;

-------------------------------------------------------------------
-- 3) Cluster connected rows into groups. Transitive: if A matches B and
--    B matches C, all three land in the same group even if A and C don't
--    directly satisfy any rule with each other. Implemented by repeatedly
--    propagating the smallest row_id across every edge until nothing
--    changes (standard "connected components" approach).
-------------------------------------------------------------------
DECLARE @changed INT = 1;
WHILE @changed > 0
BEGIN
    UPDATE b
    SET b.grp = e.min_grp
    FROM #base b
    JOIN (
        SELECT e.row_a AS row_id, MIN(g.grp) AS min_grp
        FROM #edges e
        JOIN #base g ON g.row_id = e.row_b
        GROUP BY e.row_a
    ) e ON e.row_id = b.row_id
    WHERE e.min_grp < b.grp;

    SET @changed = @@ROWCOUNT;
END

-------------------------------------------------------------------
-- 4) Build the merged output: one row per group
-------------------------------------------------------------------
IF OBJECT_ID('tempdb..#groups') IS NOT NULL DROP TABLE #groups;
SELECT DISTINCT grp INTO #groups FROM #base;

SELECT
    g.grp AS [Merge Group],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[DOCIDs])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[DOCIDs])), '') IS NOT NULL) d
    ) AS [DOCIDs],

    -- Names: keep the longest non-blank / non-placeholder value seen in the group
    (SELECT TOP 1 LTRIM(RTRIM(b.[First Name])) FROM #base b
     WHERE b.grp = g.grp AND b.n_first <> ''
     ORDER BY LEN(LTRIM(RTRIM(b.[First Name]))) DESC) AS [First Name],

    (SELECT TOP 1 LTRIM(RTRIM(b.[Last Name])) FROM #base b
     WHERE b.grp = g.grp AND b.n_last <> ''
     ORDER BY LEN(LTRIM(RTRIM(b.[Last Name]))) DESC) AS [Last Name],

    (SELECT TOP 1 LTRIM(RTRIM(b.[Middle Name])) FROM #base b
     WHERE b.grp = g.grp AND b.n_middle <> ''
     ORDER BY LEN(LTRIM(RTRIM(b.[Middle Name]))) DESC) AS [Middle Name],

    (SELECT TOP 1 LTRIM(RTRIM(b.[Suffix])) FROM #base b
     WHERE b.grp = g.grp AND b.n_suffix <> ''
     ORDER BY LEN(LTRIM(RTRIM(b.[Suffix]))) DESC) AS [Suffix],

    -- DOB: majority value wins (Rule 1); ties broken arbitrarily (most recent)
    (SELECT TOP 1 b.n_dob FROM #base b
     WHERE b.grp = g.grp AND b.n_dob IS NOT NULL
     GROUP BY b.n_dob
     ORDER BY COUNT(*) DESC, b.n_dob DESC) AS [Date of Birth],

    -- A single value, not semicolon-joined: the SSN-conflict guard above
    -- guarantees every real SSN in this group is identical.
    (SELECT TOP 1 LTRIM(RTRIM(b.[Social Security Number])) FROM #base b
     WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Social Security Number])), '') IS NOT NULL
     ORDER BY LEN(LTRIM(RTRIM(b.[Social Security Number]))) DESC) AS [Social Security Number],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Data Subject Type])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Data Subject Type])), '') IS NOT NULL) d
    ) AS [Data Subject Type],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Birth Information])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Birth Information])), '') IS NOT NULL) d
    ) AS [Birth Information],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Residential Address])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Residential Address])), '') IS NOT NULL) d
    ) AS [Residential Address],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[City])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[City])), '') IS NOT NULL) d
    ) AS [City],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[State of Residence (if US)])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[State of Residence (if US)])), '') IS NOT NULL) d
    ) AS [State of Residence (if US)],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Province of Residence (if Canada)])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Province of Residence (if Canada)])), '') IS NOT NULL) d
    ) AS [Province of Residence (if Canada)],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Zip Code])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Zip Code])), '') IS NOT NULL) d
    ) AS [Zip Code],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Country of Residence])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Country of Residence])), '') IS NOT NULL) d
    ) AS [Country of Residence],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Address Comments])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Address Comments])), '') IS NOT NULL) d
    ) AS [Address Comments],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Email Address - Personal])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Email Address - Personal])), '') IS NOT NULL) d
    ) AS [Email Address - Personal],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Phone Number])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Phone Number])), '') IS NOT NULL) d
    ) AS [Phone Number],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Contact Information])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Contact Information])), '') IS NOT NULL) d
    ) AS [Contact Information],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Driver's License Number])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Driver's License Number])), '') IS NOT NULL) d
    ) AS [Driver's License Number],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[DL Issuing Country])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[DL Issuing Country])), '') IS NOT NULL) d
    ) AS [DL Issuing Country],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[DL Issuing Province (if Canada)])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[DL Issuing Province (if Canada)])), '') IS NOT NULL) d
    ) AS [DL Issuing Province (if Canada)],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[DL Issuing State (if US)])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[DL Issuing State (if US)])), '') IS NOT NULL) d
    ) AS [DL Issuing State (if US)],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Passport Country])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Passport Country])), '') IS NOT NULL) d
    ) AS [Passport Country],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Passport Number])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Passport Number])), '') IS NOT NULL) d
    ) AS [Passport Number],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Government ID Issuing Country])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Government ID Issuing Country])), '') IS NOT NULL) d
    ) AS [Government ID Issuing Country],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Government- Issued Identification])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Government- Issued Identification])), '') IS NOT NULL) d
    ) AS [Government- Issued Identification],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Government-Issued ID Number])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Government-Issued ID Number])), '') IS NOT NULL) d
    ) AS [Government-Issued ID Number],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Health Related Information])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Health Related Information])), '') IS NOT NULL) d
    ) AS [Health Related Information],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Employee Identification Number])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Employee Identification Number])), '') IS NOT NULL) d
    ) AS [Employee Identification Number],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Work-Related Information])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Work-Related Information])), '') IS NOT NULL) d
    ) AS [Work-Related Information],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Family Information])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Family Information])), '') IS NOT NULL) d
    ) AS [Family Information],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Financial Account Information])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Financial Account Information])), '') IS NOT NULL) d
    ) AS [Financial Account Information],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Student-Related Information])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Student-Related Information])), '') IS NOT NULL) d
    ) AS [Student-Related Information],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Demographic Information])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Demographic Information])), '') IS NOT NULL) d
    ) AS [Demographic Information],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Biometric Data])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Biometric Data])), '') IS NOT NULL) d
    ) AS [Biometric Data],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[PI Notes])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[PI Notes])), '') IS NOT NULL) d
    ) AS [PI Notes],

    (SELECT STRING_AGG(v, '; ') WITHIN GROUP (ORDER BY v)
     FROM (SELECT DISTINCT LTRIM(RTRIM(b.[Access Credentials (Non-Financial Account)])) AS v FROM #base b
           WHERE b.grp = g.grp AND NULLIF(LTRIM(RTRIM(b.[Access Credentials (Non-Financial Account)])), '') IS NOT NULL) d
    ) AS [Access Credentials (Non-Financial Account)],

    (SELECT COUNT(*) FROM #base b WHERE b.grp = g.grp) AS [Rows Merged]

FROM #groups g
ORDER BY [Rows Merged] DESC, g.grp;

-------------------------------------------------------------------
-- 5) Rule 8 review list: row pairs that WOULD have matched (base rule,
--    or fuzzy-name + SSN/DOB corroboration) except that a real,
--    differing suffix (Jr vs Sr, II vs III, etc.) blocked the merge.
--    These are NOT merged above - review them manually.
-------------------------------------------------------------------
SELECT
    a.[DOCIDs] AS [DOCID A], a.[First Name] AS [First Name A], a.[Last Name] AS [Last Name A], a.[Suffix] AS [Suffix A],
    b.[DOCIDs] AS [DOCID B], b.[First Name] AS [First Name B], b.[Last Name] AS [Last Name B], b.[Suffix] AS [Suffix B]
FROM #base a
JOIN #base b ON b.row_id > a.row_id
WHERE a.n_suffix <> '' AND b.n_suffix <> '' AND a.n_suffix <> b.n_suffix
AND (
    (a.n_ssn IS NOT NULL AND a.n_ssn = b.n_ssn AND a.n_dob IS NOT NULL AND a.n_dob = b.n_dob)
    OR (a.n_last <> '' AND b.n_last <> '' AND a.n_last = b.n_last
        AND (a.n_first = b.n_first OR b.n_first LIKE a.n_first + '%' OR a.n_first LIKE b.n_first + '%')
        AND ((a.n_ssn IS NOT NULL AND a.n_ssn = b.n_ssn) OR (a.n_dob IS NOT NULL AND a.n_dob = b.n_dob))
       )
);

-------------------------------------------------------------------
-- 6) Name-conflict review list: row pairs that share an SSN + DOB match
--    (the Base rule) but have two genuinely different, non-typo names.
--    This usually signals a fake/reused/incorrect SSN rather than the
--    same person - NOT merged above; review manually.
-------------------------------------------------------------------
SELECT
    a.[DOCIDs] AS [DOCID A], a.[First Name] AS [First Name A], a.[Last Name] AS [Last Name A], a.[Social Security Number] AS [SSN A],
    b.[DOCIDs] AS [DOCID B], b.[First Name] AS [First Name B], b.[Last Name] AS [Last Name B], b.[Social Security Number] AS [SSN B]
FROM #base a
JOIN #base b ON b.row_id > a.row_id
WHERE a.n_ssn IS NOT NULL AND a.n_ssn = b.n_ssn
  AND a.n_dob IS NOT NULL AND a.n_dob = b.n_dob
  AND a.n_first <> '' AND a.n_last <> '' AND b.n_first <> '' AND b.n_last <> ''
  AND NOT (
      (a.n_last = b.n_last
           OR (LEN(a.n_last) >= 3 AND b.n_last LIKE a.n_last + '%')
           OR (LEN(b.n_last) >= 3 AND a.n_last LIKE b.n_last + '%'))
      AND (a.n_first = b.n_first
           OR (LEN(a.n_first) >= 1 AND b.n_first LIKE a.n_first + '%')
           OR (LEN(b.n_first) >= 1 AND a.n_first LIKE b.n_first + '%'))
      AND (a.n_middle = b.n_middle OR a.n_middle = '' OR b.n_middle = ''
           OR b.n_middle LIKE a.n_middle + '%' OR a.n_middle LIKE b.n_middle + '%')
  );
