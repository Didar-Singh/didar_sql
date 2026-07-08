/* ============================================================
   FLIP NAME  -  reversed / re-ordered name duplicate detector
   ------------------------------------------------------------
   Read-only SELECT. Source table is never modified.

   Output columns (in this order):
     1. [<UniqueIdColumn>] - the unique id already in the table
     2. [DOCID]            - already in the table
     3. [Full Name Count]  - # rows sharing the same Full Name
     4. [Flip Name Count]  - # rows sharing the same Flip Name
     5. [Full Name]        - First Name + ' ' + Last Name
     6. [Flip Name]        - ALL words of Full Name sorted A->Z
     7. ...all columns of the table, as-is

   Flip Name logic:
     Full Name 'Singh Kumar Didar' -> words Singh,Kumar,Didar
     sorted A->Z -> 'Didar Kumar Singh'
   So records with the SAME set of name words (however they were
   split into First/Last) share one Flip Name.

   Output: only rows where [Flip Name Count] > 1.

   Requires SQL Server 2016+ (STRING_SPLIT) and 2017+ (STRING_AGG).

   >>> BEFORE RUNNING <<<
     - Replace [YourDatabase].[dbo].[YourTable].
     - Replace the TWO occurrences of  [UNIQUE_ID_COL]  with your
       real unique-id column name (see helper query at the bottom
       to find it).
     - Adjust [First Name] / [Last Name] / [DOCID] if they differ.
   ============================================================ */

SELECT
    q.[UNIQUE_ID_COL],          -- <-- 1) your unique id column
    q.[DOCID],                  -- 2)
    q.[Full Name Count],        -- 3)
    q.[Flip Name Count],        -- 4)
    q.[Full Name],              -- 5)
    q.[Flip Name],              -- 6)
    q.*                         -- 7) all columns of the table, as-is
FROM
(
    SELECT
        t.*,                                   -- all base columns ONCE
        COUNT(*) OVER (PARTITION BY fn.FullName) AS [Full Name Count],
        COUNT(*) OVER (PARTITION BY fx.FlipName) AS [Flip Name Count],
        fn.FullName AS [Full Name],
        fx.FlipName AS [Flip Name]
    FROM [YourDatabase].[dbo].[YourTable] AS t

    /* Clean Full Name: trim parts, normalise tab / nbsp to space */
    CROSS APPLY (
        SELECT LTRIM(RTRIM(
                 REPLACE(REPLACE(
                     LTRIM(RTRIM(ISNULL(t.[First Name], ''))) + N' ' +
                     LTRIM(RTRIM(ISNULL(t.[Last Name],  ''))),
                 CHAR(160), N' '), CHAR(9), N' ')
               )) AS FullName
    ) AS fn

    /* Flip Name: split into words, drop empties, sort A->Z, rejoin */
    CROSS APPLY (
        SELECT STRING_AGG(s.[value], N' ')
               WITHIN GROUP (ORDER BY s.[value]) AS FlipName
        FROM STRING_SPLIT(fn.FullName, N' ') AS s
        WHERE s.[value] <> N''
    ) AS fx
) AS q
WHERE q.[Flip Name Count] > 1              -- only possible duplicates
ORDER BY q.[Flip Name], q.[Full Name], q.[UNIQUE_ID_COL];


/* ------------------------------------------------------------
   HELPER: find the exact column names in your table
   (run this alone to see what the unique id column is called)
   ------------------------------------------------------------ */
-- SELECT COLUMN_NAME, DATA_TYPE, ORDINAL_POSITION
-- FROM INFORMATION_SCHEMA.COLUMNS
-- WHERE TABLE_NAME = 'YourTable'      -- and TABLE_SCHEMA = 'dbo'
-- ORDER BY ORDINAL_POSITION;


/* ------------------------------------------------------------
   VARIATIONS
   ------------------------------------------------------------
   * See all names (incl. singles): remove the WHERE line.
   * Only names actually FLIPPED (Full Name differs but words match):
        AND q.[Full Name Count] < q.[Flip Name Count]
   * 20M rows: two window sorts - ensure tempdb has room; export
     via SSMS "Save Results As".
   ------------------------------------------------------------ */
