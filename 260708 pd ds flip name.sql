/* ============================================================
   FLIP NAME  -  reversed / re-ordered name duplicate detector
   ------------------------------------------------------------
   Read-only SELECT. Source table is never modified.

   Columns produced (in this order):
     1. [Unique_ID]        - already in the table
     2. [DocID]            - already in the table
     3. [Full Name Count]  - # rows sharing the same Full Name
     4. [Flip Name Count]  - # rows sharing the same Flip Name
     5. [Full Name]        - First Name + ' ' + Last Name
     6. [Flip Name]        - ALL words of Full Name sorted A->Z
     7. ...all columns of the table, as-is

   Flip Name logic:
     Full Name  'Singh Kumar Didar'
     split into words -> Singh, Kumar, Didar
     sort A->Z        -> Didar, Kumar, Singh
     Flip Name  = 'Didar Kumar Singh'
   So records with the SAME set of name words (regardless of how
   they were split into First/Last) share one Flip Name.

   Output: only rows where [Flip Name Count] > 1  (i.e. a possible
   flipped / duplicate name exists).

   Requires SQL Server 2016+ (STRING_SPLIT) and 2017+ (STRING_AGG).
   SSMS 22 compatible.

   HOW TO USE: replace [YourDatabase].[dbo].[YourTable] and the
   [First Name] / [Last Name] column names if yours differ.
   ============================================================ */

SELECT *
FROM
(
    SELECT
        t.[Unique_ID],
        t.[DocID],
        COUNT(*) OVER (PARTITION BY fn.FullName) AS [Full Name Count],
        COUNT(*) OVER (PARTITION BY fx.FlipName) AS [Flip Name Count],
        fn.FullName AS [Full Name],
        fx.FlipName AS [Flip Name],
        t.*                                    -- all columns as-is
    FROM [YourDatabase].[dbo].[YourTable] AS t

    /* Build a clean Full Name: trim each part, normalise tab /
       non-breaking space to a normal space. */
    CROSS APPLY (
        SELECT LTRIM(RTRIM(
                 REPLACE(REPLACE(
                     LTRIM(RTRIM(ISNULL(t.[First Name], ''))) + N' ' +
                     LTRIM(RTRIM(ISNULL(t.[Last Name],  ''))),
                 CHAR(160), N' '), CHAR(9), N' ')
               )) AS FullName
    ) AS fn

    /* Flip Name: split Full Name into words, drop empties from
       double spaces, sort alphabetically, rejoin with a space. */
    CROSS APPLY (
        SELECT STRING_AGG(s.[value], N' ')
               WITHIN GROUP (ORDER BY s.[value]) AS FlipName
        FROM STRING_SPLIT(fn.FullName, N' ') AS s
        WHERE s.[value] <> N''
    ) AS fx
) AS q
WHERE q.[Flip Name Count] > 1              -- only possible duplicates
ORDER BY q.[Flip Name], q.[Full Name], q.[Unique_ID];


/* ------------------------------------------------------------
   VARIATIONS
   ------------------------------------------------------------
   * Also see single (non-duplicate) names: remove the WHERE line.

   * Only rows where the name was actually FLIPPED (Full Name
     differs but Flip Name matches) - add to the WHERE:
         AND q.[Full Name Count] < q.[Flip Name Count]

   * Case-insensitive matching is driven by the column collation
     (default). For strict case handling add COLLATE as needed.

   * 20M rows: the two window COUNTs sort the whole set. Ensure
     tempdb has room; export via SSMS "Save Results As".
   ------------------------------------------------------------ */
