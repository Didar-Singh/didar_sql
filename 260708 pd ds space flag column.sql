/* ============================================================
   EXTRACTION (read-only) - flag "not clean" values
   AUTO-GENERATES a [Space] column for ALL text columns
   ------------------------------------------------------------
   For each row, [Space] lists the HEADER NAME(S) whose value
   is not trimmed/clean, joined with '; '.
     e.g.  FirstName + Address dirty  ->  'FirstName; Address'
           all clean                  ->  '' (empty string)

   You do NOT list columns by hand - the script reads them from
   the catalog, so it scales to 50+ columns automatically.
   Source table is never modified.

   Requires SQL Server 2017+ (STRING_AGG / CONCAT_WS). SSMS 22 OK.

   HOW TO USE:
     1. Set @schema and @table below.
     2. Run.  Review the generated SQL in the Messages tab.
     3. Un-comment the EXEC line to actually run the extract.
   ============================================================ */

DECLARE @schema sysname = N'dbo';          -- <-- your schema
DECLARE @table  sysname = N'YourTable';    -- <-- your table

DECLARE @full   NVARCHAR(300) = QUOTENAME(@schema) + N'.' + QUOTENAME(@table);
DECLARE @cases  NVARCHAR(MAX);
DECLARE @sql    NVARCHAR(MAX);

/* Build one CASE ... THEN 'ColumnName' END per CHARACTER column.
   collation_id IS NOT NULL  ==> char/varchar/nchar/nvarchar/text/ntext.
   (Numeric/date columns can't hold spaces, so they are skipped.
    To check EVERY column regardless of type, delete the AND line
    marked below and every reference will be CAST implicitly.)      */
SELECT @cases = STRING_AGG(
      N'    CASE WHEN ' + QUOTENAME(c.name) + N' IS NOT NULL AND ('
    + QUOTENAME(c.name) + N' <> LTRIM(RTRIM(' + QUOTENAME(c.name) + N'))'
    + N' OR ' + QUOTENAME(c.name) + N' LIKE ''%  %'''                 -- double space
    + N' OR ' + QUOTENAME(c.name) + N' LIKE ''%'' + CHAR(9)   + ''%'''-- tab
    + N' OR ' + QUOTENAME(c.name) + N' LIKE ''%'' + CHAR(13)  + ''%'''-- CR
    + N' OR ' + QUOTENAME(c.name) + N' LIKE ''%'' + CHAR(10)  + ''%'''-- LF
    + N' OR ' + QUOTENAME(c.name) + N' LIKE ''%'' + CHAR(160) + ''%'')'-- nbsp
    + N' THEN ''' + REPLACE(c.name, '''', '''''') + N''' END'
    , N',' + CHAR(13) + CHAR(10))
    WITHIN GROUP (ORDER BY c.column_id)
FROM sys.columns AS c
WHERE c.object_id = OBJECT_ID(@full)
  AND c.collation_id IS NOT NULL;          -- <-- delete this line to check ALL columns

/* Assemble the final read-only extraction query.
   Wrapped in a derived table so we can filter to ONLY the rows
   that have a trim/space/clean issue ([Space] <> ''). */
SET @sql =
      N'SELECT *' + CHAR(13) + CHAR(10)
    + N'FROM (' + CHAR(13) + CHAR(10)
    + N'    SELECT *,' + CHAR(13) + CHAR(10)
    + N'        CONCAT_WS(''; '',' + CHAR(13) + CHAR(10)
    + @cases + CHAR(13) + CHAR(10)
    + N'        ) AS [Space]' + CHAR(13) + CHAR(10)
    + N'    FROM ' + @full + CHAR(13) + CHAR(10)
    + N') AS x' + CHAR(13) + CHAR(10)
    + N'WHERE x.[Space] <> '''';';   -- only rows with an issue

/* 1) REVIEW the generated SQL (Messages tab).
      Note: PRINT truncates ~4000 chars; use the SELECT below to
      see the whole thing if it is long. */
PRINT @sql;
-- SELECT @sql AS generated_sql;   -- click the cell to view full text

/* 2) RUN the extraction (un-comment when ready) */
-- EXEC sys.sp_executesql @sql;


/* ------------------------------------------------------------
   VARIATIONS
   ------------------------------------------------------------
   * Only rows that need cleaning:
       change the SELECT header the script builds to wrap in a
       CTE, or after running once add:  WHERE [Space] <> ''

   * Show clean rows as NULL instead of '':
       replace  CONCAT_WS('; ',   with   NULLIF(CONCAT_WS('; ',
       and add a matching  , '')  before  AS [Space]

   * Extract into a NEW table (source stays untouched):
       inject   INTO [dbo].[YourTable_Extract]   before 'FROM' in
       the @sql assembly above.

   20M rows: this scans the whole table once. Export results via
   SSMS "Save Results As", or use the INTO option to persist.
   ------------------------------------------------------------ */
