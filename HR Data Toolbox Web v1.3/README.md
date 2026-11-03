# HR Data Toolbox Web v1.3

A local Streamlit app for enriching an Excel or CSV working file from one or more Excel/CSV lookup files. Output is always a new `.xlsx` workbook.

## Start

1. Extract the ZIP.
2. Run `install.bat` once.
3. Run `run.bat`.
4. Use the browser page that opens. Close the command window to stop the app.

## Version 1.3

- Each lookup file can use a different identifier from the main working file
- Example: lookup File 1 by EEID and lookup File 2 by Location in the same run
- Composite identifiers remain available independently for any lookup
- Per-lookup match method:
  - Exact only
  - Exact plus standardized
  - Exact plus standardized plus soft contains
- Soft contains supports cases such as `012345` matching `012345 - PC RMH`
- A separate match-status column is added for every lookup file for auditability
- Soft matches across multiple lookup keys are visibly flagged and their selected values follow the duplicate handling policy
- Blank single or composite identifiers are safely treated as unmatched

## Existing features

- Any number of Excel or CSV lookup files
- Duplicate aggregation with unique values by default
- Comma, semicolon, space, new-line, and custom delimiters
- Per-column output naming and conflict handling
- Optional statistics and data-quality sheets, off by default
- Excel-only output
- Local processing with no database

## Important

Soft matching is optional and disabled by default. It is less strict than exact matching, so review rows marked `Soft match` or `Soft match (n keys)` in the output. The app never modifies source files.
