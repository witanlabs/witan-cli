# Changelog

## Unreleased

- New: --create flag in `witan xlsx exec` enables creating and populating a workbook in a single command
- New: --locale flag in `witan xlsx exec` enables controlling which locale is used for formula calculation, number formatting, string comparison, etc. It accepts values such as en-US, which can also be passed as WITAN_LOCALE env var
- New: --version flag prints the API version in addition to CLI version
- New: CELL function is now implemented in calculation engine
- Updated: Add 3 unit aliases missing in CONVERT function: d, s, L
- Updated: Rounding functions now coerce empty reference args as 0, matching Excel
- Updated: COUNTIF/etc criteria functions now reject criteria strings over 255 characters, matching Excel
- Updated: COUNTIF and D* database functions now handle external workbook ranges as expected
- New: 3D references are now supported in calculation engine
- Updated: COUNTIF/SUMIF/AVERAGEIF tilde escaping updated to match Excel behavior
- Updated: LINEST/LOGEST/TREND/GROWTH regression statistics now match Excel more closely, including exact-fit and near-exact-fit cases
- Updated: LINEST(..., FALSE, TRUE) now returns Excel-style #N/A in the second-row/second-column stats slot
- Updated: LOGEST(..., FALSE, TRUE) standard-error output now matches Excel instead of incorrectly exponentiating the slope standard error
- Updated: MAP/SCAN/BYROW/BYCOL/MAKEARRAY/REDUCE now treat 1x1 lambda results as scalars in array-evaluation contexts, matching Excel
- Updated: MAKEARRAY now accepts single-cell references for rows/cols arguments, matching Excel
- Updated: REGEXREPLACE negative occurrence handling now matches Excel, including counting from the end and returning the original text when out of range
- Updated: SUM/AVERAGE/COUNT/MIN/MAX/PRODUCT/SUMSQ/AVEDEV/STDEV*/VAR* reference-text coercion now matches Excel for numeric-looking text cells
- Updated: AVERAGEA/MAXA/MINA/STDEVA/VARA/VARPA and related *A functions now distinguish plain text, quote-prefixed text, and formula-text in references the same way Excel does

## 0.6.0

- New: All API calls are now org-scoped (`/v0/orgs/:org_id/...`) when authenticated, enabling multi-org support
- New: Login flow prompts for org selection when the user belongs to multiple organizations
- Updated: File and revision cache keys now include org ID to prevent cross-org collisions

## 0.5.0

- New: Array formulas are now supported for authoring, calculation, dependency tracking, with full open-save roundtrip fidelity
- New: LAMBDA/LET/REDUCE/MAP/SCAN/MAKE_ARRAY/BYROW/CYCOL functions are now fully supported for authoring, calculation, dependency tracking, with full open-save roundtrip fidelity
- New: Calculation of What-If Data Tables is now fully supported, automatically recalculating as needed when upstream cells are modified through `setCells`
- Updated: External workbook references in formulas are now resolved against the workbook cache when recalculating
- Updated: All methods that output formulas now do in display-format style, ie. without `_xlfn`-like prefixes used in stored formulas
- Updated: All methods that accept formulas as input now consistently accept them both with and without leading `=`, with and without stored formula prefixes like `_xlfn`
- Updated: Iterative calculation worbook settings can now be read and modified from `get/setWorkbookProperties`, and they are persisted accordingly on save
- Updated: `readRange`, `readRangeTsv`, `previewStyles` and all other methods that accept range addresses now also resolve table names, defined names, structured references
- Updated: `listSheets` now omits null fields in the response, to save tokens on commonly missing/empty fields
- Updated: `listSheets` now returns list of what-if data table addresses for each sheet, which can be passed directly to `readRangeTsv` etc to inspect
- Updated: `listSheets` now returns list of ListObject table names for each sheet, which can be passed directly to `readRangeTsv` etc to inspect

## 0.4.0

- New: `sortRange` operation to sort rows by column keys.
- New: `autoFitColumns` operation to auto-fit column widths to cell content.
- New: `findAndReplace` for bulk text substitution with regex and formula support.
- New: `copyRange` operation to copy ranges with formula reference adjustment.
- New: `sweepInputs` operation for batch what-if analysis with compact TSV output.
- New: `getConditionalFormatting`, `setConditionalFormatting`, `removeConditionalFormatting` for reading, adding, and removing conditional formatting rules (`iconSet` currently read-only in write payloads).
- Breaking: `detectTables` replaced by `describeSheets` — returns per-sheet tables + compact ASCII structure map showing cell types, row collapsing, and inline label annotations.
- Updated: `readRange`, `readRow`, `readColumn`, `readCell`: now include `note`, `link`, `thread` fields when cells have comments, hyperlinks, or threaded comments.
- Updated: `setCells`: now supports `note`, `link`, and `thread` fields for setting/clearing comments, hyperlinks, and threaded comments (with inline person upsert).
- Updated: `listSheets`: now returns list of dependent and precedent sheets for each sheet

## 0.3.0

### Commands

- `witan xlsx calc`: Recalculates workbook formulas, surfaces formula errors, and can run in `--verify` mode to detect value drift without mutating the file.
- `witan xlsx exec`: Runs sandboxed JavaScript against a workbook to read, search, trace, and (optionally) persist edits.
- `witan xlsx lint`: Performs semantic formula analysis to catch spreadsheet risks like double counting, lookup issues, and coercion surprises.
- `witan xlsx render`: Renders a sheet range to an image for visual layout, formatting, and diff-based change checks.

### Skills

- `xlsx-code-mode`: Workbook exploration/editing workflow centered on `witan xlsx exec` for reading data, tracing logic, and running what-if updates.
