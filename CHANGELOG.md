# Changelog

## Unreleased

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
