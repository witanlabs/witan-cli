# Changelog

## Unreleased

- New: `sortRange` operation to sort rows by column keys.
- New: `autoFitColumns` operation to auto-fit column widths to cell content.
- New: `findAndReplace` for bulk text substitution with regex and formula support.
- New: `copyRange` operation to copy ranges with formula reference adjustment.
- New: `scenarios` operation for batch what-if analysis with compact TSV output.
- New: `getConditionalFormatting`, `setConditionalFormatting`, `removeConditionalFormatting` for reading, adding, and removing conditional formatting rules (`iconSet` currently read-only in write payloads).
- Breaking: `detectTables` replaced by `describeSheets` — returns per-sheet tables + compact ASCII structure map showing cell types, row collapsing, and inline label annotations.
- `readRange`, `readRow`, `readColumn`, `readCell`: now include `note`, `link`, `thread` fields when cells have comments, hyperlinks, or threaded comments.
- `setCells`: now supports `note`, `link`, and `thread` fields for setting/clearing comments, hyperlinks, and threaded comments (with inline person upsert).

## 0.3.0

### Commands

- `witan xlsx calc`: Recalculates workbook formulas, surfaces formula errors, and can run in `--verify` mode to detect value drift without mutating the file.
- `witan xlsx exec`: Runs sandboxed JavaScript against a workbook to read, search, trace, and (optionally) persist edits.
- `witan xlsx lint`: Performs semantic formula analysis to catch spreadsheet risks like double counting, lookup issues, and coercion surprises.
- `witan xlsx render`: Renders a sheet range to an image for visual layout, formatting, and diff-based change checks.

### Skills

- `xlsx-code-mode`: Workbook exploration/editing workflow centered on `witan xlsx exec` for reading data, tracing logic, and running what-if updates.
