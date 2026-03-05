# Changelog

## Unreleased

## 0.3.0

### Commands

- `witan xlsx calc`: Recalculates workbook formulas, surfaces formula errors, and can run in `--verify` mode to detect value drift without mutating the file.
- `witan xlsx exec`: Runs sandboxed JavaScript against a workbook to read, search, trace, and (optionally) persist edits.
- `witan xlsx lint`: Performs semantic formula analysis to catch spreadsheet risks like double counting, lookup issues, and coercion surprises.
- `witan xlsx render`: Renders a sheet range to an image for visual layout, formatting, and diff-based change checks.

### Skills

- `xlsx-code-mode`: Workbook exploration/editing workflow centered on `witan xlsx exec` for reading data, tracing logic, and running what-if updates.
