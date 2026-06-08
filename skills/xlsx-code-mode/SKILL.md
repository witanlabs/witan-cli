---
name: xlsx-code-mode
description: Use this skill for any Excel workbook task involving .xls, .xlsx, or .xlsm files: read, inspect, search, understand, create, modify, repair, author, or verify spreadsheets. Trigger for workbook creation from scratch; edits to cells, formulas, sheets, styles, tables, charts, images, conditional formatting, data validation, defined names, metadata, or sheet properties; dependency tracing; what-if analysis; and Excel-grade verification with lint, calc, render. Also trigger whenever the user references a spreadsheet file by name or path, even casually. The tool runs sandboxed JavaScript against the workbook server-side via `witan xlsx exec`."
---

> **Running in Claude Cowork?** The `witan` CLI isn't preinstalled in the sandbox. Read [references/cowork-setup.md](references/cowork-setup.md) first for install, PATH and network-allowlist steps.

> **No `witan` on PATH?** Prefix commands with `npx witan` (e.g. `npx witan xlsx exec ...`).

## Setup

The CLI supports `.xls`, `.xlsx`, and `.xlsm`; legacy `.xls` files are converted to `.xlsx` when needed. New workbook creation is `.xlsx` only.

## Quick Reference

```bash
# Create a workbook (.xlsx only)
witan xlsx exec model.xlsx --create --save --stdin <<'WITAN'
await xlsx.addSheet(wb, "Inputs")
await xlsx.setCells(wb, [{ address: "Inputs!A1", value: "Revenue" }])
return await xlsx.listSheets(wb)
WITAN

# Read awkward sheet names; apostrophes inside sheet names are doubled
witan xlsx exec model.xlsx --stdin <<'WITAN'
const a = await xlsx.readCell(wb, "'Workers'' Compensation'!B50")
const b = await xlsx.readRangeTsv(wb, { sheet: "Reserve Summary (Net)", from: {row:1,col:1}, to: {row:10,col:5} })
return { a: a.value, b }
WITAN

# Sensitivity sweep
witan xlsx exec model.xlsx --stdin <<'WITAN'
const result = await xlsx.sweepInputs(wb, {
	inputs: [
		{ address: "Inputs!B5", values: [0.02, 0.04, 0.06] },
		{ address: "Inputs!B6", values: [0.08, 0.10, 0.12] },
	],
	outputs: ["Output!C30", "Output!C45"],
	mode: "cartesian",
	includeStats: true,
})
console.log(result.tsv)
WITAN

# Author a chart and preview it
witan xlsx exec model.xlsx --stdin <<'WITAN'
await xlsx.addChart(wb, "Sheet1", {
	name: "Revenue",
	position: { from: { cell: "F2" }, to: { cell: "N18" } },
	groups: [
		{
			type: "column",
			series: [
				{
					name: { ref: "Sheet1!B1" },
					categories: "Sheet1!A2:A9",
					values: "Sheet1!B2:B9"
				}
			]
		}
	],
	title: { text: "Revenue" },
	legend: { position: "right" }
})
await xlsx.previewStyles(wb, "Sheet1!F2:N18")
WITAN

# Simple one-liner
witan xlsx exec model.xlsx --expr 'xlsx.listSheets(wb)'
```

## exec — Workbook Scripting

Runs server-side JavaScript against a workbook through globals `xlsx`, `wb`, `input`, and `print`. Top-level `await` is supported; imports are not.

Use `--create` only with new `.xlsx` paths. Add `--save` to write bytes to disk; without `--save`, creation and edits are session-only.

### Invocation patterns

Prefer `--stdin` with a single-quoted heredoc for multi-line scripts and any sheet name with spaces, apostrophes, parentheses, quotes, or glob characters:

```bash
witan xlsx exec report.xlsx --stdin <<'WITAN'
const sheets = await xlsx.listSheets(wb)
const cell = await xlsx.readCell(wb, "'My Sheet'!A1")
return { sheets, cell }
WITAN
```

Use exactly one code source: `--stdin`, `--expr`, `--code`, or `--script`. `--expr` is fine for simple one-liners; use `--script file.js --input-json ...` for reusable scripts.

### Flags

- Code source: exactly one of `--stdin`, `--expr`, `--code`, `--script`
- Inputs/control: `--input-json`, `--input-file key=@path`, `--timeout-ms`, `--max-output-chars`, `--stdin-timeout-ms`, `--locale`
- Lifecycle/output: `--create`, `--save`, `--json`

### Runtime globals

- `xlsx`: curated API; pass `wb` as first arg to all functions
- `wb`: opened workbook handle
- `input`: parsed `--input-json`; `--input-file logo=@./logo.png` adds `input.logo` as `data:image/...;base64,...`
- `print`: stdout helper like `console.log`

### API reference

All functions are async and take `wb` first.

- Read: `listSheets`, `getWorkbookProperties`, `getSheetProperties`, `listDefinedNames`, `readCell`, `readRange`, `readRow`, `readColumn`, TSV variants, `getStyle`
- Search: `findCells`, `findRows`, `describeSheet`, `tableLookup`, `getListObject`, `getDataTable`; `matcher` accepts string/string[]/number/boolean/RegExp/RegExp[]
- Trace/compute/render: `getCellPrecedents`, `getCellDependents`, `traceToInputs`, `traceToOutputs`, `sweepInputs`, `evaluateFormula(s)`, `lint`, `previewStyles`
- Write/structure: `setCells`, `findAndReplace`, `scaleRange`, `copyRange`, `sortRange`, row/column insert/delete, `autoFitColumns`, sheet/name/property/style setters
- Objects: `add/set/deleteListObject`, `add/deleteDataTable`, conditional formatting and data validation APIs
- Charts/images: `list/get/add/set/deleteChart`, `list/get/add/set/deleteImage`; charts support combo, secondary axes, stock, bubble, radar, top-view surface, waterfall, histogram/Pareto, funnel, formatting, labels, number formats, style IDs

Common options: searches use `opts.in`, `context`, `limit`, `offset`; TSV reads use `includeEmpty`/`includeFormulas`; `scaleRange.skipFormulas` defaults true; `sortRange.hasHeader` defaults true; `copyRange.pasteType` supports `all`, `values`, `formulas`, `formats`; `setConditionalFormatting` and `setDataValidations` use `opts.clear` to replace existing rules.

For local images, prefer `--input-file logo=@./logo.png` then `source:{base64:input.logo}`. `source.base64` also accepts raw base64 or data URLs; image responses return metadata only.

### Write and What-If Rules

`exec` is ephemeral unless `--save` is passed. Each invocation starts from the original workbook; writes such as `setCells`, `scaleRange`, inserts, deletes, formatting, charts, and images affect only the server-side session unless saved. `setCells` creates a missing referenced sheet and returns:

```ts
{ touched:Record<string,string>; changed:string[]; errors:Diag[] }
```

Read output values from `result.touched["Sheet!Address"]`; do not recompute answers in JavaScript.

For "what happens to Y if X changes?":

1. Separate exec: find the output cell first with `findCells(wb, matcher, { context: 2 })`, synonyms, or `readRangeTsv`; choose the formula/output cell, not the label.
2. Separate exec: `traceToInputs(wb, outputAddr)`, confirm the user-named input drives it, `setCells`, then read `result.touched[outputAddr]`.
3. Report baseline and changed values, and check `result.errors`.

Use `sweepInputs` for multi-value sensitivities. If tracing returns hundreds of cells, filter by `nearbyLabel` or context before printing. For iterative/circular models, `setCells` uses workbook iterative settings and reports convergence errors.

### calc — Full-workbook verification

`setCells` already recalculates edited cells and downstream dependents. Use `calc` for standalone workbook-wide checks or final audits:

```bash
witan xlsx calc model.xlsx
witan xlsx calc model.xlsx --verify
witan xlsx calc model.xlsx --show-touched
```

It prints all formula errors, not just a count. Exit code `2` means formula errors; with `--verify`, exit code `2` can also mean computed values changed. `--show-touched` prints every recalculated cell.

### Response format

With `--json`, success includes `ok`, `stdout`, `result`, `writes_detected`, and `accesses`; failure includes `ok:false`, `stdout`, and `error:{type,code,message}`. `accesses` records cell reads/writes.

## render — Visual Screenshot

Renders a sheet range as PNG/WebP for layout, merged cells, formatting, charts, and labels.

```bash
witan xlsx render report.xlsx -r "Sheet1!A1:Z50"
witan xlsx render report.xlsx -r "'My Sheet'!B5:H20" -o snapshot.png
witan xlsx render report.xlsx -r "Sheet1!A1:F10" --dpr 2 --diff before.png
```

Flags: `--range/-r` required, `--output/-o` optional, `--dpr` 1-3 or `auto`, `--format png|webp`, `--diff baseline.png`. Inside `exec`, use `previewStyles`.

## Error Guide

- Code source errors: provide exactly one non-empty `--stdin`, `--expr`, `--code`, or `--script`.
- `Import statements are not allowed`: use the `xlsx` global, no static/dynamic imports.
- `EXEC_SYNTAX_ERROR` / `EXEC_RUNTIME_ERROR`: fix script syntax/runtime issue.
- `EXEC_RESULT_TOO_LARGE`: return less data; print/log large output.
- `invalid --input-json`, `--timeout-ms must be > 0`: fix flag value.
- `Sheet 'X' not found`: run `listSheets`; for shell quoting issues use `--stdin <<'WITAN'`.
- `findCells` empty: broaden search, use synonyms, or inspect nearby ranges.
- `setCells` missing expected output in `touched`: wrong output cell or not a dependent; trace the chain.

### Full Type Definitions

Exact signatures are split by API area under `references/api/`. Open only the file relevant to the task:

- `references/api/workbook-sheet.d.ts`: workbook properties, sheet properties, defined names, sheet lifecycle
- `references/api/read-search.d.ts`: cell/range/row/column reads, TSV helpers, cell/row search, sheet description, table lookup
- `references/api/write-structure.d.ts`: `setCells`, range transforms, copy/sort, row/column insertion/deletion, autofit
- `references/api/styles-validation.d.ts`: styles, conditional formatting, data validations, validation checks
- `references/api/compute-trace-render.d.ts`: `sweepInputs`, precedents/dependents, input/output tracing, formula evaluation, lint, `previewStyles`
- `references/api/charts.d.ts`: chart specs and chart list/get/add/set/delete APIs
- `references/api/images.d.ts`: image specs and image list/get/add/set/delete APIs
