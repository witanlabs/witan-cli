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
# Create a new workbook from scratch (.xlsx only)
witan xlsx exec model.xlsx --create --save --stdin <<'WITAN'
await xlsx.addSheet(wb, "Inputs")
await xlsx.setCells(wb, [{ address: "Inputs!A1", value: "Revenue" }])
return await xlsx.listSheets(wb)
WITAN

# Read from sheets with spaces, apostrophes, or parentheses — note inner apostrophes are doubled (Excel convention)
witan xlsx exec model.xlsx --stdin <<'WITAN'
const a = await xlsx.readCell(wb, "'Workers'' Compensation'!B50")
const b = await xlsx.readRangeTsv(wb, { sheet: "Reserve Summary (Net)", from: {row:1,col:1}, to: {row:10,col:5} })
return { a: a.value, b }
WITAN

# Multi-input sweep — compare several input values at once
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

# Conditional formatting — add a highlight rule and a color scale
witan xlsx exec model.xlsx --stdin <<'WITAN'
await xlsx.setConditionalFormatting(wb, "Sheet1", [
	{
		type: "cellValue",
		address: "A1:A100",
		operator: "greaterThan",
		formula: "100",
		style: { fill: { color: "#FF0000" } }
	},
	{
		type: "twoColorScale",
		address: "B1:B100",
		lowValue: { type: "min", color: "#FFFFFF" },
		highValue: { type: "max", color: "#FF0000" }
	}
])
WITAN

# Chart authoring — add an embedded chart and verify the rendered result
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

# Image authoring — add a local PNG/JPEG and verify placement
witan xlsx exec model.xlsx --save --input-file logo=@./logo.png --stdin <<'WITAN'
await xlsx.addImage(wb, "Sheet1", {
	name: "Logo",
	position: { from: { cell: "A1" }, to: { cell: "D6" } },
	source: { base64: input.logo },
	altText: "Company logo"
})
return await xlsx.listImages(wb, { sheet: "Sheet1" })
WITAN

# Waterfall chart authoring with totals and connector lines
witan xlsx exec model.xlsx --save --stdin <<'WITAN'
await xlsx.addChart(wb, "Sheet1", {
	name: "Bridge",
	position: { from: { cell: "F2" }, to: { cell: "N18" } },
	groups: [
		{
			type: "waterfall",
			series: [
				{
					name: { text: "Movement" },
					categories: "Sheet1!A2:A8",
					values: "Sheet1!B2:B8",
					totalIndexes: [0, 6],
					showConnectorLines: true,
					dataLabels: { showValue: true, position: "outsideEnd" }
				}
			]
		}
	],
	title: { text: "Revenue Bridge" },
	plotVisibleOnly: true,
	styleId: 395
})
await xlsx.previewStyles(wb, "Sheet1!F2:N18")
WITAN

# ListObject authoring — create an Excel table and read it back by table name
witan xlsx exec model.xlsx --stdin <<'WITAN'
await xlsx.addListObject(wb, "Sheet1", {
	name: "SalesTable",
	ref: "A1:C4",
	showTotalsRow: true,
	columns: [
		{ name: "Region", totalsRowLabel: "Total" },
		{ name: "Sales", totalsRowFunction: "sum" },
		{ name: "DoubleSales", calculatedColumnFormula: "=B2*2" }
	],
	rows: [
		[{ value: "North" }, { value: 10 }, {}],
		[{ value: "South" }, { value: 20 }, {}]
	]
})
return {
	meta: await xlsx.getListObject(wb, "SalesTable"),
	data: await xlsx.readRange(wb, "SalesTable")
}
WITAN

# What-If Data Table authoring — create a visible one-variable table block
witan xlsx exec model.xlsx --stdin <<'WITAN'
await xlsx.addDataTable(wb, "Sheet1", {
	type: "oneVariableColumn",
	ref: "E1:F4",
	columnInputCell: "H1",
	inputValues: [5, 10, 15],
	formulas: ["=H1*2"]
})
return await xlsx.getDataTable(wb, "Sheet1!E1:F4")
WITAN

# Simple one-liner (--expr is fine when there are no special characters)
witan xlsx exec model.xlsx --expr 'xlsx.listSheets(wb)'
```

## exec — Workbook Scripting

Runs JavaScript against a workbook via the Witan API. The workbook is opened server-side; scripts interact through the `xlsx` and `wb` globals.

If the target workbook does not exist yet, use `--create` with a new `.xlsx` path.

- Use `--create --save` when you want to produce a real workbook file on disk.
- Use `--create` without `--save` when you want to prototype workbook structure, test generation logic, inspect returned data, or validate formulas/layout without leaving a file behind.

### Invocation patterns

**Recommended: `--stdin` with heredoc** — safe for all sheet names, supports multi-line scripts, and batches multiple operations into a single CLI invocation:

```bash
witan xlsx exec report.xlsx --stdin <<'WITAN'
const sheets = await xlsx.listSheets(wb)
const cell = await xlsx.readCell(wb, "'My Sheet'!A1")
return { sheets, cell }
WITAN
```

The single-quoted heredoc delimiter (`<<'WITAN'`) prevents all shell expansion. Apostrophes, parentheses, double quotes, and glob characters in sheet names pass through verbatim to JavaScript — no escaping needed.

**Other invocation patterns** (use only when `--stdin` is impractical):

```bash
# Expression — simple one-liners with no special characters
witan xlsx exec report.xlsx --expr 'xlsx.listSheets(wb)'

# Script file — reusable scripts, e.g. parameterized scenarios
witan xlsx exec report.xlsx --script scenario.js --input-json '{"rate": 1.05}'
```

Provide exactly one code source: `--expr`, `--code`, `--script`, or `--stdin`. They are mutually exclusive.

### Flags

- Code source: exactly one of `--stdin`, `--expr`, `--code`, or `--script`.
- Inputs/control: `--input-json`, `--input-file key=@path`, `--timeout-ms`, `--max-output-chars`, `--stdin-timeout-ms`, `--locale`.
- Workbook lifecycle: `--create` creates a new `.xlsx` session; `--save` persists changes; `--json` prints the full response envelope.

### Runtime globals

- `xlsx` (`object`): curated API surface; all functions listed below
- `wb` (`WorkbookContext`): opened workbook handle; pass as first arg to all `xlsx.*` calls
- `input` (`any`): parsed value from `--input-json` (defaults to `{}`); `--input-file key=@path` adds a PNG/JPEG as a `data:image/...;base64,...` string on that top-level key.
- `print` (`function`): output to stdout, like `console.log` but captured in the response

Top-level `await` is supported. No imports allowed (static or dynamic).

### API reference

Functions are grouped by purpose. All are async and take `wb` as the first argument.

**Reading**

- `listSheets`: sheet inventory with used ranges, visibility, and cross-sheet dependencies
- `getWorkbookProperties`, `getSheetProperties`: workbook/sheet metadata, formatting, view, outline, and merge info
- `listDefinedNames`: workbook and sheet-scoped names
- `readCell`, `readRange`, `readRow`, `readColumn`: structured cell reads
- `readRangeTsv`, `readRowTsv`, `readColumnTsv`: TSV reads for prompt-friendly extraction
- `getStyle`: resolved style properties for a cell

**Searching**

- `findCells`, `findRows`: fuzzy search by value or pattern
- `describeSheet`: sheet structure with detected tables
- `tableLookup`: lookup by row and column labels inside a table
- `getListObject`, `getDataTable`: metadata for existing Excel table / What-If Data Table objects

`matcher` may be a string, string array, number, boolean, `RegExp`, or `RegExp[]`. Searches are fuzzy and case-insensitive by default.

**Tracing**

- `getCellPrecedents`, `getCellDependents`: local dependency traversal; `depth` defaults to `1`
- `traceToInputs`, `traceToOutputs`: full graph tracing to leaf inputs or terminal outputs

**Computing**

- `sweepInputs`: batch what-if sweeps with TSV and structured outputs
- `evaluateFormula`, `evaluateFormulas`: evaluate one or more formulas in sheet context

**Validating**

- `lint`: potential workbook issues and diagnostics
- `getDataValidations`, `validateCells`: inspect data validation rules and find invalid cells

**Rendering**

- `previewStyles`: generate a PNG preview for a sheet range; image is auto-registered

**Charts**

- `listCharts`: chart summaries for the workbook or a single sheet
- `getChart`: canonical spec for an existing chart
- `addChart`, `setChart`, `deleteChart`: create, replace, or remove embedded charts
- Supported specs include combo charts, secondary axes, stock charts, bubble charts, radar charts, 2-D top-view surface charts, waterfall charts, histogram/Pareto charts, funnel charts, chart/plot-area formatting, group/series data labels, linked number formats, and style IDs. Use `previewStyles` after authoring to inspect rendered placement and labels.

**Images**

- `listImages`: image metadata for the workbook or a single sheet
- `getImage`: metadata for one worksheet image by `{ name }` or `{ id }`
- `addImage`, `setImage`, `deleteImage`: create, update, replace, or remove embedded PNG/JPEG images
- For `addImage`/`setImage` from local files, prefer `witan xlsx exec ... --input-file logo=@./logo.png` and use `source: { base64: input.logo }`. `source.base64` also accepts raw base64 or `data:image/png;base64,...` / `data:image/jpeg;base64,...`; responses return metadata only, not image bytes. `preserveAspectRatio` defaults to `true` when adding or replacing image bytes.

**Conditional Formatting**

- `getConditionalFormatting`: read all sheet rules; `iconSet` is read-only
- `setConditionalFormatting`: add writable rules; `opts.clear` replaces all rules
- `removeConditionalFormatting`: remove rules by index

**Writing (ephemeral)**

- Cells/ranges: `setCells`, `findAndReplace`, `scaleRange`, `copyRange`, `sortRange`
- Data validation: `setDataValidations`, `removeDataValidations`; `setCells` supports `opts.validationMode: "reject"` to reject invalid writes
- Structure: `insertRowAfter`, `deleteRows`, `insertColumnAfter`, `deleteColumns`, `autoFitColumns`
- Sheets/names: `addSheet`, `deleteSheet`, `renameSheet`, `addDefinedName`, `deleteDefinedName`
- Properties/styles: `setWorkbookProperties`, `setSheetProperties`, `setRowProperties`, `setColumnProperties`, `setStyle`
- Objects: `addListObject`, `setListObject`, `deleteListObject`, `addDataTable`, `deleteDataTable`

Common option patterns:

- Search: `opts.in`, `context`, `limit`, `offset`; `findCells` also supports `formulas`
- Row/column reads: `startRow/endRow` or `startCol/endCol`
- TSV reads: `includeEmpty`, `includeFormulas`
- `scaleRange`: `opts.skipFormulas` defaults to `true`
- `sortRange`: `opts.hasHeader` defaults to `true`
- `copyRange`: `opts.pasteType` supports `all`, `values`, `formulas`, `formats`
- `setConditionalFormatting`: `opts.clear` replaces all existing rules
- `setDataValidations`: `opts.clear` replaces all existing rules on the sheet

### The ephemeral write contract

By default, `exec` **does not write workbook bytes back to disk**. All write operations (`setCells`, `scaleRange`, inserts, deletes) take effect in the server-side session only. The `result.touched` map contains the recalculated formatted text values — read answers from there.

This means:

- No risk of corrupting the original file
- No `reset()` needed — each invocation starts clean
- For independent what-if tests, each `exec` invocation starts from the original file

To persist changes back to the workbook file, pass the `--save` flag.

When `--create` is set, the same ephemeral rule applies to the new workbook session: no local file is created unless `--save` is also passed. New workbook creation is `.xlsx` only.

### setCells result shape

```ts
{
  touched: Record<string, string>  // address → formatted text value
  changed: string[]                // addresses whose values changed
  errors: Diag[]                   // cells that errored after recalc
}
```

Read the output value from `result.touched["Sheet!Address"]`. Never compute the answer in JavaScript. `setCells` implicitly creates a sheet if `address` references one that does not yet exist.

### What-if / sensitivity workflow

For questions like "what happens to Y if X changes?", follow this sequence. **Steps 1 and 2 should be separate `exec` calls** — review the output of step 1 before proceeding.

1. **Find the output cell (separate exec call)** — search for what the user is asking about (e.g., "net income", "reserve"). Use `xlsx.findCells` with `context: 2` and review the candidates. Labels and formula cells often share the same text — pick the formula cell, not the label. This step may take more than one attempt: try synonym arrays, read surrounding rows with `xlsx.readRangeTsv`, or check multiple sheets.
2. **Trace + run the what-if (second exec call)** — once you have the output address, call `xlsx.traceToInputs(wb, outputAddr)` to confirm the user's input actually drives it. Trace results can contain hundreds of cells — filter by `nearbyLabel` matching the user's term instead of printing them all. Then `xlsx.setCells` to make the change and read the answer from `result.touched[outputAddr]`. If the output address is missing from `touched`, the cell didn't recalculate — you likely have the wrong address.
3. **Report before and after** — always include the baseline value (read before the edit) and the new value from `touched`.

For sweeping multiple values (sensitivity tables), use `sweepInputs` instead — it runs all combinations in one call and returns structured before/after data.

Practical tips:

- Don't search for the output _after_ `setCells` — find it first so you know the exact address to check in `touched`.
- If `findCells` returns several hits for the same label, use context or `readRangeTsv` to disambiguate (e.g., "Net Income" may appear as both a label and a formula cell on different rows).
- After `setCells`, verify `result.errors` is empty. New errors mean the edit introduced or surfaced a calculation problem downstream.
- If the question names a specific metric (e.g., "loss ratio"), don't just search for that term — also check the sheet/row the question references, since models often have multiple versions of the same metric across sheets.

### Iterative / circular models

When a workbook has **iterative calculation** enabled, `setCells` recalculates
the edited cells and downstream dependents using the workbook's iterative
settings. If the model does not converge within `maxIterations`, the result
includes convergence errors.

Use `witan xlsx calc` when you want a standalone seeded or full-workbook
verification pass, or when you want CLI reporting of all calculation errors or
changed cells. `--show-touched` only changes output verbosity.

### calc — Full-workbook verification

`setCells` already recalculates the edited cells and all downstream dependents,
so you do **not** need `witan xlsx calc` after every normal edit. Use `calc`
mainly as a standalone verification/reporting command:

- See every formula error in one CLI call
- In `--verify` mode, see which cell values changed without modifying the file
- Run a workbook-wide verification pass outside your `exec` script

```bash
# Recalculate the workbook and print all formula errors, if any
witan xlsx calc model.xlsx

# Verification pass only — no file write, but prints changed addresses
witan xlsx calc model.xlsx --verify

# Verbose output — every touched cell with formula/value or error code
witan xlsx calc model.xlsx --show-touched
```

Default output is concise:

- If there are no formula errors, it prints a one-line summary like `428 cells recalculated, 0 errors, 3 changed`
- If there are formula errors, it prints **all** of them, not just a count
- Without `--verify`, it still exits with code `2` if any formula errors are found
- With `--verify`, it also prints a sorted `Changed (N):` list of addresses
- `--show-touched` is the verbose mode when you need every recalculated cell

Example error output:

```text
2 errors:
  Summary!C18          =A18/B18                      #DIV/0!
  Revenue!F42          =VLOOKUP(A42,$A$2:$C$10,3,0) #N/A
```

Example verify output:

```text
428 cells recalculated, 0 errors, 3 changed

Changed (3):
  Inputs!B5
  Revenue!F42
  Summary!C18
```

`witan xlsx calc` exits with code `2` if any formula errors are found.
`witan xlsx calc --verify` also exits with code `2` if any computed values
changed. That makes `--verify` useful as a final audit step after fixing a
workbook.

### Response format

When `--json` is used, the full response envelope is returned:

**Success:**

```json
{
  "ok": true,
  "stdout": "...",
  "result": "<json>",
  "writes_detected": false,
  "accesses": [...]
}
```

**Failure:**

```json
{
  "ok": false,
  "stdout": "...",
  "error": { "type": "...", "code": "...", "message": "..." }
}
```

The `accesses` array documents all cell reads and writes with operation type and address.

## render — Visual Screenshot

Renders a sheet range as a PNG image, useful for inspecting layout, merged cells, formatting, charts, and labels.

```bash
# Render a range to a temporary file (path printed to stdout)
witan xlsx render report.xlsx -r "Sheet1!A1:Z50"

# Render to a specific output path
witan xlsx render report.xlsx -r "'My Sheet'!B5:H20" -o snapshot.png

# Higher resolution (DPR 1-3; auto picks 1 or 2)
witan xlsx render report.xlsx -r "Sheet1!A1:F10" --dpr 2

# Diff against a baseline — highlights changes in a new PNG
witan xlsx render report.xlsx -r "Sheet1!A1:F10" --diff before.png
```

- `--range`, `-r`: sheet-qualified range to render; required
- `--output`, `-o`: output path; defaults to a temporary file
- `--dpr` (default `auto`): device pixel ratio `1`-`3`
- `--format` (default `png`): `png` or `webp`
- `--diff`: compare against a baseline PNG and write a diff image

The `previewStyles` exec function (see Rendering in the API reference) provides the same capability from within a script.

## Error Guide

- `exactly one of --code, --script, --stdin, or --expr is required`: provide exactly one code source flag
- `--code, --script, --stdin, and --expr are mutually exclusive`: use only one code source flag per invocation
- `exec code must not be empty`: provide non-empty code
- `Import statements are not allowed`: no `import` in exec scripts; use the `xlsx` global
- `EXEC_SYNTAX_ERROR`: fix JavaScript syntax in your script
- `EXEC_RUNTIME_ERROR`: fix the runtime error; check the message for details
- `EXEC_RESULT_TOO_LARGE`: return less data; use `console.log()` for large output instead of return values
- `--timeout-ms must be > 0`: omit the flag or provide a positive value
- `invalid --input-json`: provide valid JSON
- `Sheet 'X' not found`: check the sheet name; use `listSheets` to enumerate
- `Shell quoting errors with sheet names`: use `--stdin <<'WITAN'`; it avoids shell quoting issues
- `findCells` returns empty: try synonym arrays, broader search, or check spelling
- `setCells` result missing expected output: the output cell may not be a dependent; trace the formula chain

### Full Type Definitions

Exact signatures are split by API area under `references/api/`. Open only the file relevant to the task:

- `references/api/workbook-sheet.d.ts`: workbook properties, sheet properties, defined names, sheet lifecycle
- `references/api/read-search.d.ts`: cell/range/row/column reads, TSV helpers, cell/row search, sheet description, table lookup
- `references/api/write-structure.d.ts`: `setCells`, range transforms, copy/sort, row/column insertion/deletion, autofit
- `references/api/styles-validation.d.ts`: styles, conditional formatting, data validations, validation checks
- `references/api/compute-trace-render.d.ts`: `sweepInputs`, precedents/dependents, input/output tracing, formula evaluation, lint, `previewStyles`
- `references/api/charts.d.ts`: chart specs and chart list/get/add/set/delete APIs
- `references/api/images.d.ts`: image specs and image list/get/add/set/delete APIs
