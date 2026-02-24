---
name: xlsx-code-mode
description: Witan CLI exec API reference — workbook scripting, function tables, type definitions.
---

## Setup

Files are cached server-side by content hash so repeated operations skip re-upload. If `WITAN_STATELESS=1` is set (or `--stateless` is passed), files are processed but not stored.

The CLI automatically applies per-attempt request timeouts and retries transient API failures (`408`, `429`, `500`, `502`, `503`, `504`, plus timeout/network errors). Non-retryable `4xx` responses fail immediately.

## Quick Reference

```bash
# Explore — map out sheets and find data
witan xlsx exec model.xlsx --expr 'xlsx.listSheets(wb)'
witan xlsx exec model.xlsx --expr 'xlsx.findCells(wb, "Revenue", { context: 1 })'
witan xlsx exec model.xlsx --expr 'xlsx.readCell(wb, "Summary!C10")'

# What-if — change an input and read the recalculated output
witan xlsx exec model.xlsx --code '
  const result = await xlsx.setCells(wb, [
    { address: "Inputs!B5", value: 1.10 }
  ]);
  return { touched: result.touched, errors: result.errors };
'

# Parameterized what-if — same script, different scenarios
witan xlsx exec model.xlsx --script scenario.js --input-json '{"rate": 1.05}'
witan xlsx exec model.xlsx --script scenario.js --input-json '{"rate": 1.10}'
```

## Exit Codes

| Command | Code | Meaning                                                             |
| ------- | ---- | ------------------------------------------------------------------- |
| `exec`  | `0`  | Script completed successfully (`ok: true`)                          |
| `exec`  | `1`  | Transport/API error, invalid request, or script error (`ok: false`) |

## exec — Workbook Scripting

Runs JavaScript against a workbook via the Witan API. The workbook is opened server-side; scripts interact through the `xlsx` and `wb` globals.

### Invocation patterns

```bash
# Expression — wraps as return (<expr>);
witan xlsx exec report.xlsx --expr 'xlsx.listSheets(wb)'

# Inline code — full script body
witan xlsx exec report.xlsx --code '
  const sheets = await xlsx.listSheets(wb);
  print(sheets);
  return sheets;
'

# Script file
witan xlsx exec report.xlsx --script explore.js

# Stdin
cat explore.js | witan xlsx exec report.xlsx --stdin
```

Provide exactly one code source: `--expr`, `--code`, `--script`, or `--stdin`. They are mutually exclusive.

### Shell quoting for sheet names with spaces

Sheet names with spaces or special characters require care when passed through bash. Single-quoted bash (`--expr '...'`) cannot contain literal single quotes, so Excel-style `'Sheet Name'!A1` addresses break.

**Use coordinate objects** (always safe inside single-quoted bash):

```bash
witan xlsx exec file.xlsx --expr 'xlsx.readRangeTsv(wb, { sheet: "My Sheet", from: { row: 1, col: 1 }, to: { row: 50, col: 10 } })'
witan xlsx exec file.xlsx --expr 'xlsx.readCell(wb, { sheet: "My Sheet", row: 5, col: 3 })'
```

**Or use functions with a separate `sheet` parameter** (no address quoting needed):

```bash
witan xlsx exec file.xlsx --expr 'xlsx.readRowTsv(wb, "My Sheet", 1)'
witan xlsx exec file.xlsx --expr 'xlsx.readRow(wb, "My Sheet", 5)'
```

**If you must use A1-style addresses with sheet names containing spaces**, use double-quoted bash for `--expr`:

```bash
witan xlsx exec file.xlsx --expr "xlsx.readRangeTsv(wb, \"'My Sheet'!A1:J50\")"
```

**Never** embed single-quoted Excel sheet names inside single-quoted bash — it breaks the shell:

```bash
# BROKEN — inner ' terminates the outer bash string:
witan xlsx exec file.xlsx --expr 'xlsx.readRangeTsv(wb, "'My Sheet'!A1:J50")'
```

### Flags

| Flag                 | Short | Default | Description                                               |
| -------------------- | ----- | ------- | --------------------------------------------------------- |
| `--expr`             |       | —       | Expression shorthand; wraps as `return (<expr>);`         |
| `--code`             |       | —       | Inline JavaScript source                                  |
| `--script`           |       | —       | Path to a JavaScript file                                 |
| `--stdin`            |       | —       | Read JavaScript source from stdin                         |
| `--input-json`       |       | `{}`    | JSON value passed as `input` to the script                |
| `--timeout-ms`       |       | `0`     | Execution timeout in milliseconds (must be > 0 if set)    |
| `--max-output-chars` |       | `0`     | Maximum stdout characters to capture (must be > 0 if set) |
| `--json`             |       | `false` | Print the full response envelope as JSON                  |

### Runtime globals

| Name    | Type            | Description                                                         |
| ------- | --------------- | ------------------------------------------------------------------- |
| `xlsx`  | object          | Curated API surface — all functions listed below                    |
| `wb`    | WorkbookContext | The opened workbook handle; pass as first arg to all `xlsx.*` calls |
| `input` | any             | Parsed value from `--input-json` (defaults to `{}`)                 |
| `print` | function        | Output to stdout (like `console.log` but captured in response)      |

Top-level `await` is supported. No imports allowed (static or dynamic).

### API reference

Functions are grouped by purpose. All are async and take `wb` as the first argument.

**Reading**

| Function                | Signature                 | Description                                               |
| ----------------------- | ------------------------- | --------------------------------------------------------- |
| `listSheets`            | `(wb)`                    | List all sheets with used ranges                          |
| `getWorkbookProperties` | `(wb)`                    | Workbook-level metadata                                   |
| `getUsedRange`          | `(wb, sheetName)`         | Used range for a sheet                                    |
| `listDefinedNames`      | `(wb)`                    | All defined names                                         |
| `readCell`              | `(wb, cell, opts?)`       | Read a single cell; `opts.context` adds surrounding cells |
| `readRange`             | `(wb, range)`             | Read all cells in a range                                 |
| `readRow`               | `(wb, sheet, row, opts?)` | Read a row; `opts.startCol/endCol` to limit               |
| `readColumn`            | `(wb, sheet, col, opts?)` | Read a column; `opts.startRow/endRow` to limit            |
| `readRangeTsv`          | `(wb, range, opts?)`      | Read range as TSV text; `opts.includeEmpty`               |
| `readColumnTsv`         | `(wb, sheet, col, opts?)` | Read column as TSV text                                   |
| `readRowTsv`            | `(wb, sheet, row, opts?)` | Read row as TSV text                                      |

**Searching**

| Function       | Signature                                | Description                                                                         |
| -------------- | ---------------------------------------- | ----------------------------------------------------------------------------------- |
| `findCells`    | `(wb, matcher, opts?)`                   | Find cells by value or pattern; `opts.in`, `context`, `limit`, `offset`, `formulas` |
| `findRows`     | `(wb, matcher, opts?)`                   | Find rows by value or pattern; `opts.in`, `context`, `limit`, `offset`              |
| `detectTables` | `(wb)`                                   | Auto-detect table-like regions                                                      |
| `tableLookup`  | `(wb, { table, rowLabel, columnLabel })` | Look up a value by row and column labels                                            |

`matcher` accepts: string, string array (OR match), number, boolean, RegExp, or RegExp array. Searches are fuzzy and case-insensitive by default.

**Tracing**

| Function            | Signature               | Description                                           |
| ------------------- | ----------------------- | ----------------------------------------------------- |
| `getCellPrecedents` | `(wb, address, depth?)` | Cells that feed into this cell; `depth` defaults to 1 |
| `getCellDependents` | `(wb, address, depth?)` | Cells that depend on this cell                        |
| `traceToInputs`     | `(wb, cell)`            | Trace all the way to leaf input cells (no formula)    |
| `traceToOutputs`    | `(wb, cell)`            | Trace all the way to terminal output cells            |

**Computing**

| Function           | Signature               | Description                                  |
| ------------------ | ----------------------- | -------------------------------------------- |
| `evaluateFormula`  | `(wb, sheet, formula)`  | Evaluate a formula string in a sheet context |
| `evaluateFormulas` | `(wb, sheet, formulas)` | Evaluate multiple formulas at once           |

**Writing (ephemeral)**

| Function                | Signature                    | Description                                                            |
| ----------------------- | ---------------------------- | ---------------------------------------------------------------------- |
| `setCells`              | `(wb, cells)`                | Write values/formulas to cells; returns `{ touched, changed, errors }` |
| `scaleRange`            | `(wb, range, factor, opts?)` | Multiply numeric cells by a factor; `opts.skipFormulas` (default true) |
| `insertRowAfter`        | `(wb, sheet, row, count?)`   | Insert rows after a given row                                          |
| `deleteRows`            | `(wb, sheet, row, count?)`   | Delete rows starting at a given row                                    |
| `insertColumnAfter`     | `(wb, sheet, col, count?)`   | Insert columns after a given column                                    |
| `deleteColumns`         | `(wb, sheet, col, count?)`   | Delete columns starting at a given column                              |
| `addSheet`              | `(wb, name)`                 | Add a new sheet                                                        |
| `deleteSheet`           | `(wb, name)`                 | Delete a sheet                                                         |
| `renameSheet`           | `(wb, oldName, newName)`     | Rename a sheet                                                         |
| `addDefinedName`        | `(wb, name, range, scope?)`  | Add a defined name                                                     |
| `setWorkbookProperties` | `(wb, properties)`           | Set workbook-level properties                                          |
| `setSheetProperties`    | `(wb, sheet, properties)`    | Set sheet-level properties (columns, rows, merges, view)               |
| `setStyle`              | `(wb, target, style)`        | Apply styles to a cell or range                                        |

### The ephemeral write contract

`exec` **never writes workbook bytes back to disk**. All write operations (`setCells`, `scaleRange`, inserts, deletes) take effect in the server-side session only. The `result.touched` map contains the recalculated formatted text values — read answers from there.

This means:

- No risk of corrupting the original file
- No `reset()` needed — each invocation starts clean
- Multiple scenarios = multiple `exec` invocations

### setCells result shape

```ts
{
  touched: Record<string, string>  // address → formatted text value
  changed: string[]                // addresses whose values changed
  errors: Diagnostic[]             // cells that errored after recalc
}
```

Read the output value from `result.touched["Sheet!Address"]`. Never compute the answer in JavaScript.

## Error Guide

| Error                                                             | Fix                                                                       |
| ----------------------------------------------------------------- | ------------------------------------------------------------------------- |
| `exactly one of --code, --script, --stdin, or --expr is required` | Provide exactly one code source flag                                      |
| `--code, --script, --stdin, and --expr are mutually exclusive`    | Only use one code source flag per invocation                              |
| `exec code must not be empty`                                     | Provide non-empty code                                                    |
| `Import statements are not allowed`                               | No `import` in exec scripts; use the `xlsx` global                        |
| `EXEC_SYNTAX_ERROR`                                               | Fix JavaScript syntax in your script                                      |
| `EXEC_RUNTIME_ERROR`                                              | Fix runtime error (check the message for details)                         |
| `EXEC_RESULT_TOO_LARGE`                                           | Return less data; use `print()` for large output instead of return values |
| `--timeout-ms must be > 0`                                        | Omit the flag (no timeout) or provide a positive value                    |
| `invalid --input-json`                                            | Provide valid JSON                                                        |
| `Sheet 'X' not found`                                             | Check the sheet name; use `listSheets` to enumerate                       |
| Sheet names with spaces fail in `--expr`                          | See "Shell quoting for sheet names with spaces" above                     |
| `findCells` returns empty                                         | Try synonym arrays, broader search, or check spelling                     |
| `setCells` result missing expected output                         | The output cell may not be a dependent; trace the formula chain           |
