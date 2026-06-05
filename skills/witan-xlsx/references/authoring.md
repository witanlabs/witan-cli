# Authoring high-quality workbooks

This is the craft layer. The [SKILL.md](../SKILL.md) tells you how to drive the engine; this tells you what good output looks like and how to reach it with the Witan API. For financial models, also read [financial-models.md](financial-models.md).

A workbook is good when three things are true: it **computes correctly**, it is **structured so a human can follow the logic**, and it is **formatted so values read clearly**. Everything below serves one of those three.

---

## 1. Modelling discipline

**Derived values are formulas, not hardcoded.** If a number can be computed from other cells, write the formula, not the answer — `=SUM(B2:B9)`, `=C4/C2-1`, `=AVERAGE(D2:D19)` — never the precomputed result. The whole value of a spreadsheet is that it recalculates; a hardcoded total is a screenshot.

**No magic numbers inside formulas.** Every assumption — a growth rate, tax rate, price, multiple — lives in its own labelled cell and is referenced. Write `=C7*(1+$C$3)`, not `=C7*1.08`. A reader should be able to change one input cell and watch the model respond.

**Calculate once, then link.** Compute a quantity in exactly one place. When you need it again, reference that source cell (`=Calc!B12`) rather than re-deriving it with a second formula that can drift out of step after an input changes. Point each link at the canonical source, never at a cell that merely links to it.

**Get references right for fill.** Anchor what shouldn't move when a formula is copied: `=B5*$B$6` fills across columns while still pointing at the single rate in `$B$6`. Mixed anchors (`$B6`, `B$6`) matter in grids. Keep a row's formula identical across every period so column 12 behaves exactly like column 2 — inconsistent formulas across a projection are the most common modelling bug.

**Name things that are referenced widely.** `addDefinedName(wb, "TaxRate", "Assumptions!$B$3")` makes `=Pretax*TaxRate` readable and stable. Use named ranges for key drivers, not for every cell.

**A cell whose visible text begins with `=`, `+`, or `-`** is read as a formula. To keep it as text (e.g. a label like `=high`), write it as a plain string value rather than a `formula`.

---

## 2. Structure and layout

**Separate inputs, calculations, and outputs.** The clearest models put assumptions/inputs in one place (often their own sheet or a top block), calculations in the middle, and presentation/outputs (summaries, charts) where they're read. A reader should always know whether a cell is something they can change or something the model derived.

**One concept per sheet, in reading order.** Don't crowd unrelated tables onto one sheet. Lay each sheet top-to-bottom / left-to-right in the order someone would think about it. Use `listSheets` and `describeSheet` to check an existing layout before adding to it.

**Give it room.** Leave a blank row/column between sections. Dense, wall-to-wall grids are hard to scan. Whitespace is free and does more for readability than colour.

**Section headers.** Label each block with a heading cell (bold, larger, or filled). For a heading that spans columns, either merge (`setSheetProperties` `merges`) or use `centerContinuousSpan` to centre across without merging (which keeps the cells independently addressable).

**Column widths and row heights.** Size columns to their content: `autoFitColumns(wb, sheet)` then cap anything extreme. Very wide text columns should be bounded (set an explicit `setColumnProperties` width and turn on `wrapText` in the style) rather than stretched to the longest string. Give title rows a little more height with `setRowProperties`.

**Alignment carries meaning.** Right-align numbers and their column headers; left-align text labels; indent sub-items beneath their parent (`alignment.indent`). Consistent alignment lets the eye separate labels from values instantly.

**Gridlines and view.** For a presentation/summary sheet, turning gridlines off (`setSheetProperties` `view.showGridLines: false`) and relying on deliberate borders looks far more finished. Keep gridlines on for working/data sheets.

---

## 3. Number formats

Formatting is not decoration — an unformatted `0.0834` and a `1234567` are unreadable. Apply a number format to every numeric cell (via `setCells` `format`, or `setStyle` `numberFormat`). A practical cookbook:

| Intent                    | Format string          | Renders      |
| ------------------------- | ---------------------- | ------------ |
| Whole currency            | `"$#,##0"`             | `$1,235`     |
| Currency, 2dp             | `"$#,##0.00"`          | `$1,234.57`  |
| Thousands, no symbol      | `"#,##0"`              | `1,235`      |
| Percent, 1dp              | `"0.0%"`               | `8.3%`       |
| Multiple                  | `"0.0\"x\""`           | `5.2x`       |
| Negatives in red parens   | `"#,##0;[Red](#,##0)"` | `(1,235)`    |
| Zero shown as dash        | `"#,##0;-#,##0;\"-\""` | `-`          |
| Date                      | `"yyyy-mm-dd"`         | `2026-06-05` |
| Plain text (don't coerce) | `"@"`                  | as typed     |

Conventions worth defaulting to:

- **Put units in the header**, not the cells: a column titled `Revenue ($mm)` with bare numbers beats `$` on every row.
- **Negatives and zeros are choices.** Decide whether negatives show as `(123)` and whether zeros show as `-`; apply consistently down a column.
- **Dates must be real dates** with a date format, never strings — otherwise sorting, charting, and date maths break. Years used as labels (`2024`, `2025`) are the exception: keep them as text/labels so they don't get thousands separators.
- **Don't store numbers as text.** A number entered as `"1,200"` won't sum. Write the numeric value and format it.
- **Consistent decimals** within a column.

---

## 4. Styling

**One professional font** across the workbook (set `defaultFont` via `setWorkbookProperties`, e.g. a clean sans-serif). Don't mix faces.

**Style the header row** of every table: bold, a fill, white text on a dark fill if you like, centred. This single move makes a sheet look intentional.

**Borders, sparingly.** A light bottom border under a header row and a border above a total is usually all you need. Avoid boxing every cell. `setStyle` `border` takes `top/bottom/left/right` with a `style` (`thin`/`medium`/`thick`) and `color`.

**Colour and fill, sparingly.** Use fill to separate input cells from calculated cells and to flag headers — not to paint the whole sheet. Pull from the workbook theme (`themeColors`) so it stays coherent. (Financial models have specific colour conventions — see [financial-models.md](financial-models.md).)

**Merges only where they earn it** (titles, banners). Merged cells complicate selection, sorting, and references; prefer `centerContinuousSpan` for spanning a heading visually without merging.

---

## 5. Tables, conditional formatting, and signal

**Use real Excel tables for tabular data.** `addListObject` gives a named table with header styling, banding, a totals row, and structured references (`=SUM(Sales[Amount])`) that auto-expand as rows are added. Read one back with `readRange(wb, "TableName")`. Give every table a unique, descriptive name.

**Use conditional formatting for signal, not noise.** A two/three-colour scale on a metric column, a data bar, an icon set, or a rule that turns negatives red (`setConditionalFormatting`) helps the reader find what matters — and is the right tool for flagging out-of-range or out-of-list values. Don't formatting-spam.

---

## 6. Charts

- **Pick the type that fits the question:** trends over time → line; composition → stacked column or pie (sparingly); comparison across categories → column/bar; correlation → scatter.
- **Always title the chart and label axes**, including units. A chart without units is a guess.
- **Anchor it where it doesn't cover data** — position it in empty space beside or below the table it visualises (`addChart` `position` with `from`/`to` anchors).
- **Point series at ranges**, so the chart updates with the data; don't bake values in.
- **Verify it visually.** After `addChart`/`setChart`, `render` the chart's range (or `previewStyles` in-script) and actually look at placement, labels, and legend.

---

## 7. Verification — the standard is zero errors

Authoring isn't finished until it's verified. Run this loop, fix, repeat:

1. **`xlsx calc <file> --verify`** — non-mutating recalculation; exit code **2** if any formula error exists _or_ any computed value changed. The definitive "is it consistent" check.
2. **`xlsx lint <file>`** — semantic and presentation problems a recalc can't see. A representative set (the linter ships more; run `witan xlsx lint --help` for the current list):

   | Rule | Severity | Catches                                                          |
   | ---- | -------- | ---------------------------------------------------------------- |
   | D001 | Warning  | Double counting from overlapping ranges                          |
   | D002 | Warning  | Approximate-match `*LOOKUP`/`MATCH` on an unsorted range         |
   | D003 | Warning  | Empty cells coerced to 0/FALSE                                   |
   | D005 | Warning  | Aggregates silently ignoring text/booleans                       |
   | D006 | Warning  | Unintended scalar broadcast in elementwise ops                   |
   | D007 | Warning  | Lookup with duplicate keys returns first match                   |
   | D008 | Error    | Mixed currencies added/aggregated together                       |
   | D009 | Warning  | Percent mixed with non-percent in +/-                            |
   | D023 | Warning  | Currency mixed with non-currency semantic formats                |
   | D030 | Warning  | Formula references a non-anchor cell of a merged range           |
   | D032 | Warning  | Text display clipped by a non-empty neighbour (widen the column) |

   Scope with `-r RANGE`, focus with `--only-rule`, or silence a checked false-positive with `--skip-rule`.

3. **`xlsx render <file> -r RANGE`** — look at the actual output: clipped text, blown-out widths, broken merges, charts over data, misaligned numbers. `--diff baseline.png` highlights what an edit changed.
4. **In-script assertions** — `evaluateFormula(wb, sheet, "=...")` to assert invariants (a check row sums to zero, a balance sheet balances).
5. **Test the unhappy branch** — an error can hide where the formula isn't currently looking. `=IF(flag=1, x, 1/divisor)` reports no error while `flag=1`, and neither `calc --verify` nor `lint` sees the broken `1/divisor` branch — it only bites when the condition flips. Where an `IF` guards a bad case (a missing lookup, a divide-by-zero), flip the guard with `sweepInputs` or a temporary `setCells` and re-run `calc --verify` with the other branch active.

Don't declare success on an unverified workbook. "It probably calculates" is not done.

---

## 8. Editing an existing workbook

Most real tasks are edits, and the prime directive is **do no harm to what's already there**.

- **Read before you write.** `listSheets` → `describeSheet` → `readRangeTsv` the area; `render` it so you can see the house style.
- **Match the existing conventions.** Read a neighbouring cell's `getStyle` and number format and reuse them. The workbook's established fonts, colours, and layout always win over your defaults — don't restyle unless asked.
- **Make the smallest change that works,** and keep new formulas identical in shape to the rows around them.
- **Edit structure carefully.** `insertRowAfter`/`insertColumnAfter` shift references for you; still re-verify, because inserts/deletes are where `#REF!` and off-by-one errors appear.
- **Re-verify and diff.** `calc --verify`, `lint`, and a `render --diff` against a before image confirm you changed only what you meant to.

---

## 9. Provenance

Numbers people can't trace, they can't trust. For any **hardcoded input**, attach where it came from to the cell itself — a note or a hyperlink via `setCells` — so the provenance travels with the value:

```js
await xlsx.setCells(wb, [
  {
    address: "Assumptions!B4",
    value: 0.21,
    format: "0.0%",
    note: { text: "US federal statutory corporate rate, FY2025 (IRS)." },
  },
]);
```

The point is that a reviewer can land on any input and immediately see its origin without hunting through a separate document.

---

## 10. Efficiency (latency matters)

The engine round-trips per `exec` call, so:

- **Batch.** One `setCells` with many cells beats many calls. Build a sheet in as few invocations as you can.
- **Explore for free.** Reads and what-ifs are ephemeral — no `--save`, no risk — so prototype layout and validate formulas before committing.
- **Verify in one pass.** Lint and calc the whole workbook once near the end rather than after every tiny edit.

---

## Error codes (quick reference)

- **File/context:** `FILE_NOT_FOUND`, `CONTEXT_CLOSED`
- **Sheets:** `SHEET_NOT_FOUND`, `DUPLICATE_SHEET`
- **Address/range:** `ADDRESS_PARSE_ERROR`, `RANGE_PARSE_ERROR`, `ADDRESS_MISSING_SHEET`, `RANGE_MISSING_SHEET`, `ADDRESS_TYPE_MISMATCH`, `ADDRESS_INVALID_COORDINATES`, `RANGE_INVALID_COORDINATES`, `COLUMN_OUT_OF_BOUNDS`, `ROW_OUT_OF_BOUNDS`
- **Data/style:** `DATA_SHAPE_MISMATCH` (array shape ≠ target range), `INVALID_COLOR`, `MISSING_VALUE`, `AMBIGUOUS_VALUE`

Every string address needs a sheet qualifier (`"Sheet1!A1"`). When writing a 2-D block, the array shape must match the target range exactly.
