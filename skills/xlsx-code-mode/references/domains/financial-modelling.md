# Financial modelling conventions

Read this for any financial model, valuation, projection, or investment-banking workbook. These conventions sit **on top of** the universal quality floor in `SKILL.md`. They are the defaults bankers and analysts expect — if the user supplies a template or house style, that always wins.

## Colour coding (font colour by cell role)

The single most recognisable convention — get it right and the model reads as professional.

| Role | Colour | Hex |
|---|---|---|
| Hardcoded inputs / assumptions a user will change | **Blue** | `#0000FF` |
| Formulas & calculations | **Black** | `#000000` |
| Links to another sheet in the same workbook | **Green** | `#008000` |
| Links to an external file | **Red** | `#FF0000` |
| Key assumption needing attention (this one is a *fill*, not font) | **Yellow** | `#FFFF00` |

## Number formats

- **Zeros show as `-`** (including percentages and currency) — e.g. `#,##0;(#,##0);-`
- **Negatives in red parentheses**, never a minus sign — e.g. `$#,##0_);[Red]($#,##0)`
- **Multiples** as `0.0x` (e.g. `5.2x`) — format `0.0"x"`
- **Percentages** default to one decimal — `0.0%`
- **Years** are text labels, not numbers — `2024`, never `2,024` — write `'2024` (a plain `"2024"` coerces to a number, even with format `@`)
- **Currency** uses `$#,##0`; always state units in the header — `Revenue ($mm)`

## Layout

- **Totals sum the range directly above them**, with a thin top border spanning the label and data columns.
- **Hide gridlines** — `setSheetProperties(wb, sheet, { view: { showGridLines: false } })`.
- **Section headers** spanning columns: merged cell, dark-blue/black fill, white bold text, left-aligned.
- **Column labels** (dates, periods) and the numbers beneath them are **right-aligned**. **Row labels** are left-aligned; sub-metrics (e.g. `% growth`) are left-aligned and indented.

## Sourcing

- Every hardcoded input cites its source in a cell comment: `Source: [System/Doc], [Date], [Reference], [URL]` — e.g. `Source: Company 10-K, FY2024, p.45`.
- Raw inputs and assumptions live in dedicated cells/sheets, never inside formulas.

## Worked example

A revenue build that follows every rule above — blue sourced inputs, black formulas referencing a separated growth assumption, finance number formats, a bordered total, merged dark header, gridlines off:

```bash
witan xlsx exec model.xlsx --create --save --stdin <<'WITAN'
await xlsx.addSheet(wb, "Model")
await xlsx.setSheetProperties(wb, "Model", { view: { showGridLines: false }, merges: ["A1:D1"] })

const num = "#,##0;(#,##0);-"

// Section header — merged, dark fill, white bold, left-aligned
await xlsx.setCells(wb, [{ address: "Model!A1", value: "Revenue Build ($000s)" }])
await xlsx.setStyle(wb, "Model!A1:D1", {
  fill: { color: "#1F3864" },
  font: { color: "#FFFFFF", bold: true },
  alignment: { horizontal: "left" },
})

// Period labels — text years ('-prefix), right-aligned, bold
await xlsx.setCells(wb, [
  { address: "Model!B2", value: "'2024" },
  { address: "Model!C2", value: "'2025" },
  { address: "Model!D2", value: "'2026" },
])
await xlsx.setStyle(wb, "Model!B2:D2", { font: { bold: true }, alignment: { horizontal: "right" } })

// Revenue lines — base year is a BLUE sourced input; later years are BLACK formulas
await xlsx.setCells(wb, [
  { address: "Model!A3", value: "Product" },
  { address: "Model!B3", value: 600, format: num, note: { text: "Source: Company 10-K, FY2024, p.45" } },
  { address: "Model!C3", formula: "=B3*(1+$B$8)", format: num },
  { address: "Model!D3", formula: "=C3*(1+$B$8)", format: num },
  { address: "Model!A4", value: "Services" },
  { address: "Model!B4", value: 400, format: num, note: { text: "Source: Company 10-K, FY2024, p.45" } },
  { address: "Model!C4", formula: "=B4*(1+$B$8)", format: num },
  { address: "Model!D4", formula: "=C4*(1+$B$8)", format: num },
])
await xlsx.setStyle(wb, "Model!B3:B4", { font: { color: "#0000FF" } }) // blue inputs

// Total — sums the range directly above, with a top border
await xlsx.setCells(wb, [
  { address: "Model!A5", value: "Total revenue" },
  { address: "Model!B5", formula: "=SUM(B3:B4)", format: num },
  { address: "Model!C5", formula: "=SUM(C3:C4)", format: num },
  { address: "Model!D5", formula: "=SUM(D3:D4)", format: num },
])
await xlsx.setStyle(wb, "Model!A5:D5", { border: { top: { style: "thin", color: "#000000" } } })

// Assumption — separated from the calc, blue input, one-decimal percent
await xlsx.setCells(wb, [
  { address: "Model!A8", value: "Growth rate" },
  { address: "Model!B8", value: 0.12, format: "0.0%" },
])
await xlsx.setStyle(wb, "Model!B8", { font: { color: "#0000FF" } })

// Widen the label column so nothing clips
await xlsx.autoFitColumns(wb, "Model", ["A"])

return await xlsx.readRangeTsv(wb, { sheet: "Model", from: {row:1,col:1}, to: {row:8,col:4} }, { includeFormulas: true })
WITAN
```

Then verify, as always: `witan xlsx calc model.xlsx` (must be `0 errors`), `witan xlsx lint model.xlsx` (must exit clean), and `witan xlsx render model.xlsx -r "Model!A1:D8"` to confirm the look.
