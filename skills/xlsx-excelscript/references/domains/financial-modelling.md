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
- **Hide gridlines** — `sheet.setShowGridlines(false)`.
- **Section headers** spanning columns: merged cell, dark-blue/black fill, white bold text, left-aligned.
- **Column labels** (dates, periods) and the numbers beneath them are **right-aligned**. **Row labels** are left-aligned; sub-metrics (e.g. `% growth`) are left-aligned and indented.

## Sourcing

- Every hardcoded input cites its source in a cell comment via `workbook.addComment(address, "Source: …")`: `Source: [System/Doc], [Date], [Reference], [URL]` — e.g. `Source: Company 10-K, FY2024, p.45`.
- Raw inputs and assumptions live in dedicated cells/sheets, never inside formulas.

## Worked example

A revenue build that follows every rule above — blue sourced inputs, black formulas referencing a separated growth assumption, finance number formats, a bordered total, merged dark header, gridlines off:

```bash
witan xlsx exec model.xlsx --create --save --stdin <<'JS'
// @office-script
function main(workbook) {
  const sheet = workbook.addWorksheet("Model");
  sheet.setShowGridlines(false);

  const num = "#,##0;(#,##0);-";

  // Section header — merged, dark fill, white bold, left-aligned
  sheet.getRange("A1").setValue("Revenue Build ($000s)");
  const header = sheet.getRange("A1:D1");
  header.merge();
  header.getFormat().getFill().setColor("#1F3864");
  header.getFormat().getFont().setColor("#FFFFFF");
  header.getFormat().getFont().setBold(true);
  header.getFormat().setHorizontalAlignment("Left");

  // Period labels — text years ('-prefix), right-aligned, bold
  const periods = sheet.getRange("B2:D2");
  periods.setValues([["'2024", "'2025", "'2026"]]);
  periods.getFormat().getFont().setBold(true);
  periods.getFormat().setHorizontalAlignment("Right");

  // Revenue lines — base year is a BLUE sourced input; later years are BLACK formulas
  sheet.getRange("A3").setValue("Product");
  sheet.getRange("B3").setValue(600);
  sheet.getRange("C3").setFormula("=B3*(1+$B$8)");
  sheet.getRange("D3").setFormula("=C3*(1+$B$8)");
  sheet.getRange("A4").setValue("Services");
  sheet.getRange("B4").setValue(400);
  sheet.getRange("C4").setFormula("=B4*(1+$B$8)");
  sheet.getRange("D4").setFormula("=C4*(1+$B$8)");
  sheet.getRange("B3:D4").setNumberFormat(num);
  sheet.getRange("B3:B4").getFormat().getFont().setColor("#0000FF"); // blue inputs
  workbook.addComment("Model!B3", "Source: Company 10-K, FY2024, p.45");
  workbook.addComment("Model!B4", "Source: Company 10-K, FY2024, p.45");

  // Total — sums the range directly above, with a thin top border
  sheet.getRange("A5").setValue("Total revenue");
  sheet.getRange("B5").setFormula("=SUM(B3:B4)");
  sheet.getRange("C5").setFormula("=SUM(C3:C4)");
  sheet.getRange("D5").setFormula("=SUM(D3:D4)");
  sheet.getRange("B5:D5").setNumberFormat(num);
  const topBorder = sheet.getRange("A5:D5").getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop);
  topBorder.setStyle(ExcelScript.BorderLineStyle.continuous);
  topBorder.setWeight(ExcelScript.BorderWeight.thin);
  topBorder.setColor("#000000");

  // Assumption — separated from the calc, blue input, one-decimal percent
  sheet.getRange("A8").setValue("Growth rate");
  sheet.getRange("B8").setValue(0.12);
  sheet.getRange("B8").setNumberFormat("0.0%");
  sheet.getRange("B8").getFormat().getFont().setColor("#0000FF");

  sheet.getRange("A3:A8").getFormat().autofitColumns();

  // Read the recalculated values + formulas back to confirm
  return {
    values: sheet.getRange("A1:D8").getValues(),
    formulas: sheet.getRange("A1:D8").getFormulas(),
  };
}
JS
```

Then verify, as always: `witan xlsx calc model.xlsx` (must be `0 errors`), `witan xlsx lint model.xlsx` (must exit clean), and `witan xlsx render model.xlsx -r "Model!A1:D8"` to confirm the look.
