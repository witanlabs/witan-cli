---
name: witan-pptx-officejs
description: Use this skill when a PPTX file needs to be rendered, inspected, created, or modified through Witan PPTX. The tool runs sandboxed Office.js-compatible JavaScript plus Witan PPTX chart extensions against PPTX files via `witan pptx exec`.
license: Apache-2.0
metadata:
  version: "1.0.0"
  author: witanlabs
  source: https://github.com/witanlabs/witan-cli
---

> **Running in Claude Cowork?** The `witan` CLI isn't preinstalled — see [references/cowork-setup.md](references/cowork-setup.md) for install steps.

## Quick Reference

Render a slide:

```bash
witan pptx render deck.pptx --slide 1 -o slide-1.png
```

Run Office.js-compatible JavaScript:

```bash
witan pptx exec deck.pptx --stdin <<'JS'
return await PowerPoint.run(async context => {
  const count = context.presentation.slides.getCount();
  await context.sync();
  return count.value;
});
JS
```

Save changes:

```bash
witan pptx exec deck.pptx --save --stdin <<'JS'
return await PowerPoint.run(async context => {
  const slide = context.presentation.slides.getItemAt(0);
  slide.shapes.addTextBox("Updated", { left: 60, top: 60, width: 300, height: 80 });
  await context.sync();
  return true;
});
JS
```

Create a new deck:

```bash
witan pptx exec new.pptx --create --save --stdin <<'JS'
return await PowerPoint.run(async context => {
  const slides = context.presentation.slides;
  slides.add();
  const count = slides.getCount();
  await context.sync();
  const slide = slides.getItemAt(count.value - 1);
  slide.shapes.addTextBox("Created", { left: 60, top: 60, width: 300, height: 80 });
  await context.sync();
  return count.value;
});
JS
```

Create and style a chart:

```bash
witan pptx exec charts.pptx --create --save --stdin <<'JS'
return await PowerPoint.run(async context => {
  const slide = context.presentation.slides.getItemAt(0);
  const shape = slide.shapes.addChart(PowerPoint.ChartType.columnClustered, [
    ["Quarter", "Revenue", "Margin"],
    ["Q1", 12, 0.31],
    ["Q2", 18, 0.34],
    ["Q3", 25, 0.39]
  ], { left: 48, top: 72, width: 420, height: 240, name: "Revenue Chart", seriesBy: "columns" });

  const chart = shape.getChart();
  chart.title.text = "Revenue and Margin";
  chart.legend.position = "Top";
  chart.axes.valueAxis.title.text = "Revenue";
  chart.series.getItemAt(1).chartType = PowerPoint.ChartType.lineMarkers;
  chart.series.getItemAt(1).axisGroup = PowerPoint.ChartAxisGroup.secondary;
  await context.sync();
  return chart.name;
});
JS
```

## Guidance

- Use `--stdin` with a quoted heredoc for multi-line scripts.
- Use `--save` only when changes should be written back.
- Use `--create` with a new `.pptx` path to run against an empty PPTX file; add `--save` to write the created file locally.
- Use `--json` when automation needs the full response envelope.
- Use `--input-json` to pass structured data as the `input` global.
- Use `witan pptx render deck.pptx --slide 1 --diff baseline.png` for visual regression checks.
- After chart authoring or mutation, render the affected slide with `witan pptx render` to verify visual output.

Top-level `await` is supported. Static and dynamic imports are not available.

## Chart Extensions

Witan PPTX includes an Excel-style chart API adapted to PPTX. This is a Witan extension: upstream Office.js PowerPoint declarations do not include these chart APIs.

- Create charts with `slide.shapes.addChart(PowerPoint.ChartType.columnClustered, values, options)`.
- Access charts with `shape.getChart()` or `shape.getChartOrNullObject()`.
- Mutate chart titles, legends, axes, series, data labels, fills, borders, and fonts with property-style accessors such as `chart.title.text`, `chart.axes.valueAxis.maximum`, and `chart.series.getItemAt(0).name`.
- Replace chart data with `chart.setData(values, options)` or repoint embedded workbook ranges with `chart.setDataRange("Sheet1!A1:B4", { seriesBy: "columns" })`.
- Delete charts with `chart.delete()` or delete the containing shape with `shape.delete()`.
- Use `references/witan-pptx-chart.d.ts` as the authoritative reference for Witan-only chart members and enum values.

## References

Office.js declarations are available in `references/office-js.d.ts`. Witan PPTX chart extension declarations are available in `references/witan-pptx-chart.d.ts`.

- Prefer targeted lookup with `rg`, for example `rg -n "SlideCollection|addTextBox|ClientResult|addChart|ChartSeries" references`.
- `PowerPoint.createPresentation(...)` is intentionally not implemented; use `witan pptx exec <file> --create --save` to create a PPTX file through the CLI.

Witan PPTX follows the Office.js PowerPoint API surface where implemented and extends it with chart APIs. `OfficeExtension.ClientResult` values must be read after `await context.sync()`, and APIs returning `void` should not be treated as object factories.
