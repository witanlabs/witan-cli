---
name: pptx-code-mode
description: Use this skill when a PPTX file needs to be rendered, inspected, created, or modified through Witan PPTX. The tool runs sandboxed Office.js-compatible JavaScript against PPTX files via `witan pptx exec`.
---

## Quick Reference

Render a slide:

```bash
witan pptx render deck.pptx --slide 1 -o slide-1.png
```

Run Office.js-compatible JavaScript:

```bash
witan pptx exec deck.pptx --stdin <<'JS'
return await PowerPoint.run(async context => {
  return context.presentation.slides.getCount().value;
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
  const slide = slides.getItemAt(slides.getCount().value - 1);
  slide.shapes.addTextBox("Created", { left: 60, top: 60, width: 300, height: 80 });
  await context.sync();
  return slides.getCount().value;
});
JS
```

## Guidance

- Use `--stdin` with a quoted heredoc for multi-line scripts.
- Use `--save` only when changes should be written back.
- Use `--json` when automation needs the full response envelope.
- Use `--input-json` to pass structured data as the `input` global.
- Use `witan pptx render --diff baseline.png` for visual regression checks.

Top-level `await` is supported. Static and dynamic imports are not available.

## Office.js Reference

Office.js declarations are available in `references/office-js.d.ts`.

- Prefer targeted lookup with `rg`, for example `rg -n "SlideCollection|addTextBox|ClientResult" references/office-js.d.ts`.
- `PowerPoint.createPresentation(...)` is intentionally not implemented; use `witan pptx exec <file> --create --save` to create a PPTX file through the CLI.

Witan PPTX follows the Office.js PowerPoint API surface. `OfficeExtension.ClientResult` values must be read after `await context.sync()`, and APIs returning `void` should not be treated as object factories.
