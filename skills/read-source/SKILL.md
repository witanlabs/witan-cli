---
name: read-source
description: Use this skill to extract text content from non-spreadsheet documents (PDF, DOCX, PPTX, HTML, text). Trigger when you need to read, search, or understand the contents of a PDF report, Word document, PowerPoint deck, HTML page, or plain-text file. The tool extracts text server-side via `witan read`.
---

## Quick Reference

```bash
witan read <file>                    # Full text extraction
witan read <file> --outline          # Document structure only
witan read report.pdf --pages 1-5    # Specific PDF pages
witan read slides.pptx --slides 1-3  # Specific slides
witan read notes.docx --offset 50 --limit 100  # Paginate long output
```

## Supported Formats

| Format | Extensions |
|--------|-----------|
| PDF | `.pdf` |
| Word | `.doc`, `.docx` |
| PowerPoint | `.ppt`, `.pptx` |
| HTML | `.html`, `.htm` |
| Text | `.txt`, `.md`, `.csv`, `.json`, `.xml`, `.yaml`, `.toml` |

## Flags

| Flag | Description |
|------|-------------|
| `--outline` | Show document structure instead of full content |
| `--pages <range>` | PDF page range (e.g. `1-5`, `1,3,5`) |
| `--slides <range>` | Presentation slide range (e.g. `1-3`) |
| `--offset <n>` | Start line (1-indexed) |
| `--limit <n>` | Max lines to return |
| `--json` | Output full JSON response |

## Strategy

1. **Start with `--outline`** for large documents to understand the structure.
2. **Target specific sections** with `--pages` or `--slides` to reduce output size.
3. **For small documents**, read the full text directly (no flags needed).
4. **Paginate** with `--offset` and `--limit` when output is very long.

## Scope

This skill is for non-spreadsheet documents only. For Excel files (`.xlsx`, `.xlsm`), use the `xlsx-code-mode` or `xlsx-verify` skills.
