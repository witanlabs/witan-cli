# Changelog

For spreadsheet product and runtime changes, see the [spreadsheet changelog](https://docs.witanlabs.com/products/spreadsheet/changelog).

## Unreleased

- New: [Skill] `xlsx-mcp` brings the `xlsx-code-mode` read/author/what-if/verify workflow to agents using the Witan MCP server instead of the CLI — `xlsx_exec` scripting plus the presigned `prepare_*` file round-trip. Requires the server's merged `xlsx_exec` (create-by-`filename`) to be deployed.
- New: [Skill] All skills now declare a per-skill semver in SKILL.md frontmatter (`metadata.version`, starting at 1.0.0). CI fails a PR that changes a skill's files without bumping its version; skill changelog entries now include the version. See `skills/README.md`.
- New: [Skill] `xlsx-excelscript` drives `witan xlsx exec` with the Office Scripts (ExcelScript) dialect — the same read/author/what-if/verify workflow as `xlsx-code-mode`, written in `function main(workbook)` style behind the `// @office-script` pragma. Sibling to `xlsx-code-mode` for comparing the two dialects.
- Updated: [Skill] `pptx-code-mode` now documents Witan install steps for the Claude Cowork sandbox.

## 0.12.0

- Updated: [CLI] `witan xlsx exec` now accepts TypeScript as well as JavaScript for `--code`, `--script`, and `--stdin` sources.
- New: [CLI] `witan xlsx exec` and `witan pptx exec` now accept repeatable `--input-file key=@path` flags for passing local PNG/JPEG files to scripts as data URI input values.
- Updated: [CLI] `witan xlsx lint` now describes itself as semantic workbook checking and lists rule `D043` for data validation violations.
- Breaking: [CLI] [JS SDK] [Python SDK] [Skill] `validateCells` was removed from the xlsx method surface; data validation failures are now reported through lint rule `D043`.
- Updated: [JS SDK] [Python SDK] xlsx types now cover richer workbook, sheet, style, chart, image, data validation, ListObject, and What-If Data Table payloads.
- Updated: [Skill] `xlsx-code-mode` now documents `--input-file`, data validation linting, and expanded chart/image authoring surfaces including radar, surface, histogram, Pareto, funnel, and box-whisker charts.

## 0.11.0

- New: [CLI] `witan pptx exec` runs Office.js-compatible scripts against PPTX files, with `--create`, `--save`, `--json`, `--input-json`, `--locale`, timeout, and output-size controls.
- New: [CLI] `witan pptx render` renders a slide to PNG and can produce visual diffs against a baseline PNG.
- New: [Skill] `pptx-code-mode` provides the workflow and Office.js reference guidance for inspecting, creating, editing, rendering, and visually checking PPTX files.

## 0.10.1

- New: [JS SDK] Introduced the Node.js SDK, with `Workbook.open`, async disposal, binary discovery, typed errors, request timeouts, and xlsx workbook methods backed by `witan xlsx rpc`.
- Fixed: [CLI] Expired saved auth sessions are cleared automatically after `401` or `403` token exchange failures, allowing unauthenticated stateless usage to resume.
- Updated: [Skill] `xlsx-code-mode` and `read-source` now document the `npx witan` fallback when `witan` is not already on `PATH`.

## 0.10.0

- New: [Python SDK] The `witan` PyPI package now ships a Python SDK alongside the bundled CLI binary, exposing `witan.Workbook` and `witan.AsyncWorkbook` for synchronous and asyncio access to xlsx workbook sessions over `witan xlsx rpc`.
- New: [CLI] `witan xlsx rpc <file>` opens a persistent xlsx session and relays newline-delimited JSON requests over stdio, enabling low-latency multi-op workflows without re-uploading the workbook per call.
- Updated: [CLI] Cached `witan xlsx exec` sessions now pin to a backend instance via affinity cookies, improving warm-cache hit rates across sequential calls against the same workbook.
- Updated: [CLI] `witan xlsx exec` cache keys now include the local file path, preventing collisions between distinct workbooks that share identical content but live at different paths.

## 0.9.0

- New: [CLI] `witan auth status` reports the active credential source, validation state, selected organization, ignored lower-priority credentials, and supports `--json` output.
- Updated: [CLI] `witan xlsx lint` help no longer lists removed rule `D031`; examples now use active rule IDs such as `D001` and `D030`.

## 0.8.0

No CLI, JS SDK, Python SDK, or skill surface changes.

## 0.7.1

- Fixed: [CLI] `witan xlsx exec --locale` now sends the locale in the request query as well as the request payload/header, so locale-sensitive execution is honored across stateless and files-backed modes.

## 0.7.0

- New: [CLI] `witan xlsx exec --create` creates and populates a workbook in a single command.
- New: [CLI] `witan xlsx exec --locale` controls the locale used for formula calculation, number formatting, and string comparison; the same value can be supplied through `WITAN_LOCALE`.

## 0.6.2

- New: [CLI] `--version` now prints the API version in addition to the CLI version.
- Fixed: [CLI] `witan xlsx exec --json --save` writes returned workbook bytes locally but omits the transport-only base64 file payload from JSON output.
- Fixed: [CLI] Temporary image files emitted by `witan xlsx exec` now use owner-only file permissions.

## 0.6.1

- Fixed: [CLI] File uploads now use stable built-in MIME type detection for spreadsheet and document extensions, avoiding host-dependent MIME sniffing issues.

## 0.6.0

- New: [CLI] Authenticated API calls are now org-scoped (`/v0/orgs/:org_id/...`), enabling multi-org support.
- New: [CLI] `witan auth login` prompts for org selection when the user belongs to multiple organizations.
- Updated: [CLI] File and revision cache keys now include org ID to prevent cross-org collisions.

## 0.5.0

No CLI, JS SDK, Python SDK, or skill surface changes.

## 0.4.0

- Fixed: [CLI] README and command help examples now use the documented `xlsx.readCell(wb, "Summary!A1")` exec API instead of the obsolete `wb.sheet(...).cell(...)` style.
- Updated: [Skill] `xlsx-code-mode` now documents `includeFormulas` for TSV reads and expands its workbook scripting reference for the method surface available through `witan xlsx exec`.

## 0.3.1

- New: [CLI] `witan read` extracts text and structure from supported source documents for downstream workbook and document workflows.
- Updated: [CLI] `witan xlsx exec` now writes image outputs to temporary files and prints their paths.
- New: [CLI] `witan xlsx exec --stdin-timeout-ms` prevents stalled stdin reads when an input stream never reaches EOF.
- Updated: [CLI] API requests now include the CLI version in the `User-Agent` header.
- Fixed: [CLI] XLSX calculation and execution cache updates now preserve the resolved local workbook path.

## 0.3.0

- [CLI] `witan xlsx calc`: Recalculates workbook formulas, surfaces formula errors, and can run in `--verify` mode to detect value drift without mutating the file.
- [CLI] `witan xlsx exec`: Runs sandboxed JavaScript against a workbook to read, search, trace, and (optionally) persist edits.
- [CLI] `witan xlsx lint`: Performs semantic formula analysis to catch spreadsheet risks like double counting, lookup issues, and coercion surprises.
- [CLI] `witan xlsx render`: Renders a sheet range to an image for visual layout, formatting, and diff-based change checks.

- [Skill] `xlsx-code-mode`: Workbook exploration/editing workflow centered on `witan xlsx exec` for reading data, tracing logic, and running what-if updates.
