---
name: witan-gsheets
description: "Read, explore, create, and modify Google Sheets spreadsheets via the witan CLI. You cannot read a Google Sheet with normal file tools — running JavaScript against it via `witan gsheets exec` is the only way to open or inspect it. Trigger whenever a Google Sheet is referenced — a `gs://SPREADSHEET_ID` ref, a `https://docs.google.com/spreadsheets/d/...` URL, or a request to read/build/edit a sheet 'in Google Sheets'/'on Drive'. Also trigger when a `gsheets` command fails with an authorization error so you can run the connect/authorize flow. For Excel files (.xlsx/.xls/.xlsm) on disk, use the witan-xlsx skill instead — the in-sheet JavaScript API is identical, only the access flow differs."
license: Apache-2.0
metadata:
  version: "1.0.0"
  author: witanlabs
  source: https://github.com/witanlabs/witan-cli
---

> **Running in Claude Cowork?** The `witan` CLI isn't preinstalled — see the witan-xlsx skill's [references/cowork-setup.md](../witan-xlsx/references/cowork-setup.md) for install steps.

## What this skill is for

Google Sheets adds **two things** on top of normal spreadsheet work: an OAuth-based access flow, and `gs://` references instead of file paths. Everything else — the JavaScript you run against the sheet — is the **same engine and the same `xlsx.*` API** as the witan-xlsx skill.

- **The in-sheet API surface is identical.** `witan gsheets exec` exposes the same `xlsx` global and `wb` workbook, the same `setCells`/`readRangeTsv`/`findCells`/`traceToInputs`/charts/formatting functions. After reading this file, read **[../witan-xlsx/references/api.md](../witan-xlsx/references/api.md)** before your first `exec`, and follow the witan-xlsx **Quality floor** (formulas not values, no magic numbers, format to type) and **Verify before done** rules. This skill does **not** repeat them.
- **What's different is access.** A Google Sheet is reached over OAuth, not the filesystem. That's the rest of this document.

## Access model — read this before any gsheets command

Two independent gates, both required for an existing sheet:

1. **Connection (account-level, once)** — the user links their Google account to witan. Without it, *every* gsheets command fails.
2. **Authorization (per-sheet)** — under the `drive.file` OAuth scope, connecting grants access to **nothing** on its own. Each spreadsheet the user did **not** create through witan must be authorized individually by picking it in Google's file picker. A sheet you **create** via witan (`--create`, or the `new` ref) is authorized automatically.

Authorization persists at Google, so a given sheet only needs the picker **once** (until the user disconnects). Check before assuming you need it.

## The golden path for agents — never block, never make the user report back

`connect` and `authorize` need a human to click through a Google page. **In agent mode (always pass `--json`) they do NOT open a browser and do NOT wait** — they print a URL and return immediately. Your job is to (a) hand the URL to the human and (b) **poll for completion yourself** with `gsheets status … --wait`, which blocks until the gate clears. Do **not** print the URL and then ask the user "tell me when you're done" — poll instead.

### Connect (once per account)

```bash
# 1. Get the consent URL (returns immediately; does not block)
witan gsheets connect --json
# -> {"connected": false, "authorization_url": "https://accounts.google.com/o/oauth2/v2/auth?..."}
#    (if already connected: {"connected": true} — skip the rest)

# 2. Give that authorization_url to the human to open and approve.

# 3. Poll until the account is connected (blocks until done, then exits 0):
witan gsheets status --wait --json
# -> {"status": "connected", ...}
```

### Authorize a specific sheet (once per sheet you didn't create)

```bash
# 1. Get the picker URL (returns immediately; idempotent — {"authorized": true} if already done)
witan gsheets authorize gs://SPREADSHEET_ID --json
# -> {"authorized": false, "picker_url": "https://...", "expires_in_seconds": 600}

# 2. Give picker_url to the human — it EXPIRES IN 10 MINUTES, so hand it over promptly.

# 3. Poll the per-sheet gate until it clears:
witan gsheets status gs://SPREADSHEET_ID --wait --json
# -> {"authorized": true}
```

`--wait` is the agent's blocking primitive — same role that polling plays in `witan auth login`. If the human is slow, `--wait` times out after ~5 minutes; just re-run it (the grant, once made, sticks).

## Referencing a sheet

`exec`, `lint`, `render`, `rpc`, `authorize`, and `status` accept either form interchangeably:

- **Short:** `gs://SPREADSHEET_ID`
- **Full URL:** `https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit` (gid/fragment ignored)
- **Create instead of open:** the literal ref `new`, or the `--create` flag (with optional `--title`). Created sheets are auto-authorized and their URL is in the result.

## Running JavaScript — same as witan xlsx exec

Identical globals (`xlsx`, `wb`), identical functions ([../witan-xlsx/references/api.md](../witan-xlsx/references/api.md)). Prefer `--stdin` with a single-quoted heredoc (no shell escaping); `--expr` is a one-expression shorthand.

```bash
# Read
witan gsheets exec gs://SPREADSHEET_ID --stdin <<'WITAN'
const sheets = await xlsx.listSheets(wb)
const tsv = await xlsx.readRangeTsv(wb, { sheet: "Sheet1", from: {row:1,col:1}, to: {row:20,col:8} })
return { sheets, tsv }
WITAN

# Write — Google Sheets writes are LIVE (there is no ephemeral/--save distinction;
# changes persist to the user's sheet immediately). Read recalculated values back
# from result.touched["Sheet!A1"]; never recompute in JS.
witan gsheets exec gs://SPREADSHEET_ID --expr "await xlsx.setCells(wb, [{address:'Sheet1!B10', formula:'=SUM(B2:B9)'}])"

# Create and populate a brand-new sheet in one call
witan gsheets exec --create --title "Q3 Model" --stdin <<'WITAN'
await xlsx.setCells(wb, [{ address: "Sheet1!A1", value: "Revenue" }])
return await xlsx.readCell(wb, "Sheet1!A1")
WITAN
```

> **Caution: writes are not ephemeral.** Unlike `witan xlsx exec` (where omitting `--save` discards changes), a `gsheets exec` write mutates the user's real spreadsheet. For what-if exploration, write to a throwaway sheet (`--create`) rather than the user's sheet, or read-only-trace first and only write when you mean it.

Siblings: `witan gsheets lint gs://ID [-r RANGE]` (semantic formula checks; exits 2 on findings), `witan gsheets render gs://ID -r RANGE` (PNG preview), `witan gsheets rpc gs://ID` (newline-delimited RPC over stdio; `--create` to bind a new sheet). Follow witan-xlsx's **Verify before done** discipline — lint and spot-check before calling an edit finished.

## Errors and exit codes

- **Exit code 3 / `needs_file_authorization`** — the sheet isn't authorized (or doesn't exist). Run the **Authorize** flow above, then retry. This is the single most common gsheets failure; handle it by authorizing, not by giving up.
- **`status: not_connected`** — the account isn't linked. Run the **Connect** flow.
- **`session expired` / `unavailable`** — the witan session lapsed. The user must re-run `witan auth login` (a separate, browser/device flow); the gsheets connection itself is fine.
- **API-key auth is not supported for Google Sheets** — these commands require a user session (`witan auth login`), not `--api-key`/`WITAN_API_KEY`.

## Disconnecting

`witan gsheets disconnect` removes the witan-side connection and revokes the Google authorization. The user re-runs `connect` to start over. (Re-connecting requests only the `drive.file` scope.)
