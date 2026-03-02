# verify.ts example — bugs from first run

## 1. ~~WORKING_DIR_POLICY violation~~ FIXED

Agent's first action was `cd /Users/sam && ls -la regional-pnl.xlsx` despite
CWD being the temp directory containing the workbook.

**Root cause:** `run.ts` passed `systemPrompt` as a plain string, which
**replaces** Claude Code's entire default system prompt. The agent lost all
standard Claude Code guidance about working directory behavior.

**Fix applied:** Changed `systemPrompt` from a plain string to the preset
form: `{ type: 'preset', preset: 'claude_code', append: ... }`. This appends
the WORKING_DIR_POLICY + skill to Claude Code's default prompt instead of
replacing it. Also reordered so WORKING_DIR_POLICY comes before the skill.

## 2. ~~Sandbox blocks localhost (witan CLI can't reach API)~~ FIXED

`witan xlsx calc` failed with:
```
dial tcp [::1]:3000: connect: operation not permitted
```

The Claude Code sandbox blocked outbound connections to localhost even with
`allowedDomains: ['localhost']` configured. The Go CLI resolves `localhost` to
IPv6 `::1` which the sandbox proxy didn't allow. Setting
`allowedDomains` to domain names doesn't cover IP-level resolution.

**Fix applied:** Disabled sandboxing (`enabled: false`) in `run.ts`. The
sandbox's `allowedDomains` doesn't reliably handle localhost/loopback. Can
revisit if the SDK adds IP-based allow lists or fixes localhost resolution.

## 3. ~~`witan xlsx lint` returns 500 on the buggy workbook~~ FIXED

Root cause: the VLOOKUP in `buggy-workbook.ts` used a full-column reference
(`'FX Rates'!A:B`) which triggers a null reference in xlsx-serve's lint
endpoint. Bounded ranges work fine.

**Fix applied:** Changed `A:B` to `A1:B10` in `buggy-workbook.ts`. Lint now
returns 66 issues (8 errors, 58 warnings) including D008, D001, D009, D023.

Note: the underlying xlsx-serve bug with full-column refs in VLOOKUP still
exists — should be reported to witan-alfred separately.

## 4. Python `!=` escaping in Bash `-c` strings

Agent tried inline Python via `Bash -c` and `!=` was escaped to `\!=`,
causing `SyntaxError` twice. Agent eventually worked around it by writing a
`.py` file.

**Fix:** Minor — agent behavior issue, not easily fixable. Could add a hint in
the system prompt to write Python scripts to files instead of using `-c`.

## 5. Sibling tool call cascade failures

When the agent fires parallel tool calls and one fails, all siblings get
`Sibling tool call errored`, wasting turns. Happened 3+ times in this run.

**Fix:** Not directly fixable (Claude Code behavior). Could mitigate by
prompting the agent to avoid unnecessary parallel tool calls during
verification.

## 6. Agent never completed the audit

The trace ends without a final audit report listing the 5 planted bugs. The
cascade of sandbox, lint, and escaping errors consumed the agent's turns
without producing useful output.

**Fix:** Resolving items 1-3 above should unblock the agent enough to complete
the task.
