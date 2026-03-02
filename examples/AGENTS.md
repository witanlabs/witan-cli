# Agent Guidelines

## Project overview

This is a TypeScript examples directory for the witan CLI. It provides two
agent wrappers (`invokeClaudeCode` and `invokeDeepAgent`), a shared runner
module (`run.ts`), and three entry points:

- `qna.ts` — ask questions about existing Excel workbooks.
- `model-builder.ts` — build financial models from a prompt specification.
- `verify.ts` — audit a workbook for formula bugs using the witan linter.

## Key files

- `qna.ts` — QnA CLI entry point. Parses args, copies the input workbook into
  a sandboxed temp directory, invokes the chosen runner, and cleans up.
- `model-builder.ts` — Model builder CLI entry point. Creates an empty workbook,
  invokes the chosen runner with a model specification prompt, and copies output
  workbooks to the output directory.
- `verify.ts` — Workbook audit CLI entry point. Generates a buggy P&L workbook
  (demo mode) or accepts a user-provided workbook, loads the xlsx-verify skill,
  and asks the agent to find formula issues. The xlsx-verify skill provides
  lint (semantic formula analysis), calc (recalculation verification), and
  render (visual inspection) — tools that catch bugs invisible to normal
  spreadsheet use. Combine with any write tooling to create an edit-verify loop.
- `lib/run.ts` — Shared runner dispatch module. Exports `runAgent()` which handles
  Claude Code and DeepAgents invocation with system prompt assembly and logging.
- `lib/setup.ts` — Loads `.env`, verifies the witan binary exists, prepends it
  to PATH.
- `lib/format.ts` — Console output helpers (ANSI formatting, message logging).
- `lib/demo-workbook.ts` — Sample workbook generator for the QnA demo mode.
- `lib/buggy-workbook.ts` — Buggy multi-region P&L workbook generator for the
  verify demo mode. Contains 5 planted formula bugs across 3 sheets.
- `prompts/loan-amortization.md` — Default model-builder prompt (loan
  amortization with circular references).
- `agents/claude-code.ts` — Async generator wrapping the Claude Code SDK
  `query()` function.
- `agents/deep-agents.ts` — Async generator wrapping DeepAgents with a local
  filesystem + shell execution backend.

## Build and run

```bash
pnpm install
pnpm qna <workbook.xlsx> <question>
pnpm model-builder [prompt-file]
pnpm verify [workbook.xlsx]
```

The witan CLI binary must exist at `../witan`. Build it with `cd .. && make build`.

## Environment

Copy `.env.example` to `.env` and fill in API keys. Required variables depend
on the runner:

- **claude-code**: `ANTHROPIC_API_KEY`
- **deep-agents**: `ANTHROPIC_API_KEY` (or `OPENAI_API_KEY` if using an OpenAI model)
- **witan CLI**: `WITAN_API_KEY`, `WITAN_API_URL`

## Running the local witan API stack

The witan CLI needs a running API server. When `.env` points to
`http://localhost:3000`, you must start the local witan-alfred stack. The
agent API (port 3000) validates API keys by calling the management API
(port 3001) — **both must be running** or all requests will fail with
`401 unauthorized`.

The witan-alfred repo lives at `../../witan-alfred` (sibling to witan-cli)
or at the path configured in the parent project. Always `source .envrc` in
witan-alfred before starting either API process.

### Starting the stack

```bash
cd <witan-alfred-root>
source .envrc

# 1. Start management API (port 3001) — needs Docker containers on :5432
pnpm --filter @witan/management-api dev:docker    # fresh start (starts Docker + seeds)
# OR if Docker containers are already running:
pnpm --filter @witan/management-api dev:docker:api  # starts Node only, skips seeding
pnpm --filter @witan/management-api run db:seed     # seed manually after dev:docker:api

# 2. Start agent API (port 3000) — needs Docker containers on :5433/:9000/:6379
#    Wait for management API to be healthy first!
pnpm --filter @witan/api dev:all      # fresh start
# OR if Docker containers are already running:
pnpm --filter @witan/api dev:all:api  # starts Node only
```

### Verifying the stack

```bash
curl http://localhost:3001/health  # management API
curl http://localhost:3000/health  # agent API
```

### Troubleshooting

- **`401 unauthorized` on all requests**: The agent API cannot reach the
  management API. Check that the management API is running on port 3001 and
  that `db:seed` has been run. The `dev:docker:api` variant skips seeding.
- **`dial tcp [::1]:3000: connect: operation not permitted`**: The API
  isn't running, or a sandbox is blocking localhost connections. Start the
  API stack. If running inside a Claude Code sandbox, note that
  `allowedDomains: ['localhost']` may not cover IPv6 loopback (`::1`) —
  disable sandboxing or use the production API instead.
- **Local dev API key**: `wk_live_localdev_org_00000000000000000000000000`

## Type checking

```bash
npx tsc --noEmit
```

There is no test suite or linter configured in this directory.
