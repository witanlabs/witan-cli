# Agent Guidelines

## Project overview

This is a TypeScript examples directory for the witan CLI. It provides two
agent wrappers (`invokeClaudeCode` and `invokeDeepAgent`), a shared runner
module (`run.ts`), and two entry points:

- `qna.ts` — ask questions about existing Excel workbooks.
- `model-builder.ts` — build financial models from a prompt specification.

## Key files

- `qna.ts` — QnA CLI entry point. Parses args, copies the input workbook into
  a sandboxed temp directory, invokes the chosen runner, and cleans up.
- `model-builder.ts` — Model builder CLI entry point. Creates an empty workbook,
  invokes the chosen runner with a model specification prompt, and copies output
  workbooks to the output directory.
- `lib/run.ts` — Shared runner dispatch module. Exports `runAgent()` which handles
  Claude Code and DeepAgents invocation with sandbox config and logging.
- `lib/setup.ts` — Loads `.env`, verifies the witan binary exists, prepends it
  to PATH.
- `lib/format.ts` — Console output helpers (ANSI formatting, message logging).
- `lib/demo-workbook.ts` — Sample workbook generator for the QnA demo mode.
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
```

The witan CLI binary must exist at `../witan`. Build it with `cd .. && make build`.

## Environment

Copy `.env.example` to `.env` and fill in API keys. Required variables depend
on the runner:

- **claude-code**: `ANTHROPIC_API_KEY`
- **deep-agents**: `ANTHROPIC_API_KEY` (or `OPENAI_API_KEY` if using an OpenAI model)
- **witan CLI**: `WITAN_API_KEY`, `WITAN_API_URL`

## Type checking

```bash
npx tsc --noEmit
```

There is no test suite or linter configured in this directory.
