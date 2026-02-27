# Agent Guidelines

## Project overview

This is a TypeScript examples directory for the witan CLI. It provides two
agent wrappers (`invokeClaudeCode` and `invokeDeepAgent`) and a `qna.ts`
entry point that runs questions against Excel workbooks.

## Key files

- `qna.ts` — CLI entry point. Parses args, sets up a sandboxed temp working
  directory, invokes the chosen runner, and cleans up.
- `setup.ts` — Loads `.env`, verifies the witan binary exists, prepends it
  to PATH.
- `format.ts` — Console output helpers (ANSI formatting, message logging).
- `agents/claude-code.ts` — Async generator wrapping the Claude Code SDK
  `query()` function.
- `agents/deep-agents.ts` — Async generator wrapping DeepAgents with a local
  filesystem + shell execution backend.

## Build and run

```bash
pnpm install
pnpm qna <workbook.xlsx> <question>
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
