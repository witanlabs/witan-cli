# Witan Agent Examples

Run questions against Excel workbooks using AI agents.

## Setup

```bash
cp .env.example .env   # fill in your API keys
pnpm install
```

The witan CLI binary must be built in the parent directory (`cd .. && make build`).

## Usage

```bash
# Claude Code (default)
pnpm qna path/to/workbook.xlsx What is the total revenue?

# DeepAgents
pnpm qna -r deep-agents path/to/workbook.xlsx What is the total revenue?
```

### Options

| Flag | Description |
|------|-------------|
| `-r, --runner` | `claude-code` (default) or `deep-agents` |
| `-m, --model` | Model ID (default: `claude-opus-4-6` / `gpt-5.2`) |
| `-v, --verbose` | Print raw JSON messages |

## Project structure

```
qna.ts              Entry point
setup.ts             Env loading + PATH setup
format.ts            Console output formatting
agents/
  claude-code.ts     Claude Code SDK wrapper
  deep-agents.ts     DeepAgents/LangChain wrapper
  index.ts           Re-exports
```
