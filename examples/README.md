# Witan Agent Examples

Give an AI agent a spreadsheet and a task. These examples show how to wire up
an agent that can read, query, and build Excel workbooks using the witan CLI.

Two runners are included — **Claude Code** (Anthropic's agent SDK) and
**DeepAgents** (LangChain-based) — so you can compare how different agent
frameworks approach the same spreadsheet problems.

## Getting started

### 1. Build the witan CLI

From the repository root:

```bash
make build
```

### 2. Set up the examples

```bash
cd examples
cp .env.example .env
pnpm install
```

Fill in your `.env`:

- `ANTHROPIC_API_KEY` — required for the default Claude Code runner
- `WITAN_API_KEY` and `WITAN_API_URL` — your witan API credentials
- `OPENAI_API_KEY` — only needed if using DeepAgents with an OpenAI model

### 3. Run the QnA demo

```bash
pnpm qna
```

This generates a small quarterly-revenue workbook, asks the agent a question
about it, and prints the answer. No spreadsheet needed — it's a self-contained
smoke test to verify everything is wired up.

## Experimenting

### Ask questions about your own workbooks

Once the demo works, point it at a real spreadsheet:

```bash
pnpm qna path/to/workbook.xlsx "What is the total revenue in Year 3?"
```

The agent reads the workbook using witan CLI tools (find, exec, render),
reasons about the data, and returns an answer.

### Build a financial model from scratch

The model-builder creates an empty workbook, gives the agent a specification,
and lets it build the model:

```bash
# Default: loan amortization with circular references
pnpm model-builder

# Your own spec
pnpm model-builder path/to/spec.md
```

Output workbooks are saved to `./output/`. Open them in Excel to inspect the
result — formulas, formatting, and structure are all agent-generated.

### Compare runners

Both examples support the `-r` flag to switch between agent frameworks:

```bash
# Claude Code (default) — uses Anthropic's agent SDK with sandbox
pnpm qna -r claude-code path/to/workbook.xlsx "What is the EBITDA margin?"

# DeepAgents — uses LangChain with local shell execution
pnpm qna -r deep-agents path/to/workbook.xlsx "What is the EBITDA margin?"
```

You can also override the model with `-m`:

```bash
pnpm qna -m claude-sonnet-4-5-20250929 workbook.xlsx "Summarize this model"
pnpm qna -r deep-agents -m gpt-4.1 workbook.xlsx "Summarize this model"
```

### Verbose mode

Pass `-v` to see the raw message stream — useful for debugging agent behavior
or understanding the tool-call sequence:

```bash
pnpm qna -v path/to/workbook.xlsx "What is cell B12?"
```

### Write your own prompts

The model-builder reads a markdown file as the agent's task specification.
Look at `prompts/loan-amortization.md` for the format, then create your own:

```bash
pnpm model-builder prompts/my-dcf-model.md
```

## Options reference

| Flag | Description | Applies to |
|------|-------------|------------|
| `-r, --runner` | `claude-code` (default) or `deep-agents` | both |
| `-m, --model` | Model ID | both |
| `-v, --verbose` | Print raw JSON messages | both |
| `-o, --output` | Output directory (default: `./output/`) | model-builder |

## Project structure

```
qna.ts                 QnA entry point
model-builder.ts       Model builder entry point
lib/
  run.ts               Shared runner dispatch
  setup.ts             Env + PATH setup
  format.ts            Console output formatting
  demo-workbook.ts     Sample workbook generator
agents/
  claude-code.ts       Claude Code SDK wrapper
  deep-agents.ts       DeepAgents wrapper
  index.ts             Re-exports
prompts/
  loan-amortization.md Default model spec
output/                Generated workbooks (gitignored)
```
