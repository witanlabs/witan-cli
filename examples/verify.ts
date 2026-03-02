import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { parseArgs } from 'node:util';
import { loadEnv, setupPythonVenv } from './lib/setup.js';
import { logWorkDir } from './lib/format.js';
import { runAgent } from './lib/run.js';
import { createBuggyWorkbook, BUGGY_FILENAME, DEMO_PROMPT } from './lib/buggy-workbook.js';

loadEnv();

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    runner: { type: 'string', short: 'r', default: 'claude-code' },
    model: { type: 'string', short: 'm' },
    verbose: { type: 'boolean', short: 'v', default: false },
    help: { type: 'boolean', short: 'h', default: false },
  },
});

if (values.help) {
  console.error(`Usage: pnpm verify [options] [workbook.xlsx]

When run with no arguments, generates a sample workbook with planted formula
bugs and asks the agent to audit it.

Options:
  -r, --runner <name>   Runner to use: claude-code | deep-agents (default: claude-code)
  -m, --model <model>   Model ID (default: claude-opus-4-6 for claude-code, gpt-5.2 for deep-agents)
  -v, --verbose         Print full JSON for each message/chunk
  -h, --help            Show this help message`);
  process.exit(0);
}

const runner = values.runner!;
const verbose = values.verbose!;
const defaultModel = runner === 'deep-agents' ? 'gpt-5.2' : 'claude-opus-4-6';
const isDemo = positionals.length === 0;

// Load the xlsx-verify skill
const verifyPath = path.resolve(import.meta.dirname, '../skill/xlsx-verify/SKILL.md');
const skill = fs.readFileSync(verifyPath, 'utf-8');

let filename: string;
const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'verify-'));

if (isDemo) {
  filename = BUGGY_FILENAME;
  await createBuggyWorkbook(path.join(workDir, filename));
  console.log(`Demo mode: generated ${filename} (contains 5 planted bugs)\n`);
} else {
  const [workbookPath] = positionals;
  const resolvedWorkbook = path.resolve(workbookPath);
  filename = path.basename(resolvedWorkbook);
  fs.copyFileSync(resolvedWorkbook, path.join(workDir, filename));
}

setupPythonVenv(workDir);
logWorkDir(workDir);

const prompt = `${DEMO_PROMPT}

Input files: ${filename}
The input files are in your current working directory.

A Python virtual environment with openpyxl is available at ./venv/.
Use \`./venv/bin/python\` to run Python scripts for reading or editing the workbook.`;

try {
  await runAgent({
    runner,
    model: values.model ?? defaultModel,
    prompt,
    workDir,
    skill,
    verbose,
  });
} finally {
  fs.rmSync(workDir, { recursive: true, force: true });
}

if (isDemo) {
  console.log('Try it with your own spreadsheet:\n');
  console.log('  pnpm verify path/to/your-workbook.xlsx\n');
}
