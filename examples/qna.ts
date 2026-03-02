import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { parseArgs } from 'node:util';
import { loadEnv } from './lib/setup.js';
import { logWorkDir } from './lib/format.js';
import { runAgent } from './lib/run.js';
import { createDemoWorkbook, DEMO_FILENAME, DEMO_QUESTION } from './lib/demo-workbook.js';

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
  console.error(`Usage: pnpm qna [options] [workbook.xlsx] [question...]

When run with no arguments, generates a sample workbook and runs a demo question.

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
const isDemo = positionals.length < 2;

const skillPath = path.resolve(import.meta.dirname, '../skills/xlsx-code-mode/SKILL.md');
const skill = fs.readFileSync(skillPath, 'utf-8');

let filename: string;
let question: string;
const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'qna-'));

if (isDemo) {
  filename = DEMO_FILENAME;
  question = DEMO_QUESTION;
  await createDemoWorkbook(path.join(workDir, filename));
  console.log(`Demo mode: generated ${filename}\n`);
} else {
  const [workbookPath, ...questionParts] = positionals;
  question = questionParts.join(' ');
  const resolvedWorkbook = path.resolve(workbookPath);
  filename = path.basename(resolvedWorkbook);
  fs.copyFileSync(resolvedWorkbook, path.join(workDir, filename));
}

logWorkDir(workDir);

const prompt = `${question}\n\nInput files: ${filename}\nThe input files are in your current working directory.`;

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
  console.log('  pnpm qna path/to/your-workbook.xlsx "Your question here"\n');
}
