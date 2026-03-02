import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { parseArgs } from 'node:util';
import { loadEnv } from './lib/setup.js';
import { logWorkDir } from './lib/format.js';
import { runAgent } from './lib/run.js';
import {
  generateAll,
  ACME_PDF_FILENAME,
  ACME_DOCX_FILENAME,
  ACME_PPTX_FILENAME,
  ACME_QUESTION,
} from './lib/acme-fixtures.js';

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
  console.error(`Usage: pnpm read [options] [file1 file2 ...] [question]

When run with no arguments, generates 3 Acme Corp demo documents (PDF, DOCX,
PPTX) and asks a cross-document question.

With arguments, all args except the last are file paths and the last is the
question. At least one file and a question are required.

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

const skillPath = path.resolve(import.meta.dirname, '../skills/read-source/SKILL.md');
const skill = fs.readFileSync(skillPath, 'utf-8');

let filenames: string[];
let question: string;
const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'read-'));

if (isDemo) {
  filenames = [ACME_PDF_FILENAME, ACME_DOCX_FILENAME, ACME_PPTX_FILENAME];
  question = ACME_QUESTION;
  await generateAll(workDir);
  console.log(`Demo mode: generated 3 Acme Corp documents\n`);
} else {
  question = positionals[positionals.length - 1];
  const filePaths = positionals.slice(0, -1);
  filenames = filePaths.map((fp) => {
    const resolved = path.resolve(fp);
    const basename = path.basename(resolved);
    fs.copyFileSync(resolved, path.join(workDir, basename));
    return basename;
  });
}

logWorkDir(workDir);

const prompt = `${question}\n\nInput files: ${filenames.join(', ')}\nThe input files are in your current working directory.`;

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
  console.log('Try it with your own documents:\n');
  console.log('  pnpm read report.pdf minutes.docx deck.pptx "Your question here"\n');
}
