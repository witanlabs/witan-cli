import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { parseArgs } from 'node:util';
import ExcelJS from 'exceljs';
import { loadEnv } from './lib/setup.js';
import { logWorkDir } from './lib/format.js';
import { runAgent } from './lib/run.js';

loadEnv();

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    runner: { type: 'string', short: 'r', default: 'claude-code' },
    model: { type: 'string', short: 'm' },
    verbose: { type: 'boolean', short: 'v', default: false },
    output: { type: 'string', short: 'o', default: './output' },
    help: { type: 'boolean', short: 'h', default: false },
  },
});

if (values.help) {
  console.error(`Usage: pnpm model-builder [options] [prompt-file]

Arguments:
  prompt-file            Path to a .md prompt file (default: prompts/loan-amortization.md)

Options:
  -r, --runner <name>   claude-code | deep-agents (default: claude-code)
  -m, --model <model>   Model ID (default: claude-opus-4-6 for claude-code, gpt-5.2 for deep-agents)
  -v, --verbose         Full JSON output
  -o, --output <path>   Where to copy result workbooks (default: ./output/)
  -h, --help            Show this help message`);
  process.exit(0);
}

const promptFile = positionals[0]
  ? path.resolve(positionals[0])
  : path.resolve(import.meta.dirname, 'prompts/loan-amortization.md');
const promptContent = fs.readFileSync(promptFile, 'utf-8');

// Extract the workbook filename from the prompt: look for "in <filename>.xlsx"
const filenameMatch = promptContent.match(/\bin\s+([\w-]+\.xlsx)\b/i);
const workbookFilename = filenameMatch?.[1] ?? 'model.xlsx';

const runner = values.runner!;
const verbose = values.verbose!;
const outputDir = path.resolve(values.output!);

const skillPath = path.resolve(import.meta.dirname, '../skills/xlsx-code-mode/SKILL.md');
const skill = fs.readFileSync(skillPath, 'utf-8');

const defaultModel = runner === 'deep-agents' ? 'gpt-5.2' : 'claude-opus-4-6';

// Create temp working directory with an empty workbook
const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'model-builder-'));
logWorkDir(workDir);

async function createEmptyWorkbook(filePath: string): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  workbook.addWorksheet('Sheet1');
  await workbook.xlsx.writeFile(filePath);
}

async function copyOutputFiles(): Promise<void> {
  const files = fs.readdirSync(workDir).filter((f) => f.endsWith('.xlsx'));
  if (files.length === 0) {
    console.log('No .xlsx files found in working directory.');
    return;
  }
  fs.mkdirSync(outputDir, { recursive: true });
  for (const file of files) {
    const src = path.join(workDir, file);
    const dest = path.join(outputDir, file);
    fs.copyFileSync(src, dest);
    console.log(`Copied: ${dest}`);
  }
}

await createEmptyWorkbook(path.join(workDir, workbookFilename));

const prompt = `${promptContent}\n\nInput files: ${workbookFilename}\nThe input files are in your current working directory.`;

try {
  await runAgent({
    runner,
    model: values.model ?? defaultModel,
    prompt,
    workDir,
    skill,
    verbose,
  });
  await copyOutputFiles();
} finally {
  fs.rmSync(workDir, { recursive: true, force: true });
}
