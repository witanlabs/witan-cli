import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { parseArgs } from 'node:util';
import { loadEnv } from './setup.js';
import { invokeClaudeCode, invokeDeepAgent } from './agents/index.js';
import type { SDKResultMessage } from './agents/index.js';
import { logClaudeCodeMessage, logDeepAgentChunk, logWorkDir, answerSeparator } from './format.js';

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

if (values.help || positionals.length < 2) {
  console.error(`Usage: pnpm qna [options] <workbook.xlsx> <question...>

Options:
  -r, --runner <name>   Runner to use: claude-code | deep-agents (default: claude-code)
  -m, --model <model>   Model ID (default: claude-opus-4-6 for claude-code, gpt-5.2 for deep-agents)
  -v, --verbose         Print full JSON for each message/chunk
  -h, --help            Show this help message`);
  process.exit(values.help ? 0 : 1);
}

const [workbookPath, ...questionParts] = positionals;
const question = questionParts.join(' ');
const runner = values.runner!;
const verbose = values.verbose!;

const skillPath = path.resolve(import.meta.dirname, '../skill/xlsx-code-mode/SKILL.md');
const skill = fs.readFileSync(skillPath, 'utf-8');
const resolvedWorkbook = path.resolve(workbookPath);
const filename = path.basename(resolvedWorkbook);
const prompt = `${question}\n\nInput files: ${filename}\nThe input files are in your current working directory.`;

/** Create a temp working directory and copy the input workbook into it. */
function prepareWorkDir(): string {
  const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'qna-'));
  fs.copyFileSync(resolvedWorkbook, path.join(workDir, filename));
  return workDir;
}

const WORKING_DIR_POLICY = `
## working-directory-policy
Your current working directory is already set to the correct location. All input files are here.
- Do NOT use \`cd\` — you are already in the right place.
- Do NOT access, read, write, or reference any files outside the current directory.
- Use relative paths only. Never use absolute paths.
- Save all output files in the current directory.
`;

const workDir = prepareWorkDir();
logWorkDir(workDir);

try {
  if (runner === 'claude-code') {
    const model = values.model ?? 'claude-opus-4-6';
    console.log(`Runner: ${runner} | Model: ${model}\n`);
    let result: SDKResultMessage | undefined;
    for await (const message of invokeClaudeCode({
      prompt,
      cwd: workDir,
      systemPrompt: skill,
      model,
      sandbox: {
        enabled: true,
        autoAllowBashIfSandboxed: true,
        network: { allowedDomains: ['localhost', 'api.witanlabs.com'] },
      },
    })) {
      logClaudeCodeMessage(message, verbose);
      if (message.type === 'result') result = message;
    }
    console.log(answerSeparator);
    if (!result || result.subtype !== 'success') {
      console.log(result?.errors?.join('\n') ?? 'No result');
    }
  } else if (runner === 'deep-agents') {
    const model = values.model ?? 'gpt-5.2';
    console.log(`Runner: ${runner} | Model: ${model}\n`);
    let answer = '';
    for await (const chunk of invokeDeepAgent({
      prompt,
      cwd: workDir,
      systemPrompt: skill + WORKING_DIR_POLICY,
      model,
    })) {
      answer = logDeepAgentChunk(chunk, verbose, answer);
    }
    console.log(answerSeparator);
    if (!answer) console.log('No result');
  } else {
    console.error(`Unknown runner: ${runner}. Use "claude-code" or "deep-agents".`);
    process.exit(1);
  }
} finally {
  fs.rmSync(workDir, { recursive: true, force: true });
}
