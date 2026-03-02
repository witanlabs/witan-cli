import type { SDKResultMessage } from '../agents/index.js';
import { invokeClaudeCode, invokeDeepAgent } from '../agents/index.js';
import { logClaudeCodeMessage, logDeepAgentChunk, answerSeparator } from './format.js';

export const WORKING_DIR_POLICY = `
## working-directory-policy
Your current working directory is already set to the correct location. All input files are here.
- Do NOT use \`cd\` — you are already in the right place.
- Do NOT access, read, write, or reference any files outside the current directory.
- Use relative paths only. Never use absolute paths.
- Save all output files in the current directory.
`;

export interface RunAgentOptions {
  runner: string;
  model: string;
  prompt: string;
  workDir: string;
  skill: string;
  verbose: boolean;
}

export async function runAgent(options: RunAgentOptions): Promise<void> {
  const { runner, model, prompt, workDir, skill, verbose } = options;

  console.log(`Runner: ${runner} | Model: ${model}\n`);

  const appendedPrompt = WORKING_DIR_POLICY + '\n\n' + skill;

  if (runner === 'claude-code') {
    let result: SDKResultMessage | undefined;
    for await (const message of invokeClaudeCode({
      prompt,
      cwd: workDir,
      systemPrompt: {
        type: 'preset' as const,
        preset: 'claude_code' as const,
        append: appendedPrompt,
      },
      model,
    })) {
      logClaudeCodeMessage(message, verbose);
      if (message.type === 'result') result = message;
    }
    console.log(answerSeparator);
    if (!result || result.subtype !== 'success') {
      console.log(result?.errors?.join('\n') ?? 'No result');
    }
  } else if (runner === 'deep-agents') {
    let answer = '';
    for await (const chunk of invokeDeepAgent({
      prompt,
      cwd: workDir,
      systemPrompt: appendedPrompt,
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
}
