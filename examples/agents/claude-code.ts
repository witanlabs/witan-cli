import { query } from '@anthropic-ai/claude-agent-sdk';
import type {
  HookCallbackMatcher,
  HookEvent,
  SDKMessage,
  SandboxSettings,
} from '@anthropic-ai/claude-agent-sdk';

export type { SDKMessage, SDKResultMessage } from '@anthropic-ai/claude-agent-sdk';

export interface ClaudeCodeAgentOptions {
  prompt: string;
  cwd: string;
  systemPrompt: string | { type: 'preset'; preset: 'claude_code'; append: string };
  model?: string;
  signal?: AbortSignal;
  permissionMode?: 'default' | 'plan' | 'acceptEdits';
  allowedTools?: string[];
  sandbox?: SandboxSettings;
  hooks?: Partial<Record<HookEvent, HookCallbackMatcher[]>>;
}

/**
 * Invoke Claude Code via the SDK.
 *
 * Yields each `SDKMessage` as it arrives. The final yielded message will
 * have `type === 'result'` (an `SDKResultMessage`). Callers can optionally
 * pass `hooks` (e.g. for tracing tool calls).
 *
 * ```ts
 * for await (const message of invokeClaudeCode({ prompt, cwd, systemPrompt })) {
 *   if (message.type === 'result') result = message;
 * }
 * ```
 */
export async function* invokeClaudeCode(
  options: ClaudeCodeAgentOptions,
): AsyncGenerator<SDKMessage> {
  const {
    prompt,
    cwd,
    systemPrompt,
    model = 'claude-opus-4-6',
    signal,
    permissionMode = 'acceptEdits',
    allowedTools = ['Bash', 'Read', 'Write', 'Edit', 'Glob', 'Grep'],
    sandbox,
    hooks,
  } = options;

  const abortController = new AbortController();
  if (signal) {
    signal.addEventListener('abort', () => abortController.abort(), { once: true });
  }

  // Allow launching from inside a Claude Code session
  delete process.env.CLAUDECODE;

  const conversation = query({
    prompt,
    options: {
      cwd,
      model,
      permissionMode,
      allowedTools,
      systemPrompt,
      abortController,
      ...(sandbox ? { sandbox } : {}),
      ...(hooks ? { hooks } : {}),
    },
  });

  for await (const message of conversation) {
    yield message;
  }
}
