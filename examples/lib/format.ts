import type { SDKMessage } from '../agents/index.js';

// ANSI escape codes — disabled when stdout is not a TTY (piped output, redirects)
const isTTY = process.stdout.isTTY ?? false;
const dim = isTTY ? '\x1b[2m' : '';
const bold = isTTY ? '\x1b[1m' : '';
const italic = isTTY ? '\x1b[3m' : '';
const cyan = isTTY ? '\x1b[36m' : '';
const yellow = isTTY ? '\x1b[33m' : '';
const reset = isTTY ? '\x1b[0m' : '';

const TRUNCATE_LIMIT = 200;

function truncate(s: string, max = TRUNCATE_LIMIT): string {
  const oneLine = s.replaceAll('\n', ' ').trim();
  return oneLine.length > max ? oneLine.slice(0, max) + '...' : oneLine;
}

function extractText(content: unknown): string | null {
  if (typeof content === 'string') return content;
  if (Array.isArray(content)) {
    const parts = (content as Array<Record<string, unknown>>)
      .filter((b) => b.type === 'text' && typeof b.text === 'string')
      .map((b) => b.text as string);
    return parts.length > 0 ? parts.join('\n') : null;
  }
  return null;
}

function formatToolCall(name: string, input: Record<string, unknown> | undefined): string {
  const desc = input?.description as string | undefined;
  const cmd = input?.command as string | undefined;
  let out = `\n${cyan}${bold}▶ ${name}${reset}${desc ? `${dim} — ${desc}${reset}` : ''}`;
  if (cmd) out += `\n${dim}  $ ${truncate(cmd)}${reset}`;
  return out;
}

function logContentBlocks(
  blocks: Array<Record<string, unknown>>,
  pendingTools: Map<string, string>,
  options?: { suppressText?: boolean },
): void {
  for (const block of blocks) {
    if (block.type === 'thinking' && typeof block.thinking === 'string') {
      console.log(`\n${dim}${italic}${block.thinking}${reset}`);
    } else if (block.type === 'tool_use') {
      const label = formatToolCall(block.name as string, block.input as Record<string, unknown> | undefined);
      if (block.id) {
        pendingTools.set(block.id as string, label);
      } else {
        console.log(label);
      }
    } else if (block.type === 'text' && typeof block.text === 'string' && !options?.suppressText) {
      console.log(`\n${block.text}`);
    }
  }
}

const ccPendingTools = new Map<string, string>();

/** Print a single Claude Code SDK message. */
export function logClaudeCodeMessage(message: SDKMessage, verbose: boolean): void {
  if (verbose) {
    console.log(`[${message.type}]`, JSON.stringify(message));
    return;
  }

  if (message.type === 'assistant') {
    type Block = Record<string, unknown>;
    const content = (message as unknown as { message?: { content?: Block[] } }).message?.content;
    if (Array.isArray(content)) logContentBlocks(content, ccPendingTools);
  } else if (message.type === 'user') {
    type Block = Record<string, unknown>;
    const content = (message as unknown as { message?: { content?: string | Block[] } }).message
      ?.content;
    if (Array.isArray(content)) {
      for (const block of content as Block[]) {
        const id = block.tool_use_id as string | undefined;
        const label = id ? ccPendingTools.get(id) : undefined;
        if (label) {
          ccPendingTools.delete(id!);
          console.log(label);
        }
        const text = extractText(block.content);
        if (text) console.log(`${dim}  ← ${truncate(text)}${reset}`);
      }
    }
  }
}

/**
 * Extract messages from a LangGraph update chunk.
 *
 * Chunks are keyed by graph node name (e.g. `model_request`, `tools`,
 * `todoListMiddleware.after_model`). Messages live inside the node's
 * partial state update under a `messages` key.
 */
function extractChunkMessages(chunk: unknown): unknown[] {
  if (chunk == null || typeof chunk !== 'object') return [];
  const entries = Object.values(chunk as Record<string, unknown>);
  for (const val of entries) {
    if (val != null && typeof val === 'object' && 'messages' in (val as object)) {
      const msgs = (val as { messages?: unknown[] }).messages;
      if (Array.isArray(msgs)) return msgs;
    }
  }
  return [];
}

const daPendingTools = new Map<string, string>();

/**
 * Print new messages from a DeepAgent state-update chunk.
 * Returns the last AI answer text seen (for final answer extraction).
 */
export function logDeepAgentChunk(
  chunk: unknown,
  verbose: boolean,
  lastAnswer: string,
): string {
  if (verbose) {
    console.log('[chunk]', JSON.stringify(chunk));
    return lastAnswer;
  }

  const messages = extractChunkMessages(chunk);
  let answer = lastAnswer;

  for (const msg of messages) {
    const m = msg as Record<string, unknown>;

    // Unwrap LangChain constructor wrappers: {lc: 1, type: "constructor", kwargs: {...}}
    const kwargs = m.lc != null ? (m.kwargs as Record<string, unknown>) ?? m : m;

    // Skip human messages
    const getType = (kwargs as { _getType?: () => string })._getType;
    if (typeof getType === 'function' && getType.call(kwargs) === 'human') continue;
    const msgType = (kwargs.type as string) ?? '';
    if (msgType === 'human') continue;

    // Tool calls (AIMessage) — buffer for pairing with results
    const toolCalls = kwargs.tool_calls as Array<Record<string, unknown>> | undefined;
    if (Array.isArray(toolCalls) && toolCalls.length > 0) {
      for (const tc of toolCalls) {
        const label = formatToolCall(tc.name as string, tc.args as Record<string, unknown> | undefined);
        const tcId = tc.id as string | undefined;
        if (tcId) {
          daPendingTools.set(tcId, label);
        } else {
          console.log(label);
        }
      }
    }

    // Content
    const content = kwargs.content;
    if (typeof content === 'string' && content) {
      const tcId = kwargs.tool_call_id as string | undefined;
      if (tcId) {
        const label = daPendingTools.get(tcId);
        if (label) {
          daPendingTools.delete(tcId);
          console.log(label);
        }
        console.log(`${dim}  ← ${truncate(content)}${reset}`);
      } else if (!toolCalls?.length) {
        console.log(`\n${content}`);
        answer = content;
      }
    } else if (Array.isArray(content)) {
      logContentBlocks(content as Array<Record<string, unknown>>, daPendingTools, {
        suppressText: !!toolCalls?.length,
      });
    }
  }

  return answer;
}

export const answerSeparator = `\n${yellow}${'─'.repeat(40)}${reset}\n`;

/** Print the temp working directory path (dimmed) so the user can find output files. */
export function logWorkDir(workDir: string): void {
  console.log(`${dim}workdir: ${workDir}${reset}`);
}
