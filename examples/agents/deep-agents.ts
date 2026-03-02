import type { ExecFileException } from 'node:child_process';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

import { createDeepAgent, FilesystemBackend } from 'deepagents';
import { initChatModel } from 'langchain';

const execFileAsync = promisify(execFile);

/**
 * A FilesystemBackend subclass that adds local shell execution support.
 *
 * DeepAgents' `FilesystemBackend` only provides file operations (read, write, edit,
 * grep, glob). The `execute` tool requires a backend that implements
 * `SandboxBackendProtocol` (i.e. has an `execute` method and `id` property).
 *
 * This class adds local subprocess execution via `sh -c`, with support for
 * custom environment variables (needed for the Python venv PATH).
 */
export class LocalExecutionBackend extends FilesystemBackend {
  readonly id: string;
  private rootDir: string;
  private execEnv: NodeJS.ProcessEnv;

  constructor(options: { rootDir: string; env?: NodeJS.ProcessEnv }) {
    super({ rootDir: options.rootDir });
    this.rootDir = options.rootDir;
    this.id = `local-${options.rootDir}`;
    this.execEnv = options.env ?? process.env;
  }

  async execute(
    command: string,
  ): Promise<{ output: string; exitCode: number | null; truncated: boolean }> {
    const MAX_OUTPUT = 100_000;
    try {
      const { stdout, stderr } = await execFileAsync('sh', ['-c', command], {
        cwd: this.rootDir,
        env: this.execEnv,
        timeout: 120_000,
        maxBuffer: 10 * 1024 * 1024,
      });
      const output = stdout + (stderr ? `\n${stderr}` : '');
      const truncated = output.length > MAX_OUTPUT;
      return {
        output: truncated ? output.slice(0, MAX_OUTPUT) : output,
        exitCode: 0,
        truncated,
      };
    } catch (err: unknown) {
      const e = err as ExecFileException;
      const output =
        (e.stdout ?? '') + (e.stderr ? `\n${e.stderr}` : '') ||
        (err instanceof Error ? err.message : String(err));
      const exitCode = typeof e.code === 'number' ? e.code : e.killed ? 137 : 1;
      const truncated = output.length > MAX_OUTPUT;
      return {
        output: truncated ? output.slice(0, MAX_OUTPUT) : output,
        exitCode,
        truncated,
      };
    }
  }
}

export interface DeepAgentOptions {
  prompt: string;
  cwd: string;
  systemPrompt: string;
  model?: string;
  signal?: AbortSignal;
  env?: NodeJS.ProcessEnv;
}

/**
 * Invoke a DeepAgent with local file system + shell execution backend.
 *
 * Yields raw state-update chunks as each graph node (model call, tool
 * execution) completes. The last chunk contains the final state with the
 * full message history.
 *
 * ```ts
 * let last: unknown;
 * for await (const chunk of invokeDeepAgent({ prompt, cwd, systemPrompt })) {
 *   last = chunk;
 * }
 * ```
 */
export async function* invokeDeepAgent(
  options: DeepAgentOptions,
): AsyncGenerator<unknown> {
  const {
    prompt,
    cwd,
    systemPrompt,
    model: modelName = 'gpt-5.2',
    signal,
    env,
  } = options;

  const model = await initChatModel(modelName);

  const backend = new LocalExecutionBackend({
    rootDir: cwd,
    env,
  });

  const agent = createDeepAgent({
    model,
    systemPrompt,
    backend,
  });

  // DeepAgent's recursive generics exceed TS instantiation depth limits,
  // so we narrow to the stream interface we actually use.
  interface Streamable {
    stream(
      input: { messages: { role: string; content: string }[] },
      config?: { signal?: AbortSignal },
    ): Promise<AsyncIterable<unknown>>;
  }
  const stream = await (agent as unknown as Streamable).stream(
    { messages: [{ role: 'user', content: prompt }] },
    { signal },
  );

  yield* stream;
}
