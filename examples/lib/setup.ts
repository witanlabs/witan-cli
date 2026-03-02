import { execSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import dotenv from 'dotenv';

/**
 * Load environment variables from `examples/.env` and ensure the witan CLI
 * binary is reachable on PATH.
 *
 * Call this once at the top of your entry-point (e.g. `qna.ts`) before
 * invoking any runner.
 */
export function loadEnv(): void {
  // 1. Load .env from the examples directory (one level up from lib/)
  const envPath = path.resolve(import.meta.dirname, '..', '.env');
  dotenv.config({ path: envPath, override: true });

  // 2. Verify the witan binary exists at the repo root (two levels up from lib/)
  const witanBinary = path.resolve(import.meta.dirname, '..', '..', 'witan');
  if (!fs.existsSync(witanBinary)) {
    console.error(
      `witan binary not found at ${witanBinary}\n` +
        'Run "make build" in the witan-cli root to compile it.',
    );
    process.exit(1);
  }

  // 3. Prepend the repo root to PATH so child processes can find witan
  const repoRoot = path.resolve(import.meta.dirname, '..', '..');
  process.env.PATH = `${repoRoot}${path.delimiter}${process.env.PATH ?? ''}`;
}


/**
 * Create a Python virtual environment inside `workDir` with openpyxl installed.
 *
 * The venv lives at `<workDir>/venv/` so the agent can use `./venv/bin/python`
 * without needing PATH manipulation (which gets reset by shell profile sourcing).
 */
export function setupPythonVenv(workDir: string): void {
  const venvDir = path.join(workDir, 'venv');
  if (fs.existsSync(venvDir)) return;

  console.log('Setting up Python venv with openpyxl...');
  execSync('python3 -m venv venv', { cwd: workDir, stdio: 'pipe' });
  execSync('./venv/bin/pip install -q openpyxl', { cwd: workDir, stdio: 'pipe' });
  console.log('Python venv ready.\n');
}
