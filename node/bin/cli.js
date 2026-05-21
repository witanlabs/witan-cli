#!/usr/bin/env node
// bin/cli.js - Standalone CLI (no dist/ dependency)
// This file is checked into source control, not compiled from TypeScript.

import { execFileSync, execSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { createRequire } from 'node:module';
import { platform, arch } from 'node:process';
import { familySync, MUSL } from 'detect-libc';

function getBinaryPath() {
  // 1. Check WITAN_BINARY environment variable
  if (process.env.WITAN_BINARY && existsSync(process.env.WITAN_BINARY)) {
    return process.env.WITAN_BINARY;
  }

  // 2. Resolve from platform-specific package
  const require = createRequire(import.meta.url);
  let suffix = `${platform}-${arch}`;
  if (platform === 'linux' && familySync() === MUSL) {
    suffix += '-musl';
  }
  const pkg = `@witan/${suffix}`;
  const bin = platform === 'win32' ? 'witan.exe' : 'witan';

  try {
    const pkgPath = dirname(require.resolve(`${pkg}/package.json`));
    const binaryPath = join(pkgPath, 'bin', bin);
    if (existsSync(binaryPath)) {
      return binaryPath;
    }
  } catch {
    // Package not installed
  }

  // 3. Check PATH
  try {
    const cmd = platform === 'win32' ? 'where witan' : 'which witan';
    const result = execSync(cmd, { encoding: 'utf8', stdio: ['pipe', 'pipe', 'ignore'] });
    const found = result.trim().split('\n')[0];
    if (found && existsSync(found)) {
      return found;
    }
  } catch {
    // Not on PATH
  }

  throw new Error(
    `witan binary not found. Install the package for your platform (${pkg}), ` +
    `set WITAN_BINARY environment variable, or add witan to PATH.`
  );
}

try {
  execFileSync(getBinaryPath(), process.argv.slice(2), { stdio: 'inherit' });
} catch (err) {
  if (err.status != null) {
    process.exit(err.status);
  }
  console.error(err.message);
  process.exit(1);
}
