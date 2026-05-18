import { existsSync } from 'node:fs';
import { platform, arch } from 'node:process';
import { execSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { familySync, MUSL } from 'detect-libc';

const require = createRequire(import.meta.url);

/**
 * Get the platform-specific package name for the witan binary.
 * Handles Linux glibc vs musl detection.
 */
export function getPlatformPackageName(): string {
  let suffix = `${platform}-${arch}`;

  // On Linux, detect glibc vs musl
  if (platform === 'linux' && familySync() === MUSL) {
    suffix += '-musl';
  }

  return `@witan/${suffix}`;
}

/**
 * Find the witan binary path.
 *
 * Resolution order:
 * 1. WITAN_BINARY environment variable
 * 2. Platform-specific npm package (@witan/darwin-arm64, etc.)
 * 3. System PATH
 *
 * @throws Error if binary cannot be found
 */
export function getBinaryPath(): string {
  // 1. Check WITAN_BINARY environment variable (matches Python behavior)
  const envPath = process.env['WITAN_BINARY'];
  if (envPath && existsSync(envPath)) {
    return envPath;
  }

  // 2. Try to resolve from platform-specific package
  const packageName = getPlatformPackageName();
  try {
    const packagePath = dirname(require.resolve(`${packageName}/package.json`));
    const binaryName = platform === 'win32' ? 'witan.exe' : 'witan';
    const binaryPath = join(packagePath, 'bin', binaryName);
    if (existsSync(binaryPath)) {
      return binaryPath;
    }
  } catch {
    // Package not installed
  }

  // 3. Check PATH (matches Python's shutil.which() fallback)
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
    `witan binary not found. Install the package for your platform (${packageName}), ` +
    `set WITAN_BINARY environment variable, or add witan to PATH.`
  );
}
