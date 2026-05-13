import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { existsSync } from 'node:fs';
import { execSync } from 'node:child_process';

// Mock modules before importing the module under test
vi.mock('node:fs', async () => {
  const actual = await vi.importActual<typeof import('node:fs')>('node:fs');
  return {
    ...actual,
    existsSync: vi.fn(),
  };
});

vi.mock('node:child_process', async () => {
  const actual = await vi.importActual<typeof import('node:child_process')>('node:child_process');
  return {
    ...actual,
    execSync: vi.fn(),
  };
});

vi.mock('node:module', async () => {
  const actual = await vi.importActual<typeof import('node:module')>('node:module');
  return {
    ...actual,
    createRequire: vi.fn(() => {
      const mockRequire = Object.assign(
        () => {
          throw new Error('Cannot find module');
        },
        {
          resolve: () => {
            throw new Error('Cannot find module');
          },
        }
      );
      return mockRequire;
    }),
  };
});

// Import after mocks are set up
import { getBinaryPath, getPlatformPackageName } from '../src/binary.js';

const mockExistsSync = vi.mocked(existsSync);
const mockExecSync = vi.mocked(execSync);

describe('getPlatformPackageName', () => {
  it('returns platform-arch format', () => {
    const name = getPlatformPackageName();
    expect(name).toMatch(/^@witan\/(darwin|linux|win32)-(x64|arm64)(-musl)?$/);
  });
});

describe('getBinaryPath', () => {
  const originalEnv = process.env;

  beforeEach(() => {
    vi.resetAllMocks();
    process.env = { ...originalEnv };
    delete process.env['WITAN_BINARY'];
  });

  afterEach(() => {
    process.env = originalEnv;
  });

  it('returns WITAN_BINARY env var when set and file exists', () => {
    process.env['WITAN_BINARY'] = '/custom/path/witan';
    mockExistsSync.mockReturnValue(true);

    const result = getBinaryPath();

    expect(result).toBe('/custom/path/witan');
    expect(mockExistsSync).toHaveBeenCalledWith('/custom/path/witan');
  });

  it('ignores WITAN_BINARY when file does not exist', () => {
    process.env['WITAN_BINARY'] = '/nonexistent/witan';
    mockExistsSync.mockImplementation((path) => {
      if (path === '/nonexistent/witan') return false;
      if (path === '/usr/bin/witan') return true;
      return false;
    });
    mockExecSync.mockReturnValue('/usr/bin/witan\n');

    const result = getBinaryPath();

    expect(result).toBe('/usr/bin/witan');
  });

  it('falls back to PATH when package not installed', () => {
    mockExistsSync.mockImplementation((path) => {
      if (path === '/usr/local/bin/witan') return true;
      return false;
    });
    mockExecSync.mockReturnValue('/usr/local/bin/witan\n');

    const result = getBinaryPath();

    expect(result).toBe('/usr/local/bin/witan');
  });

  it('throws descriptive error when binary not found', () => {
    mockExistsSync.mockReturnValue(false);
    mockExecSync.mockImplementation(() => {
      throw new Error('not found');
    });

    expect(() => getBinaryPath()).toThrow(/witan binary not found/);
    expect(() => getBinaryPath()).toThrow(/@witan\//);
    expect(() => getBinaryPath()).toThrow(/WITAN_BINARY/);
  });

  it('handles which command returning multiple paths', () => {
    mockExistsSync.mockImplementation((path) => {
      if (path === '/first/witan') return true;
      return false;
    });
    mockExecSync.mockReturnValue('/first/witan\n/second/witan\n');

    const result = getBinaryPath();

    expect(result).toBe('/first/witan');
  });

  it('skips PATH result if file does not exist', () => {
    mockExistsSync.mockReturnValue(false);
    mockExecSync.mockReturnValue('/ghost/witan\n');

    expect(() => getBinaryPath()).toThrow(/witan binary not found/);
  });
});
