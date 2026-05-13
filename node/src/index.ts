// Witan Node.js SDK

import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const pkg = JSON.parse(readFileSync(join(__dirname, '..', 'package.json'), 'utf-8'));

/** Package version (matches Python's __version__) */
export const version: string = pkg.version;

export { getBinaryPath } from './binary.js';

export {
  WitanError,
  WitanProcessError,
  WitanRPCError,
  WitanTimeoutError,
} from './errors.js';

export { serializeMatcher, dropUndefined } from './helpers.js';

export { StdioRPCProcess } from './process.js';

export { Workbook, type WorkbookOptions } from './workbook.js';

export type * from './types.js';
