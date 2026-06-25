/**
 * Witan Node.js SDK
 *
 * A Promise-based SDK for reading, writing, and manipulating Excel files
 * using the Witan CLI.
 *
 * @example
 * ```typescript
 * import { Workbook } from 'witan';
 *
 * // Open a workbook with automatic cleanup
 * {
 *   await using wb = await Workbook.open('report.xlsx');
 *   const sheets = await wb.listSheets();
 *   const data = await wb.readRange('Sheet1!A1:D10');
 *   await wb.save();
 * }
 * ```
 *
 * @packageDocumentation
 */

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
  isGoogleAuthRequired,
  isNeedsFileAuthorization,
} from './errors.js';

export { serializeMatcher, dropUndefined } from './helpers.js';

export { StdioRPCProcess } from './process.js';

export { Workbook, type WorkbookOptions } from './workbook.js';

export {
  GoogleSheet,
  DEFAULT_GOOGLE_SHEET_REQUEST_TIMEOUT_MS,
  type GoogleSheetOptions,
  type GoogleSheetOpenOptions,
  type GoogleSheetCreateOptions,
  type GoogleSheetAuthorizeOptions,
  type AuthorizeUrlResult,
} from './google-sheet.js';

export type * from './types.js';
