import type { Matcher } from './spreadsheet-base.js';

/**
 * Serialize a matcher for the RPC protocol.
 * Converts RegExp to { source, flags } format expected by the Go CLI.
 */
export function serializeMatcher(matcher: Matcher): unknown {
  if (matcher instanceof RegExp) {
    return { source: matcher.source, flags: matcher.flags };
  }

  if (Array.isArray(matcher)) {
    return matcher.map(m =>
      m instanceof RegExp
        ? { source: m.source, flags: m.flags }
        : m
    );
  }

  return matcher;
}

/**
 * Remove undefined values from an object (like Python's _drop_none).
 */
export function dropUndefined<T extends Record<string, unknown>>(obj: T): Partial<T> {
  const result: Partial<T> = {};
  for (const [key, value] of Object.entries(obj)) {
    if (value !== undefined) {
      (result as Record<string, unknown>)[key] = value;
    }
  }
  return result;
}
