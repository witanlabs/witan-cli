import { describe, it, expect } from 'vitest';
import { serializeMatcher, dropUndefined } from '../src/helpers.js';

describe('serializeMatcher', () => {
  it('serializes RegExp to source/flags object', () => {
    const result = serializeMatcher(/hello/i);
    expect(result).toEqual({ source: 'hello', flags: 'i' });
  });

  it('serializes RegExp with multiple flags', () => {
    const result = serializeMatcher(/pattern/gim);
    expect(result).toEqual({ source: 'pattern', flags: 'gim' });
  });

  it('passes string through unchanged', () => {
    const result = serializeMatcher('hello');
    expect(result).toBe('hello');
  });

  it('serializes array of RegExp', () => {
    const result = serializeMatcher([/a/i, /b/g]);
    expect(result).toEqual([
      { source: 'a', flags: 'i' },
      { source: 'b', flags: 'g' },
    ]);
  });

  it('serializes array of strings', () => {
    const result = serializeMatcher(['foo', 'bar']);
    expect(result).toEqual(['foo', 'bar']);
  });

  it('serializes mixed array of strings and RegExp', () => {
    const result = serializeMatcher(['foo', /bar/i, 'baz']);
    expect(result).toEqual(['foo', { source: 'bar', flags: 'i' }, 'baz']);
  });
});

describe('dropUndefined', () => {
  it('removes undefined values', () => {
    const result = dropUndefined({ a: 1, b: undefined, c: 'hello' });
    expect(result).toEqual({ a: 1, c: 'hello' });
  });

  it('keeps null values', () => {
    const result = dropUndefined({ a: 1, b: null, c: undefined });
    expect(result).toEqual({ a: 1, b: null });
  });

  it('keeps falsy values except undefined', () => {
    const result = dropUndefined({ a: 0, b: '', c: false, d: undefined });
    expect(result).toEqual({ a: 0, b: '', c: false });
  });

  it('returns empty object for all undefined values', () => {
    const result = dropUndefined({ a: undefined, b: undefined });
    expect(result).toEqual({});
  });

  it('handles empty object', () => {
    const result = dropUndefined({});
    expect(result).toEqual({});
  });

  it('preserves nested objects', () => {
    const nested = { x: 1, y: 2 };
    const result = dropUndefined({ a: nested, b: undefined });
    expect(result).toEqual({ a: nested });
  });
});
