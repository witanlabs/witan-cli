import { describe, it, expect } from 'vitest';
import {
  WitanError,
  WitanProcessError,
  WitanRPCError,
  WitanTimeoutError,
  isGoogleAuthRequired,
  isNeedsFileAuthorization,
} from '../src/errors.js';

describe('WitanError', () => {
  it('has correct name and message', () => {
    const err = new WitanError('test message');
    expect(err.name).toBe('WitanError');
    expect(err.message).toBe('test message');
    expect(err).toBeInstanceOf(Error);
    expect(err).toBeInstanceOf(WitanError);
  });
});

describe('WitanProcessError', () => {
  it('has correct name and message without stderr', () => {
    const err = new WitanProcessError('process failed');
    expect(err.name).toBe('WitanProcessError');
    expect(err.message).toBe('process failed');
    expect(err.stderrTail).toEqual([]);
    expect(err).toBeInstanceOf(WitanError);
  });

  it('includes stderr tail in message', () => {
    const stderr = ['line 1', 'line 2'];
    const err = new WitanProcessError('process failed', stderr);
    expect(err.message).toBe('process failed\nstderr tail:\nline 1\nline 2');
    expect(err.stderrTail).toEqual(stderr);
  });
});

describe('WitanTimeoutError', () => {
  it('extends WitanProcessError', () => {
    const err = new WitanTimeoutError('timed out');
    expect(err.name).toBe('WitanTimeoutError');
    expect(err).toBeInstanceOf(WitanProcessError);
    expect(err).toBeInstanceOf(WitanError);
  });

  it('includes stderr tail', () => {
    const stderr = ['timeout stderr'];
    const err = new WitanTimeoutError('timed out', stderr);
    expect(err.stderrTail).toEqual(stderr);
  });
});

describe('WitanRPCError', () => {
  it('has correct name and properties', () => {
    const err = new WitanRPCError('not found', {
      method: 'readCell',
      op: 'readRange',
      requestId: '42',
      code: 'NOT_FOUND',
      response: { extra: 'data' },
    });

    expect(err.name).toBe('WitanRPCError');
    expect(err.message).toBe('NOT_FOUND: not found');
    expect(err.method).toBe('readCell');
    expect(err.op).toBe('readRange');
    expect(err.requestId).toBe('42');
    expect(err.code).toBe('NOT_FOUND');
    expect(err.response).toEqual({ extra: 'data' });
    expect(err).toBeInstanceOf(WitanError);
  });

  it('uses default code when not provided', () => {
    const err = new WitanRPCError('generic error', {
      method: 'test',
      op: 'testOp',
      requestId: '1',
    });

    expect(err.message).toBe('RPC_ERROR: generic error');
    expect(err.code).toBeNull();
    expect(err.response).toEqual({});
  });
});

describe('isGoogleAuthRequired', () => {
  it('returns true for google_auth_required rpc errors', () => {
    const err = new WitanRPCError('auth', {
      method: 'listSheets',
      op: 'listSheets',
      requestId: '1',
      code: 'google_auth_required',
    });
    expect(isGoogleAuthRequired(err)).toBe(true);
  });

  it('returns true for auth failures surfaced during subprocess startup', () => {
    const err = new WitanProcessError('witan subprocess exited during startup (exit=1)', [
      "Google Sheets requires authorization. Run 'witan gsheets connect' to enable access.",
    ]);
    expect(isGoogleAuthRequired(err)).toBe(true);
  });

  it('returns false for other rpc errors', () => {
    const err = new WitanRPCError('not found', {
      method: 'readCell',
      op: 'readRange',
      requestId: '42',
      code: 'NOT_FOUND',
    });
    expect(isGoogleAuthRequired(err)).toBe(false);
  });

  it('returns false for other process errors', () => {
    const err = new WitanProcessError('witan subprocess exited before responding (exit=1)');
    expect(isGoogleAuthRequired(err)).toBe(false);
  });

  it('returns true for not-connected failures from authorize/status', () => {
    const err = new WitanProcessError('witan gsheets authorize failed', [
      "Google Sheets is not connected. Run 'witan gsheets connect' first.",
    ]);
    expect(isGoogleAuthRequired(err)).toBe(true);
  });

  it('returns true for expired/revoked failures from authorize/status', () => {
    const err = new WitanProcessError('witan gsheets authorize failed', [
      "Google authorization expired or was revoked. Run 'witan gsheets connect' to reconnect.",
    ]);
    expect(isGoogleAuthRequired(err)).toBe(true);
  });
});

describe('isNeedsFileAuthorization', () => {
  it('returns true for needs_file_authorization rpc errors', () => {
    const err = new WitanRPCError('nope', {
      method: 'listSheets',
      op: 'listSheets',
      requestId: '1',
      code: 'needs_file_authorization',
    });
    expect(isNeedsFileAuthorization(err)).toBe(true);
  });

  it('returns true for auth failures surfaced during subprocess startup', () => {
    const err = new WitanProcessError('witan subprocess exited before responding (exit=3)', [
      'This Google Sheet must be authorized before Witan can access it (it may also not exist).',
    ]);
    expect(isNeedsFileAuthorization(err)).toBe(true);
  });

  it('returns false for unrelated errors', () => {
    expect(isNeedsFileAuthorization(new WitanProcessError('some other failure'))).toBe(false);
    expect(
      isNeedsFileAuthorization(
        new WitanRPCError('x', { method: 'm', op: 'o', requestId: '1', code: 'NOT_FOUND' })
      )
    ).toBe(false);
  });
});
