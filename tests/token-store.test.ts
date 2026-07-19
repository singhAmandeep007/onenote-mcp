import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { normalizeAccessToken, isLikelyJwt, loadToken, saveToken } from '../src/token-store.js';

describe('normalizeAccessToken', () => {
  it('returns null for falsy values', () => {
    expect(normalizeAccessToken(null)).toBeNull();
    expect(normalizeAccessToken(undefined)).toBeNull();
    expect(normalizeAccessToken('')).toBeNull();
  });

  it('trims whitespace', () => {
    expect(normalizeAccessToken('  abc  ')).toBe('abc');
  });

  it('strips Bearer prefix (case-insensitive)', () => {
    expect(normalizeAccessToken('Bearer xyz')).toBe('xyz');
    expect(normalizeAccessToken('bearer xyz')).toBe('xyz');
    expect(normalizeAccessToken('BEARER xyz')).toBe('xyz');
  });

  it('returns null for empty-after-strip', () => {
    expect(normalizeAccessToken('Bearer ')).toBeNull();
  });
});

describe('isLikelyJwt', () => {
  it('returns true for three-segment tokens', () => {
    expect(isLikelyJwt('aaa.bbb.ccc')).toBe(true);
  });

  it('returns false for compact tokens', () => {
    expect(isLikelyJwt('compacttoken')).toBe(false);
  });

  it('returns false for null/undefined', () => {
    expect(isLikelyJwt(null)).toBe(false);
    expect(isLikelyJwt(undefined)).toBe(false);
  });
});

describe('loadToken / saveToken', () => {
  let tmpFile: string;

  beforeEach(() => {
    tmpFile = path.join(os.tmpdir(), `test-token-${Date.now()}.txt`);
  });

  afterEach(() => {
    try { fs.unlinkSync(tmpFile); } catch { /* ok */ }
  });

  it('round-trips a token through save and load', () => {
    saveToken('my-secret-token', tmpFile);
    expect(loadToken(tmpFile)).toBe('my-secret-token');
  });

  it('saves as JSON with owner-only permissions (unix)', () => {
    saveToken('tok123', tmpFile);
    const raw = fs.readFileSync(tmpFile, 'utf8');
    expect(JSON.parse(raw)).toEqual({ token: 'tok123' });
    if (process.platform !== 'win32') {
      const stat = fs.statSync(tmpFile);
      expect((stat.mode & 0o777).toString(8)).toBe('600');
    }
  });

  it('loads raw-string token files (legacy format)', () => {
    fs.writeFileSync(tmpFile, '  raw-token-value  ');
    expect(loadToken(tmpFile)).toBe('raw-token-value');
  });

  it('returns null when file does not exist', () => {
    expect(loadToken('/tmp/nonexistent-token-file-12345.txt')).toBeNull();
  });

  it('refuses to save an empty token', () => {
    expect(() => saveToken('', tmpFile)).toThrow();
  });
});
