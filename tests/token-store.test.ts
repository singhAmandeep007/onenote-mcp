import { describe, it, expect } from 'vitest';
import { normalizeAccessToken, isLikelyJwt } from '../src/token-store.js';

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
