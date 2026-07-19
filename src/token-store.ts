import fs from 'node:fs';
import { TOKEN_FILE_PATH } from './config.js';

/** Trim whitespace and strip an optional leading "Bearer " from a token string. */
export function normalizeAccessToken(value: string | null | undefined): string | null {
  if (!value) return null;
  let token = String(value).trim();
  const lower = token.toLowerCase();
  if (lower.startsWith('bearer ') || lower === 'bearer') {
    token = token.slice(lower.startsWith('bearer ') ? 7 : 6).trim();
  }
  return token.length > 0 ? token : null;
}

/**
 * Whether a token looks like a JWT (three base64url segments).
 *
 * Work/school (Azure AD) accounts receive JWT access tokens; personal Microsoft
 * accounts receive a compact, non-JWT token. BOTH are valid for Microsoft Graph, so
 * this is only used for diagnostics and messaging, never to reject a token.
 */
export function isLikelyJwt(token: string | null | undefined): boolean {
  return typeof token === 'string' && token.split('.').length === 3;
}

/**
 * Load the cached access token, or null if none is available.
 * Reads the token file (JSON `{ "token": "..." }` or a raw string), falling back to
 * the GRAPH_ACCESS_TOKEN environment variable.
 */
export function loadToken(tokenPath: string = TOKEN_FILE_PATH): string | null {
  if (fs.existsSync(tokenPath)) {
    const raw = fs.readFileSync(tokenPath, 'utf8');
    let value: string;
    try {
      value = (JSON.parse(raw) as { token?: string }).token ?? '';
    } catch {
      value = raw;
    }
    const normalized = normalizeAccessToken(value);
    if (normalized) return normalized;
  }
  return normalizeAccessToken(process.env.GRAPH_ACCESS_TOKEN);
}

/** Persist an access token to disk with owner-only permissions. */
export function saveToken(token: string, tokenPath: string = TOKEN_FILE_PATH): void {
  const normalized = normalizeAccessToken(token);
  if (!normalized) {
    throw new Error('Refusing to save an empty access token.');
  }
  fs.writeFileSync(tokenPath, JSON.stringify({ token: normalized }), { mode: 0o600 });
}
