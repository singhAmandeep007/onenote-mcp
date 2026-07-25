/**
 * Token utility functions.
 *
 * Token persistence is handled by auth.ts.
 * These utilities are kept for diagnostics, tests, and edge cases.
 */

/** Trim whitespace and strip an optional leading "Bearer " from a token string. */
export function normalizeAccessToken(value: string | null | undefined): string | null {
  if (!value) return null;
  let token = String(value).trim();
  const lower = token.toLowerCase();
  if (lower.startsWith("bearer ") || lower === "bearer") {
    token = token.slice(lower.startsWith("bearer ") ? 7 : 6).trim();
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
  return typeof token === "string" && token.split(".").length === 3;
}
