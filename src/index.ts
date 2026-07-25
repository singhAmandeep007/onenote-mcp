// Barrel export — public API surface for programmatic use.

export { CLIENT_ID, TENANT_ID, SCOPES, GRAPH_BASE_URL, CACHE_FILE_PATH } from "./config.js";
export { log } from "./logger.js";
export { normalizeAccessToken, isLikelyJwt } from "./token-store.js";
export {
  authenticateWithDeviceCode,
  acquireTokenSilent,
  getAccessToken,
  hasCachedAccount,
  clearCache,
} from "./auth.js";
export type { AuthenticationResult } from "./auth.js";
export { createGraphClient } from "./graph-client.js";
export type { TokenProvider } from "./graph-client.js";
export { htmlToText, escapeHtml } from "./html.js";
export { OneNoteClient } from "./onenote.js";
export type { PageContent, UserProfile } from "./onenote.js";
