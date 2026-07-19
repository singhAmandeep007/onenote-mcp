// Barrel export — public API surface for programmatic use.

export { CLIENT_ID, TENANT_ID, SCOPES, GRAPH_BASE_URL, TOKEN_FILE_PATH } from './config.js';
export { log } from './logger.js';
export { normalizeAccessToken, isLikelyJwt, loadToken, saveToken } from './token-store.js';
export { authenticateWithDeviceCode } from './auth.js';
export type { AuthResult } from './auth.js';
export { createGraphClient } from './graph-client.js';
export { htmlToText, escapeHtml } from './html.js';
export { OneNoteClient } from './onenote.js';
export type { PageContent, UserProfile } from './onenote.js';
