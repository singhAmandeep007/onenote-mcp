import path from 'node:path';
import { fileURLToPath } from 'node:url';

const moduleDir = path.dirname(fileURLToPath(import.meta.url));

/**
 * Public client id of the Microsoft Graph Explorer app. It is pre-consented for
 * delegated Graph scopes, so no Azure app registration is required. Override with
 * the GRAPH_CLIENT_ID environment variable if you register your own app.
 */
export const CLIENT_ID = process.env.GRAPH_CLIENT_ID ?? '14d82eec-204b-4c2f-b7e8-296a70dab67e';

/**
 * Authentication authority/tenant.
 *
 * - `common`        -> personal Microsoft accounts (MSA) AND work/school (Azure AD)  [default]
 * - `consumers`     -> personal Microsoft accounts only
 * - `organizations` -> work/school accounts only
 * - `<tenant-id>`   -> a single specific organization
 *
 * Override with the GRAPH_TENANT environment variable.
 */
export const TENANT_ID = process.env.GRAPH_TENANT ?? 'common';

/**
 * Resource-qualified, delegated Graph scopes.
 *
 * IMPORTANT: these are the non-".All" delegated scopes. Personal Microsoft accounts
 * cannot consent to the ".All" application-style scopes, and requesting them makes the
 * device-code flow return a token Microsoft Graph rejects with HTTP 401 (code 40001).
 * The scopes below cover reading and writing the signed-in user's own OneNote content
 * and work for both personal and work/school accounts.
 */
export const SCOPES = [
  'https://graph.microsoft.com/Notes.Read',
  'https://graph.microsoft.com/Notes.ReadWrite',
  'https://graph.microsoft.com/User.Read',
];

/** Base URL for Microsoft Graph v1.0. */
export const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

/**
 * Where the cached access token is stored. Defaults to `.access-token.txt` in the
 * project root. Override with the ONENOTE_TOKEN_PATH environment variable.
 */
export const TOKEN_FILE_PATH =
  process.env.ONENOTE_TOKEN_PATH ?? path.join(moduleDir, '..', '.access-token.txt');
