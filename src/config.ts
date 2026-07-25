import path from "node:path";
import { fileURLToPath } from "node:url";

const moduleDir = path.dirname(fileURLToPath(import.meta.url));

/**
 * Client ID for Microsoft Graph Command Line Tools — a first-party Microsoft
 * application (used by the Graph PowerShell SDK). Pre-consented for delegated
 * Graph scopes, so no Azure app registration is required. Override with
 * the GRAPH_CLIENT_ID environment variable if you register your own app.
 */
export const CLIENT_ID = process.env.GRAPH_CLIENT_ID ?? "14d82eec-204b-4c2f-b7e8-296a70dab67e";

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
export const TENANT_ID = process.env.GRAPH_TENANT ?? "common";

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
  "https://graph.microsoft.com/Notes.Read",
  "https://graph.microsoft.com/Notes.ReadWrite",
  "https://graph.microsoft.com/User.Read",
  "offline_access", // required for refresh token (MSAL added this implicitly)
  "openid", // required for id_token (username extraction)
];

/** Base URL for Microsoft Graph v1.0. */
export const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

/**
 * Where the token cache is stored. Contains access tokens, refresh tokens,
 * and account info. Secured with chmod 600 (owner-only).
 *
 * The refresh token lets the server silently renew access tokens without
 * user interaction, so auth only needs to happen once — not every hour.
 */
export const CACHE_FILE_PATH = process.env.ONENOTE_CACHE_PATH ?? path.join(moduleDir, "..", ".token-cache.json");
