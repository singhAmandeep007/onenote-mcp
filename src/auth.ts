/**
 * OAuth 2.0 authentication — direct implementation using native fetch.
 *
 * Implements the device-code flow (initial sign-in) and refresh-token
 * exchange (silent renewal) against Azure AD's v2.0 endpoints. Tokens
 * are persisted to a JSON file with owner-only permissions.
 *
 * This replaces the MSAL Node dependency, which had a v5 regression
 * that omitted `client_id` from refresh-token POST bodies (AADSTS900144).
 */

import fs from "node:fs";
import { CLIENT_ID, TENANT_ID, SCOPES, CACHE_FILE_PATH } from "./config.js";
import { log } from "./logger.js";

// ── Token cache ──────────────────────────────────────────────────────────────

interface TokenCache {
  access_token: string;
  refresh_token?: string;
  /** Unix timestamp (seconds) when the access token expires. */
  expires_at: number;
  account?: { username: string };
}

function readCache(): TokenCache | null {
  if (!fs.existsSync(CACHE_FILE_PATH)) return null;
  try {
    return JSON.parse(fs.readFileSync(CACHE_FILE_PATH, "utf8")) as TokenCache;
  } catch {
    return null;
  }
}

function writeCache(cache: TokenCache): void {
  fs.writeFileSync(CACHE_FILE_PATH, JSON.stringify(cache, null, 2), {
    mode: 0o600,
  });
}

// ── Endpoints ────────────────────────────────────────────────────────────────

const TOKEN_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
const DEVICE_CODE_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/devicecode`;

// ── Refresh token exchange ───────────────────────────────────────────────────

async function refreshAccessToken(refreshToken: string): Promise<TokenCache | null> {
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: SCOPES.join(" "),
  });

  const res = await fetch(TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });

  if (!res.ok) {
    log("Token refresh failed:", res.status, await res.text());
    return null;
  }

  const data = (await res.json()) as {
    access_token: string;
    refresh_token?: string;
    expires_in: number;
  };

  return {
    access_token: data.access_token,
    refresh_token: data.refresh_token ?? refreshToken,
    expires_at: Math.floor(Date.now() / 1000) + data.expires_in,
  };
}

// ── Silent acquisition ───────────────────────────────────────────────────────

/**
 * Get a valid access token from the cache, refreshing silently if needed.
 * Returns `null` if no cached session exists or the refresh token is expired.
 */
export async function acquireTokenSilent(): Promise<string | null> {
  const cache = readCache();
  if (!cache) return null;

  // Access token still valid (with 5-minute buffer)?
  const now = Math.floor(Date.now() / 1000);
  if (cache.access_token && cache.expires_at > now + 300) {
    return cache.access_token;
  }

  // Need to refresh
  if (!cache.refresh_token) return null;
  log("Access token expired, refreshing...");

  const refreshed = await refreshAccessToken(cache.refresh_token);
  if (!refreshed) return null;

  refreshed.account = cache.account;
  writeCache(refreshed);
  log("Token refreshed successfully.");
  return refreshed.access_token;
}

// ── Device-code flow ─────────────────────────────────────────────────────────

/** Subset of MSAL's AuthenticationResult — keeps callers compatible. */
export interface AuthenticationResult {
  accessToken: string;
  account?: { username: string };
}

/**
 * Run the interactive device-code flow. The user visits a URL and enters a
 * code; we poll until Azure AD issues tokens. The result is persisted to the
 * cache file so that `acquireTokenSilent` works for all subsequent calls.
 */
export async function authenticateWithDeviceCode(
  prompt: (message: string) => void = (message) => log(message)
): Promise<AuthenticationResult> {
  // Step 1 — request a device code
  const dcBody = new URLSearchParams({
    client_id: CLIENT_ID,
    scope: SCOPES.join(" "),
  });

  const dcRes = await fetch(DEVICE_CODE_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: dcBody.toString(),
  });

  if (!dcRes.ok) {
    const text = await dcRes.text();
    throw new Error(`Device code request failed (${dcRes.status}): ${text}`);
  }

  const dc = (await dcRes.json()) as {
    device_code: string;
    user_code: string;
    verification_uri: string;
    message: string;
    interval: number;
    expires_in: number;
  };

  prompt(dc.message);

  // Step 2 — poll for token
  const interval = (dc.interval || 5) * 1000;
  const deadline = Date.now() + dc.expires_in * 1000;

  while (Date.now() < deadline) {
    await new Promise((r) => setTimeout(r, interval));

    const pollBody = new URLSearchParams({
      client_id: CLIENT_ID,
      grant_type: "urn:ietf:params:oauth:grant-type:device_code",
      device_code: dc.device_code,
    });

    const pollRes = await fetch(TOKEN_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: pollBody.toString(),
    });

    const poll = (await pollRes.json()) as {
      access_token?: string;
      refresh_token?: string;
      expires_in?: number;
      id_token?: string;
      error?: string;
    };

    if (poll.error === "authorization_pending") continue;
    if (poll.error === "slow_down") {
      await new Promise((r) => setTimeout(r, 5000));
      continue;
    }
    if (poll.error) {
      throw new Error(`Authentication failed: ${poll.error}`);
    }
    if (!poll.access_token) {
      throw new Error("Token response missing access_token.");
    }
    if (!poll.refresh_token) {
      log(
        "Warning: no refresh_token returned — silent renewal won't work. Ensure 'offline_access' scope is requested."
      );
    }

    // Extract username from id_token JWT (if present)
    let username = "unknown";
    if (poll.id_token) {
      try {
        const payload = JSON.parse(Buffer.from(poll.id_token.split(".")[1]!, "base64url").toString()) as {
          preferred_username?: string;
          email?: string;
          name?: string;
        };
        username = payload.preferred_username ?? payload.email ?? payload.name ?? "unknown";
      } catch {
        /* ignore decode errors */
      }
    }

    const tokenCache: TokenCache = {
      access_token: poll.access_token,
      refresh_token: poll.refresh_token,
      expires_at: Math.floor(Date.now() / 1000) + (poll.expires_in ?? 3600),
      account: { username },
    };
    writeCache(tokenCache);

    log(`Authenticated as: ${username}`);
    return { accessToken: poll.access_token, account: { username } };
  }

  throw new Error("Device code expired — user did not complete sign-in in time.");
}

// ── Convenience exports ──────────────────────────────────────────────────────

/**
 * Get a valid access token — silently if possible, otherwise throws.
 * This is the main entry point used by the Graph client.
 */
export async function getAccessToken(): Promise<string> {
  const token = await acquireTokenSilent();
  if (token) return token;
  throw new Error("No valid token available. Run 'npm run auth' (or the 'authenticate' MCP tool) to sign in.");
}

/** Whether there's a cached session (user has authenticated at least once). */
export async function hasCachedAccount(): Promise<boolean> {
  return readCache() !== null;
}

/** Clear the token cache (sign out). */
export async function clearCache(): Promise<void> {
  if (fs.existsSync(CACHE_FILE_PATH)) {
    fs.unlinkSync(CACHE_FILE_PATH);
  }
}
