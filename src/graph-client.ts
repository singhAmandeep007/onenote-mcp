import { Client } from "@microsoft/microsoft-graph-client";

/**
 * A function that returns a valid access token, refreshing silently if needed.
 */
export type TokenProvider = () => Promise<string>;

/**
 * Build a Microsoft Graph client that acquires a fresh token for every request
 * via the provided token provider. This means expired tokens are automatically
 * renewed without any user interaction.
 */
export function createGraphClient(tokenProvider: TokenProvider): Client {
  return Client.initWithMiddleware({
    authProvider: {
      getAccessToken: tokenProvider,
    },
  });
}
