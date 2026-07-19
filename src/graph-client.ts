import { Client } from '@microsoft/microsoft-graph-client';

/**
 * Build a Microsoft Graph client that authenticates every request with a static
 * access token. Uses the middleware-based auth provider (the modern API).
 */
export function createGraphClient(token: string): Client {
  return Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => token,
    },
  });
}
