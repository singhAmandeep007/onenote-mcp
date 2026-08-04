# OneNote MCP Server

A TypeScript [Model Context Protocol](https://modelcontextprotocol.io) (MCP) server that gives AI assistants (Claude, Cursor, etc.) read/write access to your Microsoft OneNote notebooks via the Microsoft Graph API.

**Zero Azure setup required.** Authentication uses the [device-code flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-device-code) with a first-party Microsoft client ID, so you only need a Microsoft account. Sign in once and tokens renew automatically — you won't be asked to authenticate again.

> Fork of [azure-onenote-mcp-server](https://github.com/ZubeidHendricks/azure-onenote-mcp-server) by Zubeid Hendricks — rewritten in TypeScript with proper auth, typed tools, and a unified CLI.


## Article

[🚀 https://dev.to/singhamandeep007/i-rewrote-a-onenote-mcp-server-in-typescript-heres-what-i-learned-about-microsoft-graph-auth-5933](https://dev.to/singhamandeep007/i-rewrote-a-onenote-mcp-server-in-typescript-heres-what-i-learned-about-microsoft-graph-auth-5933)

## Diagram

![OAuth Flow](./OAuth%20device-code%20flow.png)

![Request Lifecycle](./request%20lifecycle.png)

## Features

- **One-time authentication** — sign in once via device-code flow; tokens silently refresh using a stored refresh token (access tokens expire hourly, but this is invisible to you)
- Works with both **personal Microsoft accounts** and **work/school (Entra ID)** accounts
- List notebooks, sections, and pages
- Read page content as plain text (HTML is stripped automatically)
- Create pages with HTML content
- Search pages by title across all notebooks
- Zod schemas on every MCP tool — AI clients get proper parameter descriptions and types
- Unified CLI for scripting and debugging (`onenote-cli`)
- Zero heavy dependencies — auth is ~180 lines of direct OAuth 2.0, no MSAL required

## Prerequisites

- **Node.js >= 18.18** (`.nvmrc` pins v24.14.1)
- A Microsoft account with access to OneNote

## Quick Start

```bash
git clone https://github.com/singhAmandeep007/onenote-mcp.git
cd onenote-mcp
npm install
npm run auth        # sign in with your Microsoft account (one-time)
npm run verify      # confirm the token works
```

`npm run auth` opens a device-code flow — it prints a URL and a code. Open the URL in any browser, enter the code, sign in, and you're done. The server stores a refresh token locally so it can silently renew access tokens without asking you again.

## MCP Server Configuration

Add the server to your AI assistant's MCP config. The server communicates over stdio (JSON-RPC on stdout, logs on stderr).

**Claude Desktop** (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "onenote": {
      "command": "npx",
      "args": ["tsx", "/absolute/path/to/onenote-mcp/src/mcp-server.ts"]
    }
  }
}
```

**Cursor** (Settings → MCP):

```json
{
  "mcpServers": {
    "onenote": {
      "command": "npx",
      "args": ["tsx", "/absolute/path/to/onenote-mcp/src/mcp-server.ts"]
    }
  }
}
```

This runs the TypeScript source directly via `tsx` — no build step needed. If you prefer compiled output, run `npm run build` first, then use `"command": "node"` with `"args": ["/absolute/path/to/onenote-mcp/dist/mcp-server.js"]`.

After updating your config, restart your AI assistant so it picks up the new server. Try asking: *"What notebooks do I have in OneNote?"*

## Available MCP Tools

| Tool | Parameters | Description |
|------|-----------|-------------|
| `authenticate` | — | Start device-code sign-in (one-time) |
| `listNotebooks` | — | List all notebooks |
| `getNotebook` | `notebookId` | Get notebook details by ID |
| `listSections` | `notebookId?` | List sections (all, or within a notebook) |
| `listPages` | `sectionId?` | List pages (all, or within a section) |
| `getPage` | `query` | Get page content by ID or title search |
| `createPage` | `title`, `bodyHtml`, `sectionId?` | Create a page in a section |
| `searchPages` | `query` | Search pages by title |

## CLI

A single CLI replaces the standalone scripts from the original project. Run via `npm run cli` or directly with `npx tsx src/cli.ts`.

```bash
# Auth
npm run auth                         # sign in (one-time)
npm run verify                       # verify token against Graph API
npm run logout                       # clear tokens and sign out

# Browse
npm run cli -- notebooks             # list notebooks
npm run cli -- sections              # list all sections
npm run cli -- sections --notebook <id>
npm run cli -- pages                 # list all pages
npm run cli -- pages --section <id>

# Read & write
npm run cli -- get "page title"      # get page content as text
npm run cli -- create --title "My Page" --body "<p>Hello</p>"
npm run cli -- search "meeting"      # search by title
```

## npm Scripts

| Script | Description |
|--------|-------------|
| `npm run build` | Compile TypeScript to `dist/` |
| `npm start` | Run the compiled MCP server |
| `npm run dev` | Run the MCP server via `tsx` (no build needed) |
| `npm run auth` | Device-code authentication (one-time) |
| `npm run verify` | Verify stored token against Graph API |
| `npm run logout` | Clear token cache and sign out |
| `npm run cli` | Unified CLI (see above) |
| `npm run typecheck` | Type-check without emitting |
| `npm test` | Run tests with Vitest |
| `npm run test:watch` | Run tests in watch mode |

## Project Structure

```
src/
  config.ts         — Client ID, tenant, scopes, paths
  logger.ts         — stderr-only logging (MCP uses stdout for JSON-RPC)
  token-store.ts    — Token normalization utilities
  auth.ts           — OAuth 2.0 device-code auth + refresh token renewal
  graph-client.ts   — Microsoft Graph SDK client factory
  html.ts           — HTML-to-text conversion (zero dependencies)
  onenote.ts        — Typed OneNote client (notebooks, sections, pages)
  mcp-server.ts     — MCP server entry point with Zod-typed tools
  cli.ts            — Unified CLI (replaces ~10 standalone scripts)
  index.ts          — Barrel export for programmatic use
tests/
  html.test.ts      — HTML-to-text edge cases
  token-store.test.ts — Token normalization
  onenote.test.ts   — OneNote client with mocked Graph responses
```

## Authentication Details

Authentication uses Microsoft's [device-code flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-device-code) with the [Microsoft Graph Command Line Tools](https://learn.microsoft.com/en-us/answers/questions/1619076/microsoft-graph-command-line-tools-enterprise-appl) client ID (`14d82eec-...`). This is a first-party Microsoft application used by the official Graph PowerShell SDK — it's pre-consented for delegated Graph scopes, so no Azure app registration is needed.

The server requests these delegated scopes:

- `Notes.Read` — read notebooks, sections, and pages
- `Notes.ReadWrite` — create and modify pages
- `User.Read` — verify the signed-in identity
- `offline_access` — obtain a refresh token for silent renewal
- `openid` — obtain an ID token for username extraction

These are **delegated, non-`.All`** scopes, which is important: personal Microsoft accounts cannot consent to `.All` (application-level) scopes. Requesting `.All` scopes produces a token that Graph silently rejects with HTTP 401 — no error during auth, just a broken token.

### Token Lifecycle

After the initial device-code sign-in, the server persists both the access token and a refresh token in `.token-cache.json` (chmod 600, gitignored). When the access token expires (~1 hour), it silently acquires a new one using the refresh token — no user interaction required. The refresh token itself lasts ~90 days.

You only need to re-authenticate if you explicitly sign out (`npm run logout`) or if your refresh token is revoked by an admin.

### Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `GRAPH_CLIENT_ID` | Microsoft Graph Command Line Tools ID | Use your own Azure app registration |
| `GRAPH_TENANT` | `common` (personal + work/school) | `consumers`, `organizations`, or a specific tenant ID |
| `ONENOTE_CACHE_PATH` | `.token-cache.json` | Custom token cache file location |

## Troubleshooting

**401 / error code 40001** — You're likely requesting `.All` scopes (only an issue if you've overridden the defaults), or your refresh token was revoked. Run `npm run logout` then `npm run auth`.

**"Token is a compact (non-JWT) token"** — Normal for personal Microsoft accounts. Personal accounts return opaque (non-JWT) access tokens that are perfectly valid for Graph API calls. Don't validate token format — let Graph be the authority.

**Server output corrupts the MCP stream** — All logging must go to stderr. If you add `console.log()` calls, the JSON-RPC protocol over stdout breaks. Use `log()` from `src/logger.ts` instead.

**AADSTS50105 "application blocked"** — Your organization's admin has blocked the Microsoft Graph Command Line Tools app. Set `GRAPH_CLIENT_ID` to your own Azure app registration's client ID.

## Security

- Token cache is stored locally with `chmod 600` (owner-only read/write)
- No credentials are sent to any third party — auth goes directly to `login.microsoftonline.com`
- The public client ID has no client secret; security relies entirely on the user completing the device-code flow in their browser
- `.token-cache.json` is in `.gitignore`

## Credits

Built on [azure-onenote-mcp-server](https://github.com/ZubeidHendricks/azure-onenote-mcp-server) by Zubeid Hendricks.

## License

MIT — see [LICENSE](LICENSE) for details.
