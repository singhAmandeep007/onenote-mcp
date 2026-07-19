# OneNote MCP Server

A TypeScript [Model Context Protocol](https://modelcontextprotocol.io) (MCP) server that gives AI assistants (Claude, Cursor, etc.) read/write access to your Microsoft OneNote notebooks via the Microsoft Graph API.

**Zero Azure setup required** — authentication uses the device-code flow with a pre-consented public client, so you only need a Microsoft account.

> Based on [azure-onenote-mcp-server](https://github.com/ZubeidHendricks/azure-onenote-mcp-server) by Zubeid Hendricks.

## Features

- Device-code authentication — works with both personal Microsoft accounts and work/school (Azure AD) accounts
- List notebooks, sections, and pages
- Read page content as plain text (HTML is stripped automatically)
- Create pages with HTML content
- Search pages by title across all notebooks
- Typed Zod schemas on every MCP tool for reliable AI integration
- Unified CLI for scripting and debugging (`onenote-cli`)

## Prerequisites

- **Node.js ≥ 18.18** (`.nvmrc` pins v24.14.1)
- A Microsoft account with access to OneNote

## Quick Start

```bash
git clone https://github.com/danosb/onenote-mcp.git
cd onenote-mcp
npm install
npm run build       # compile TypeScript
npm run auth        # sign in with your Microsoft account
npm run verify      # confirm the token works
```

## MCP Server Configuration

Add the server to your AI assistant's MCP config. The server communicates over stdio.

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

This runs the TypeScript source directly via `tsx` — no build step needed. If you prefer to use the compiled output instead, run `npm run build` first and use `"command": "node"` with `"args": ["/absolute/path/to/onenote-mcp/dist/mcp-server.js"]`.

After updating your config, restart your AI assistant (e.g. quit and reopen Claude Desktop) so it picks up the new MCP server. You can verify the server is working by asking Claude to list your notebooks.

## Available MCP Tools

| Tool | Parameters | Description |
|------|-----------|-------------|
| `authenticate` | — | Start device-code sign-in flow |
| `saveAccessToken` | `token` | Manually save an access token |
| `listNotebooks` | — | List all notebooks |
| `getNotebook` | `notebookId` | Get notebook details by ID |
| `listSections` | `notebookId?` | List sections (all, or within a notebook) |
| `listPages` | `sectionId?` | List pages (all, or within a section) |
| `getPage` | `query` | Get page content by ID or title search |
| `createPage` | `title`, `bodyHtml`, `sectionId?` | Create a page |
| `searchPages` | `query` | Search pages by title |

## CLI

A single CLI replaces the standalone scripts. Run via `npm run cli` or directly with `npx tsx src/cli.ts`.

```bash
npm run auth                         # device-code sign-in
npm run verify                       # verify token against Graph API

npm run cli -- notebooks             # list notebooks
npm run cli -- sections              # list all sections
npm run cli -- sections --notebook <id>
npm run cli -- pages                 # list all pages
npm run cli -- pages --section <id>
npm run cli -- get "page title"      # get page content as text
npm run cli -- create --title "My Page" --body "<p>Hello</p>"
npm run cli -- search "meeting"      # search by title
```

## Scripts

| Script | Description |
|--------|-------------|
| `npm run build` | Compile TypeScript to `dist/` |
| `npm start` | Run the compiled MCP server |
| `npm run dev` | Run the MCP server with `tsx` (no build needed) |
| `npm run auth` | Device-code authentication |
| `npm run verify` | Verify stored token |
| `npm run cli` | Unified CLI (see above) |
| `npm run typecheck` | Type-check without emitting |
| `npm test` | Run tests with Vitest |
| `npm run test:watch` | Run tests in watch mode |

## Project Structure

```
src/
  config.ts         — Client ID, tenant, scopes, paths
  logger.ts         — stderr-only logging (MCP uses stdout for JSON-RPC)
  token-store.ts    — Load/save/normalize access tokens
  auth.ts           — Device-code authentication flow
  graph-client.ts   — Microsoft Graph SDK client factory
  html.ts           — HTML-to-text conversion (dependency-free)
  onenote.ts        — Typed OneNote client (notebooks, sections, pages)
  mcp-server.ts     — MCP server entry point with Zod-typed tools
  cli.ts            — Unified CLI
  index.ts          — Barrel export
tests/
  html.test.ts
  token-store.test.ts
  onenote.test.ts
```

## Authentication Details

Authentication uses Microsoft's [device-code flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-device-code) with the Graph Explorer public client ID. No Azure app registration is needed.

The server requests these delegated scopes:

- `Notes.Read` — read notebooks, sections, and pages
- `Notes.ReadWrite` — create and modify pages
- `User.Read` — verify the signed-in identity

These are **delegated, non-`.All`** scopes, which is critical: personal Microsoft accounts cannot consent to `.All` (application-level) scopes, and requesting them produces a token that Graph rejects with HTTP 401.

Tokens are cached in `.access-token.txt` with owner-only file permissions (`600`). Tokens expire after roughly one hour and need to be refreshed by re-running `npm run auth`.

### Environment Variable Overrides

| Variable | Default | Description |
|----------|---------|-------------|
| `GRAPH_CLIENT_ID` | Graph Explorer client ID | Your own Azure app registration |
| `GRAPH_TENANT` | `common` | `consumers`, `organizations`, or a tenant ID |
| `ONENOTE_TOKEN_PATH` | `.access-token.txt` | Custom token file location |
| `GRAPH_ACCESS_TOKEN` | — | Provide a token directly (skips file) |

## Troubleshooting

**401 / error code 40001** — You're likely requesting `.All` scopes, or your token has expired. Re-run `npm run auth`.

**"Token is a compact (non-JWT) token"** — This is normal for personal Microsoft accounts. Personal accounts return opaque (non-JWT) access tokens that are perfectly valid for Graph API calls.

**Server output corrupts the MCP stream** — All logging goes to stderr. If you add `console.log()` calls, the JSON-RPC protocol over stdout will break. Use `log()` from `src/logger.ts` instead.

## Security

- Tokens are stored locally with `chmod 600` permissions
- No credentials are sent to any third party
- The public client ID has no client secret; authentication relies entirely on the user completing the device-code flow
- `.access-token.txt` is in `.gitignore`

## Credits

Built on [azure-onenote-mcp-server](https://github.com/ZubeidHendricks/azure-onenote-mcp-server) by Zubeid Hendricks.

## License

MIT — see [LICENSE](LICENSE) for details.
