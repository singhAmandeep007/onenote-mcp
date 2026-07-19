#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { log } from './logger.js';
import { loadToken, saveToken, normalizeAccessToken } from './token-store.js';
import { authenticateWithDeviceCode } from './auth.js';
import { OneNoteClient } from './onenote.js';

// ── Server setup ────────────────────────────────────────────────────────────

const server = new McpServer(
  {
    name: 'onenote',
    version: '2.0.0',
    description: 'MCP server for Microsoft OneNote — read, create, and search pages via Microsoft Graph.',
  },
  { capabilities: { tools: { listChanged: true } } },
);

// ── Helpers ─────────────────────────────────────────────────────────────────

function client(): OneNoteClient {
  return OneNoteClient.fromStoredToken();
}

function ok(text: string) {
  return { content: [{ type: 'text' as const, text }] };
}

function err(message: string) {
  return { content: [{ type: 'text' as const, text: `Error: ${message}` }], isError: true as const };
}

// ── Tools ───────────────────────────────────────────────────────────────────

server.tool(
  'authenticate',
  'Start the Microsoft device-code authentication flow. Follow the on-screen instructions to sign in.',
  async () => {
    try {
      const existing = loadToken();
      if (existing) {
        return ok('Already authenticated with a stored access token. Use saveAccessToken to replace it, or delete .access-token.txt to re-authenticate.');
      }
      await authenticateWithDeviceCode();
      return ok('Authentication successful. Token saved.');
    } catch (e) {
      return err(`Authentication failed: ${(e as Error).message}`);
    }
  },
);

server.tool(
  'saveAccessToken',
  'Manually save a Microsoft Graph access token for later use.',
  { token: z.string().describe('The access token to save') },
  // @ts-expect-error — TS2589: MCP SDK overload + Zod chain triggers deep type instantiation
  async ({ token }: { token: string }) => {
    try {
      const normalized = normalizeAccessToken(token);
      if (!normalized) return err('Provided token is empty after normalization.');
      saveToken(normalized);
      return ok('Access token saved successfully.');
    } catch (e) {
      return err(`Failed to save token: ${(e as Error).message}`);
    }
  },
);

server.tool(
  'listNotebooks',
  'List all OneNote notebooks for the signed-in user.',
  async () => {
    try {
      const notebooks = await client().listNotebooks();
      return ok(JSON.stringify(notebooks, null, 2));
    } catch (e) {
      return err((e as Error).message);
    }
  },
);

server.tool(
  'getNotebook',
  'Get details of a specific notebook by its ID.',
  { notebookId: z.string().min(1).describe('The notebook ID') },
  async ({ notebookId }) => {
    try {
      const notebook = await client().getNotebook(notebookId);
      return ok(JSON.stringify(notebook, null, 2));
    } catch (e) {
      return err((e as Error).message);
    }
  },
);

server.tool(
  'listSections',
  'List sections, optionally scoped to a notebook.',
  { notebookId: z.string().optional().describe('Notebook ID (omit to list all sections)') },
  async ({ notebookId }) => {
    try {
      const sections = await client().listSections(notebookId);
      return ok(JSON.stringify(sections, null, 2));
    } catch (e) {
      return err((e as Error).message);
    }
  },
);

server.tool(
  'listPages',
  'List pages, optionally scoped to a section.',
  { sectionId: z.string().optional().describe('Section ID (omit to list all pages)') },
  async ({ sectionId }) => {
    try {
      const pages = await client().listPages(sectionId);
      return ok(JSON.stringify(pages, null, 2));
    } catch (e) {
      return err((e as Error).message);
    }
  },
);

server.tool(
  'getPage',
  'Get page content by ID or title search. Returns both HTML and plain text.',
  { query: z.string().min(1).describe('Page ID or title substring to search for') },
  async ({ query }) => {
    try {
      const c = client();
      const page = await c.findPage(query);
      if (!page || !page.id) return err(`No page found matching "${query}".`);

      const content = await c.getPageContent(page.id);
      return ok(JSON.stringify({
        id: content.id,
        title: content.title,
        text: content.text,
        htmlLength: content.html.length,
      }, null, 2));
    } catch (e) {
      return err((e as Error).message);
    }
  },
);

server.tool(
  'createPage',
  'Create a new page in a OneNote section.',
  {
    title: z.string().min(1).describe('Page title'),
    bodyHtml: z.string().describe('HTML body content for the page'),
    sectionId: z.string().optional().describe('Section ID (omit to use first available section)'),
  },
  async ({ title, bodyHtml, sectionId }) => {
    try {
      const page = await client().createPage(title, bodyHtml, sectionId);
      return ok(JSON.stringify(page, null, 2));
    } catch (e) {
      return err((e as Error).message);
    }
  },
);

server.tool(
  'searchPages',
  'Search for pages by title across all notebooks.',
  { query: z.string().describe('Search term (case-insensitive title substring). Empty string returns all pages.') },
  async ({ query }) => {
    try {
      const pages = await client().searchPages(query);
      return ok(JSON.stringify(pages, null, 2));
    } catch (e) {
      return err((e as Error).message);
    }
  },
);

// ── Main ────────────────────────────────────────────────────────────────────

async function main(): Promise<void> {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  log('Server started. Tools: authenticate, saveAccessToken, listNotebooks, getNotebook, listSections, listPages, getPage, createPage, searchPages');
}

main().catch((e) => {
  log('Fatal error:', e);
  process.exit(1);
});
