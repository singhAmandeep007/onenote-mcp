#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { log } from "./logger.js";
import { authenticateWithDeviceCode, acquireTokenSilent, clearCache } from "./auth.js";
import { OneNoteClient } from "./onenote.js";

// ── Server setup ────────────────────────────────────────────────────────────

const server = new McpServer(
  {
    name: "onenote",
    version: "2.0.0",
    description: "MCP server for Microsoft OneNote — read, create, and search pages via Microsoft Graph.",
  },
  { capabilities: { tools: { listChanged: true } } }
);

// ── Helpers ─────────────────────────────────────────────────────────────────

function client(): OneNoteClient {
  return OneNoteClient.create();
}

type ToolResult = { content: { type: "text"; text: string }[]; isError?: true };

function ok(text: string): ToolResult {
  return { content: [{ type: "text", text }] };
}

function err(message: string): ToolResult {
  return {
    content: [{ type: "text", text: `Error: ${message}` }],
    isError: true,
  };
}

/**
 * Typed wrapper around server.tool() that avoids TS2589.
 *
 * The MCP SDK's tool() has deeply nested conditional types for overload
 * resolution with Zod schemas. TypeScript's type checker intermittently
 * hits its recursion limit (TS2589) depending on the number of tools and
 * schema complexity. This wrapper casts through `any` to bypass the deep
 * inference while keeping Zod runtime validation intact.
 */
function defineTool(
  name: string,
  description: string,
  schema: Record<string, z.ZodTypeAny>,
  handler: (params: Record<string, unknown>) => Promise<ToolResult>
): void {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-call
  (server as any).tool(name, description, schema, handler);
}

// ── Tools ───────────────────────────────────────────────────────────────────

server.tool(
  "authenticate",
  "Start the Microsoft device-code authentication flow. Only needed once — tokens renew automatically after initial sign-in.",
  async () => {
    try {
      // Try silent renewal first — if it works, we're good
      const token = await acquireTokenSilent();
      if (token) {
        return ok("Already authenticated. Tokens renew automatically — no action needed.");
      }

      // Silent failed (no account or refresh token expired) — clear stale cache and re-auth
      log("Silent renewal failed; clearing stale cache and starting fresh auth.");
      await clearCache();
      await authenticateWithDeviceCode();
      return ok(
        "Authentication successful. Tokens will renew automatically — you won't need to sign in again unless you revoke access."
      );
    } catch (e) {
      return err(`Authentication failed: ${(e as Error).message}`);
    }
  }
);

server.tool("listNotebooks", "List all OneNote notebooks for the signed-in user.", async () => {
  try {
    const notebooks = await client().listNotebooks();
    return ok(JSON.stringify(notebooks, null, 2));
  } catch (e) {
    return err((e as Error).message);
  }
});

defineTool(
  "getNotebook",
  "Get details of a specific notebook by its ID.",
  { notebookId: z.string().min(1).describe("The notebook ID") },
  async ({ notebookId }) => {
    try {
      const notebook = await client().getNotebook(notebookId as string);
      return ok(JSON.stringify(notebook, null, 2));
    } catch (e) {
      return err((e as Error).message);
    }
  }
);

defineTool(
  "listSections",
  "List sections, optionally scoped to a notebook.",
  {
    notebookId: z.string().optional().describe("Notebook ID (omit to list all sections)"),
  },
  async ({ notebookId }) => {
    try {
      const sections = await client().listSections(notebookId as string | undefined);
      return ok(JSON.stringify(sections, null, 2));
    } catch (e) {
      return err((e as Error).message);
    }
  }
);

defineTool(
  "listPages",
  "List pages, optionally scoped to a section.",
  {
    sectionId: z.string().optional().describe("Section ID (omit to list all pages)"),
  },
  async ({ sectionId }) => {
    try {
      const pages = await client().listPages(sectionId as string | undefined);
      return ok(JSON.stringify(pages, null, 2));
    } catch (e) {
      return err((e as Error).message);
    }
  }
);

defineTool(
  "getPage",
  "Get page content by ID or title search. Returns both HTML and plain text.",
  {
    query: z.string().min(1).describe("Page ID or title substring to search for"),
  },
  async ({ query }) => {
    try {
      const c = client();
      const page = await c.findPage(query as string);
      if (!page || !page.id) return err(`No page found matching "${query}".`);

      const content = await c.getPageContent(page.id);
      return ok(
        JSON.stringify(
          {
            id: content.id,
            title: content.title,
            text: content.text,
            htmlLength: content.html.length,
          },
          null,
          2
        )
      );
    } catch (e) {
      return err((e as Error).message);
    }
  }
);

defineTool(
  "createPage",
  "Create a new page in a OneNote section.",
  {
    title: z.string().min(1).describe("Page title"),
    bodyHtml: z.string().describe("HTML body content for the page"),
    sectionId: z.string().optional().describe("Section ID (omit to use first available section)"),
  },
  async ({ title, bodyHtml, sectionId }) => {
    try {
      const page = await client().createPage(title as string, bodyHtml as string, sectionId as string | undefined);
      return ok(JSON.stringify(page, null, 2));
    } catch (e) {
      return err((e as Error).message);
    }
  }
);

defineTool(
  "searchPages",
  "Search for pages by title across all notebooks.",
  {
    query: z.string().describe("Search term (case-insensitive title substring). Empty string returns all pages."),
  },
  async ({ query }) => {
    try {
      const pages = await client().searchPages(query as string);
      return ok(JSON.stringify(pages, null, 2));
    } catch (e) {
      return err((e as Error).message);
    }
  }
);

// ── Main ────────────────────────────────────────────────────────────────────

async function main(): Promise<void> {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  log("Server started. Tokens renew automatically via cached refresh token.");
}

main().catch((e) => {
  log("Fatal error:", e);
  process.exit(1);
});
