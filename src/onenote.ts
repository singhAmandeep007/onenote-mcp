import type { Client } from "@microsoft/microsoft-graph-client";
import type { Notebook, OnenoteSection, OnenotePage } from "@microsoft/microsoft-graph-types";
import { createGraphClient, type TokenProvider } from "./graph-client.js";
import { getAccessToken } from "./auth.js";
import { htmlToText } from "./html.js";
import { log } from "./logger.js";
import { GRAPH_BASE_URL } from "./config.js";

// ── Types ───────────────────────────────────────────────────────────────────

export interface PageContent {
  id: string;
  title: string;
  html: string;
  text: string;
}

export interface UserProfile {
  displayName: string;
  mail: string | null;
}

// ── Client ──────────────────────────────────────────────────────────────────

/**
 * Typed OneNote client that wraps the Microsoft Graph SDK. All OneNote
 * operations go through this class so there is a single place to manage
 * authentication, create the Graph client, and apply error handling.
 *
 * Uses an async token provider instead of a static token string. This means
 * expired access tokens are silently refreshed behind the scenes — no
 * manual re-authentication needed.
 */
export class OneNoteClient {
  private client: Client;
  private tokenProvider: TokenProvider;

  constructor(tokenProvider: TokenProvider) {
    this.tokenProvider = tokenProvider;
    this.client = createGraphClient(tokenProvider);
  }

  /**
   * Build a client that uses the token cache for automatic renewal.
   * Throws if no cached account exists (user hasn't authenticated yet).
   */
  static create(): OneNoteClient {
    return new OneNoteClient(getAccessToken);
  }

  // ── User ────────────────────────────────────────────────────────────────

  async getMe(): Promise<UserProfile> {
    const me = (await this.client.api("/me").select("displayName,mail").get()) as {
      displayName: string;
      mail?: string | null;
    };
    return { displayName: me.displayName, mail: me.mail ?? null };
  }

  // ── Notebooks ───────────────────────────────────────────────────────────

  async listNotebooks(): Promise<Notebook[]> {
    const res = await this.client.api("/me/onenote/notebooks").get();
    return (res.value ?? []) as Notebook[];
  }

  async getNotebook(notebookId: string): Promise<Notebook> {
    return (await this.client.api(`/me/onenote/notebooks/${notebookId}`).get()) as Notebook;
  }

  // ── Sections ────────────────────────────────────────────────────────────

  /** List all sections across all notebooks, or within a specific notebook. */
  async listSections(notebookId?: string): Promise<OnenoteSection[]> {
    const endpoint = notebookId ? `/me/onenote/notebooks/${notebookId}/sections` : "/me/onenote/sections";
    const res = await this.client.api(endpoint).get();
    return (res.value ?? []) as OnenoteSection[];
  }

  // ── Pages ───────────────────────────────────────────────────────────────

  /** List pages — optionally scoped to a section. */
  async listPages(sectionId?: string): Promise<OnenotePage[]> {
    const endpoint = sectionId ? `/me/onenote/sections/${sectionId}/pages` : "/me/onenote/pages";
    const res = await this.client.api(endpoint).get();
    return (res.value ?? []) as OnenotePage[];
  }

  /** Get a page's HTML content and convert it to text. */
  async getPageContent(pageId: string): Promise<PageContent> {
    // The Graph SDK's .get() on the /content endpoint returns a ReadableStream
    // which is awkward to consume. Use native fetch instead for the raw HTML.
    // Get a fresh token from the provider for the raw fetch call.
    const token = await this.tokenProvider();
    const url = `${GRAPH_BASE_URL}/me/onenote/pages/${pageId}/content`;
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!response.ok) {
      throw new Error(`Failed to fetch page content: ${response.status} ${response.statusText}`);
    }

    const html = await response.text();

    // We also need the title, which isn't in the content response.
    // Get it from the page metadata.
    let title = "";
    try {
      const meta = (await this.client.api(`/me/onenote/pages/${pageId}`).select("title").get()) as { title?: string };
      title = meta.title ?? "";
    } catch {
      log("Could not fetch page title for", pageId);
    }

    return { id: pageId, title, html, text: htmlToText(html) };
  }

  /**
   * Find a page by ID or title (case-insensitive substring match).
   * Returns null if nothing matches.
   */
  async findPage(query: string): Promise<OnenotePage | null> {
    const pages = await this.listPages();

    // Exact ID match
    const byId = pages.find((p) => p.id === query);
    if (byId) return byId;

    // Title substring match (case-insensitive)
    const lower = query.toLowerCase();
    const byTitle = pages.find((p) => p.title && p.title.toLowerCase().includes(lower));
    if (byTitle) return byTitle;

    // Partial ID match
    const byPartialId = pages.find((p) => p.id?.includes(query) || query.includes(p.id ?? ""));
    return byPartialId ?? null;
  }

  /** Search pages by title (case-insensitive substring). */
  async searchPages(query: string): Promise<OnenotePage[]> {
    const pages = await this.listPages();
    if (!query) return pages;
    const lower = query.toLowerCase();
    return pages.filter((p) => p.title && p.title.toLowerCase().includes(lower));
  }

  /**
   * Create a page in the given section (or the first available section).
   * @param title  Page title
   * @param bodyHtml  HTML body content (without the full document wrapper)
   * @param sectionId  Optional section ID; defaults to first section found
   */
  async createPage(title: string, bodyHtml: string, sectionId?: string): Promise<OnenotePage> {
    const targetSection = sectionId ?? (await this.firstSectionId());

    const html = `<!DOCTYPE html>
<html>
  <head><title>${escapeForTitle(title)}</title></head>
  <body>${bodyHtml}</body>
</html>`;

    const page = (await this.client
      .api(`/me/onenote/sections/${targetSection}/pages`)
      .header("Content-Type", "application/xhtml+xml")
      .post(html)) as OnenotePage;

    return page;
  }

  // ── Helpers ─────────────────────────────────────────────────────────────

  /** Verify that the token is valid by calling /me. */
  async verifyToken(): Promise<{ user: UserProfile; notebookCount: number }> {
    const user = await this.getMe();
    const notebooks = await this.listNotebooks();
    return { user, notebookCount: notebooks.length };
  }

  private async firstSectionId(): Promise<string> {
    const sections = await this.listSections();
    if (sections.length === 0) {
      throw new Error("No sections found. Create a section in OneNote first.");
    }
    return sections[0]!.id!;
  }
}

// ── Utilities ─────────────────────────────────────────────────────────────

/** Minimal escaping for content inside a <title> tag. */
function escapeForTitle(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
