import { describe, it, expect, vi, beforeEach } from 'vitest';

// Mock the graph-client module so we never hit a real Graph API.
vi.mock('../src/graph-client.js', () => ({
  createGraphClient: vi.fn(),
}));

// Mock token-store so OneNoteClient.fromStoredToken works without a file.
vi.mock('../src/token-store.js', () => ({
  loadToken: vi.fn(() => 'fake-token'),
  saveToken: vi.fn(),
  normalizeAccessToken: vi.fn((v: string) => v?.trim() || null),
  isLikelyJwt: vi.fn(() => false),
}));

import { createGraphClient } from '../src/graph-client.js';
import { OneNoteClient } from '../src/onenote.js';

// Helper: build a mock Graph client whose .api(path).get()/post() we control.
function mockGraphClient(responses: Record<string, unknown>) {
  const apiMock = vi.fn((path: string) => {
    const data = responses[path] ?? { value: [] };
    return {
      get: vi.fn(async () => data),
      select: vi.fn().mockReturnThis(),
      header: vi.fn().mockReturnThis(),
      post: vi.fn(async () => data),
    };
  });
  (createGraphClient as ReturnType<typeof vi.fn>).mockReturnValue({ api: apiMock });
  return apiMock;
}

describe('OneNoteClient', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe('listNotebooks', () => {
    it('returns notebooks from Graph', async () => {
      const notebooks = [{ id: 'nb1', displayName: 'Work' }];
      mockGraphClient({ '/me/onenote/notebooks': { value: notebooks } });

      const client = new OneNoteClient('fake-token');
      const result = await client.listNotebooks();
      expect(result).toEqual(notebooks);
    });
  });

  describe('listSections', () => {
    it('lists all sections when no notebookId given', async () => {
      const sections = [{ id: 's1', displayName: 'General' }];
      mockGraphClient({ '/me/onenote/sections': { value: sections } });

      const client = new OneNoteClient('fake-token');
      const result = await client.listSections();
      expect(result).toEqual(sections);
    });

    it('scopes to a notebook when notebookId given', async () => {
      const sections = [{ id: 's2', displayName: 'Notes' }];
      const apiMock = mockGraphClient({
        '/me/onenote/notebooks/nb1/sections': { value: sections },
      });

      const client = new OneNoteClient('fake-token');
      const result = await client.listSections('nb1');
      expect(result).toEqual(sections);
      expect(apiMock).toHaveBeenCalledWith('/me/onenote/notebooks/nb1/sections');
    });
  });

  describe('findPage', () => {
    it('matches by exact ID', async () => {
      const pages = [
        { id: 'page-123', title: 'A' },
        { id: 'page-456', title: 'B' },
      ];
      mockGraphClient({ '/me/onenote/pages': { value: pages } });

      const client = new OneNoteClient('fake-token');
      const found = await client.findPage('page-456');
      expect(found?.title).toBe('B');
    });

    it('falls back to title substring match', async () => {
      const pages = [{ id: 'p1', title: 'Meeting Notes 2024' }];
      mockGraphClient({ '/me/onenote/pages': { value: pages } });

      const client = new OneNoteClient('fake-token');
      const found = await client.findPage('meeting notes');
      expect(found?.id).toBe('p1');
    });

    it('returns null when nothing matches', async () => {
      mockGraphClient({ '/me/onenote/pages': { value: [] } });

      const client = new OneNoteClient('fake-token');
      expect(await client.findPage('nonexistent')).toBeNull();
    });
  });

  describe('searchPages', () => {
    it('filters pages by title', async () => {
      const pages = [
        { id: 'p1', title: 'Alpha' },
        { id: 'p2', title: 'Beta' },
        { id: 'p3', title: 'Alpha Beta' },
      ];
      mockGraphClient({ '/me/onenote/pages': { value: pages } });

      const client = new OneNoteClient('fake-token');
      const results = await client.searchPages('alpha');
      expect(results).toHaveLength(2);
      expect(results.map((p) => p.id)).toEqual(['p1', 'p3']);
    });

    it('returns all pages when query is empty', async () => {
      const pages = [{ id: 'p1', title: 'X' }];
      mockGraphClient({ '/me/onenote/pages': { value: pages } });

      const client = new OneNoteClient('fake-token');
      expect(await client.searchPages('')).toHaveLength(1);
    });
  });

  describe('createPage', () => {
    it('posts XHTML to the correct section endpoint', async () => {
      const created = { id: 'new-page', title: 'Test' };
      const apiMock = mockGraphClient({
        '/me/onenote/sections': { value: [{ id: 'sec1' }] },
        '/me/onenote/sections/sec1/pages': created,
      });

      const client = new OneNoteClient('fake-token');
      const page = await client.createPage('Test', '<p>body</p>');
      expect(page).toEqual(created);
      expect(apiMock).toHaveBeenCalledWith('/me/onenote/sections/sec1/pages');
    });

    it('uses provided sectionId when given', async () => {
      const created = { id: 'new-page', title: 'T' };
      const apiMock = mockGraphClient({
        '/me/onenote/sections/custom-sec/pages': created,
      });

      const client = new OneNoteClient('fake-token');
      await client.createPage('T', '<p>hi</p>', 'custom-sec');
      expect(apiMock).toHaveBeenCalledWith('/me/onenote/sections/custom-sec/pages');
    });
  });

  describe('fromStoredToken', () => {
    it('creates a client from the stored token', () => {
      mockGraphClient({});
      const client = OneNoteClient.fromStoredToken();
      expect(client).toBeInstanceOf(OneNoteClient);
    });
  });
});
