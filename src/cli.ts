#!/usr/bin/env node

/**
 * Unified CLI for the OneNote MCP project.
 *
 * Replaces ~10 standalone scripts (authenticate.js, verify-token.js,
 * simple-onenote.js, list-sections.js, list-pages.js, create-page.js,
 * get-page.js, get-page-content.js, etc.) with subcommands:
 *
 *   onenote-cli auth         Run the device-code sign-in flow
 *   onenote-cli verify       Verify the stored token against Graph
 *   onenote-cli notebooks    List notebooks
 *   onenote-cli sections     List sections [--notebook <id>]
 *   onenote-cli pages        List pages [--section <id>]
 *   onenote-cli get <query>  Get page content by ID or title
 *   onenote-cli create       Create a page [--title <t>] [--body <html>] [--section <id>]
 *   onenote-cli search <q>   Search pages by title
 */

import { authenticateWithDeviceCode } from './auth.js';
import { OneNoteClient } from './onenote.js';
import { isLikelyJwt, loadToken } from './token-store.js';

// ── Arg parsing ─────────────────────────────────────────────────────────────

const args = process.argv.slice(2);
const command = args[0]?.toLowerCase() ?? 'help';

function flag(name: string): string | undefined {
  const idx = args.indexOf(`--${name}`);
  return idx >= 0 ? args[idx + 1] : undefined;
}

function positional(index: number): string | undefined {
  // Return args after the command, skipping flags
  let pos = 0;
  for (let i = 1; i < args.length; i++) {
    if (args[i]!.startsWith('--')) {
      i++; // skip flag value
      continue;
    }
    if (pos === index) return args[i];
    pos++;
  }
  return undefined;
}

// ── Commands ────────────────────────────────────────────────────────────────

async function run(): Promise<void> {
  switch (command) {
    case 'auth':
    case 'authenticate': {
      console.log('Starting device-code authentication...\n');
      const result = await authenticateWithDeviceCode((msg) => console.log(msg));
      console.log(`\nAuthentication successful.`);
      console.log(`Token type: ${result.isJwt ? 'JWT (work/school account)' : 'compact (personal account)'}`);
      break;
    }

    case 'verify': {
      const client = OneNoteClient.fromStoredToken();
      const { user, notebookCount } = await client.verifyToken();
      console.log(`Signed in as: ${user.displayName} (${user.mail ?? 'no email'})`);
      console.log(`Notebooks: ${notebookCount}`);

      const token = loadToken();
      if (token) {
        console.log(`Token format: ${isLikelyJwt(token) ? 'JWT (work/school)' : 'compact (personal account)'}`);
      }
      break;
    }

    case 'notebooks': {
      const notebooks = await OneNoteClient.fromStoredToken().listNotebooks();
      if (notebooks.length === 0) {
        console.log('No notebooks found.');
      } else {
        for (const nb of notebooks) {
          console.log(`  ${nb.displayName}  [${nb.id}]`);
        }
      }
      break;
    }

    case 'sections': {
      const notebookId = flag('notebook');
      const sections = await OneNoteClient.fromStoredToken().listSections(notebookId);
      if (sections.length === 0) {
        console.log('No sections found.');
      } else {
        for (const s of sections) {
          console.log(`  ${s.displayName}  [${s.id}]`);
        }
      }
      break;
    }

    case 'pages': {
      const sectionId = flag('section');
      const pages = await OneNoteClient.fromStoredToken().listPages(sectionId);
      if (pages.length === 0) {
        console.log('No pages found.');
      } else {
        for (const p of pages) {
          console.log(`  ${p.title ?? '(untitled)'}  [${p.id}]`);
        }
      }
      break;
    }

    case 'get': {
      const query = positional(0);
      if (!query) {
        console.error('Usage: onenote-cli get <page-id-or-title>');
        process.exit(1);
      }
      const client = OneNoteClient.fromStoredToken();
      const page = await client.findPage(query);
      if (!page || !page.id) {
        console.error(`No page found matching "${query}".`);
        process.exit(1);
      }
      const content = await client.getPageContent(page.id);
      console.log(`# ${content.title}\n`);
      console.log(content.text);
      break;
    }

    case 'create': {
      const title = flag('title') ?? `New Page — ${new Date().toLocaleDateString()}`;
      const body = flag('body') ?? '<p>Created via onenote-cli.</p>';
      const sectionId = flag('section');
      const page = await OneNoteClient.fromStoredToken().createPage(title, body, sectionId);
      console.log(`Created page: ${page.title ?? title}  [${page.id}]`);
      break;
    }

    case 'search': {
      const query = positional(0) ?? '';
      const pages = await OneNoteClient.fromStoredToken().searchPages(query);
      if (pages.length === 0) {
        console.log('No pages found.');
      } else {
        for (const p of pages) {
          console.log(`  ${p.title ?? '(untitled)'}  [${p.id}]`);
        }
      }
      break;
    }

    case 'help':
    default:
      console.log(`
onenote-cli — OneNote MCP command-line interface

Commands:
  auth                     Run device-code sign-in
  verify                   Verify stored token against Graph
  notebooks                List notebooks
  sections [--notebook id] List sections
  pages    [--section id]  List pages
  get <id-or-title>        Get page content as plain text
  create   [--title t] [--body html] [--section id]
                           Create a page
  search <query>           Search pages by title
  help                     Show this message
`.trim());
      break;
  }
}

run().catch((e: unknown) => {
  console.error(`Error: ${(e as Error).message}`);
  process.exit(1);
});
