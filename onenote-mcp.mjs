#!/usr/bin/env node
/**
 * OneNote MCP server.
 *
 * Built on @modelcontextprotocol/server v2, which implements the 2026-07-28 MCP
 * specification. `serveStdio` handles the era decision per connection, so this
 * one server answers both modern stateless clients and 2025-era clients that
 * still open with `initialize` -- no branching needed here.
 *
 * Design notes that matter for correctness:
 *
 *  - Every tool declares a real Zod input schema. The previous version registered
 *    tools with no schema and then read `params.random_string`, an artifact of a
 *    Cursor placeholder argument. Clients had nothing to fill in, so
 *    `listSections` ignored the notebook, `listPages` always used the first
 *    section, and `createPage` wrote a fixed placeholder page.
 *
 *  - `authenticate` never blocks on the human. It returns the device code or
 *    browser instruction immediately and lets sign-in finish in the background,
 *    which is what makes it work in Claude Desktop, Cursor, Trae, and anything
 *    else that enforces a tool timeout (danosb/onenote-mcp#6, #1).
 *
 *  - Nothing is ever written to stdout except JSON-RPC. Diagnostics go to stderr.
 */

import { McpServer } from '@modelcontextprotocol/server';
import { serveStdio } from '@modelcontextprotocol/server/stdio';
import { z } from 'zod';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

import { loadConfig, ConfigError } from './src/config.mjs';
import { OneNoteAuth, AuthSession, inspectCache } from './src/auth.mjs';
import { OneNoteClient } from './src/onenote.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const { version } = JSON.parse(readFileSync(path.join(__dirname, 'package.json'), 'utf8'));

/** stdout belongs to the protocol; diagnostics go to stderr. */
const log = (message) => process.stderr.write(`${message}\n`);

// .env is loaded only if present, and only for keys not already set, so a host
// like Claude Desktop passing env through its config always wins.
await loadDotEnvIfPresent();

let config;
try {
  config = loadConfig();
} catch (error) {
  if (error instanceof ConfigError) {
    log(`\nOneNote MCP server cannot start.\n\n${error.message}\n`);
    process.exit(1);
  }
  throw error;
}

const auth = new OneNoteAuth(config, log);
const session = new AuthSession(auth);
const onenote = new OneNoteClient(auth, config, log);

// ---------------------------------------------------------------- tool result

/**
 * Text plus machine-readable payload.
 *
 * TOKEN BUDGET: `content` and `structuredContent` are BOTH likely to land in the
 * model's context, so anything present in both is paid for twice. Keep them
 * complementary: prose in `content`, identifiers in `structuredContent`, and
 * never the same large blob in both. Returning a 6,000-character page in each
 * doubled its cost for no benefit.
 */
function ok(summary, data) {
  return {
    content: [{ type: 'text', text: summary }],
    structuredContent: data
  };
}

/** `2026-07-30` rather than `2026-07-30T17:33:12Z` — same use, a third the size. */
function shortDate(iso) {
  return typeof iso === 'string' ? iso.slice(0, 10) : null;
}

/**
 * Compact page record for list-shaped results.
 *
 * Drops createdDateTime (rarely relevant) and webUrl (~120 characters each, and
 * reconstructible on demand). Across 50 pages those two fields alone cost more
 * than everything else combined.
 */
function leanPage(page) {
  return {
    id: page.id,
    title: page.title || '(untitled)',
    section: page.sectionName ?? undefined,
    modified: shortDate(page.lastModifiedDateTime)
  };
}

function fail(message) {
  return {
    content: [{ type: 'text', text: message }],
    isError: true
  };
}

/**
 * Wrap a handler so unauthenticated and Graph errors become actionable text
 * rather than an opaque protocol-level exception.
 */
function handler(fn) {
  return async (args, ctx) => {
    try {
      return await fn(args ?? {}, ctx);
    } catch (error) {
      if (error?.name === 'AuthError' || error?.status === 401) {
        return fail(
          `${error.message}\n\n` +
            'Call the `authenticate` tool to sign in, or run `npm run auth` in a terminal on ' +
            'the machine hosting this server.'
        );
      }
      log(`Tool error: ${error?.stack || error?.message || error}`);
      return fail(error?.message ? `${error.message}` : 'The request failed for an unknown reason.');
    }
  };
}

// --------------------------------------------------------------------- schemas

/**
 * Sections are addressed by name wherever possible. Requiring a GUID forced a
 * discovery round-trip (and an approval click) before every real action.
 */
const SectionRef = z
  .string()
  .min(1)
  .describe(
    'Section name ("Ideas"), path ("Learn / Cooking"), or ID. Names are matched ' +
      'case-insensitively; if several match you get the list back to choose from.'
  );

const PageRef = z
  .string()
  .min(1)
  .describe(
    'Page TITLE (e.g. "AI Search") or ID. Titles are matched case-insensitively; ' +
      'if several pages share one you get the list back to choose from. You do NOT ' +
      'need to look up an ID first.'
  );

const Limit = z
  .number()
  .int()
  .min(1)
  .max(200)
  .describe('Maximum number of items to return. Results are paginated internally.');

/**
 * Default cap on page text returned to the model.
 *
 * Long enough for essentially every real note, short enough that one runaway
 * page cannot consume the whole context. The tool reports truncation and how to
 * ask for more, so nothing is silently lost.
 */
const DEFAULT_PAGE_CHARS = 12_000;

/**
 * Default page-list size.
 *
 * Was 100, which routinely returned far more than any question needed. Asking
 * for more is one cheap call; sending it unprompted costs every time.
 */
const DEFAULT_PAGE_LIMIT = 25;

// ---------------------------------------------------------------------- server

function createServer() {
  const server = new McpServer(
    { name: 'onenote', version, title: 'Microsoft OneNote' },
    {
      capabilities: { tools: { listChanged: true } },
      instructions:
        'Read and write Microsoft OneNote for the signed-in user.\n\n' +
        'EFFICIENCY -- do not crawl the hierarchy:\n' +
        '- To see what exists, call `getOutline` ONCE. It returns every notebook, section group, ' +
        'and section with IDs and readable paths. Do not chain listNotebooks -> listSections.\n' +
        '- Tools accept section NAMES and page TITLES directly ("Ideas", "Learn / Cooking", ' +
        '"AI Search"). You never need to look up an ID first — pass the name straight through ' +
        'from what searchPages or getOutline showed you.\n' +
        '- To find content, call `searchPages` directly, then pass a matching title straight to ' +
        '`getPage`. Add `section` to narrow either one.\n' +
        '- Reserve `listNotebooks` / `listSections` for when the user explicitly asks to browse.\n\n' +
        'Note that sections often live inside section groups; `getOutline` shows that nesting, ' +
        'and every tool handles it transparently.'
    }
  );

  // ------------------------------------------------------------------- auth

  server.registerTool(
    'authStatus',
    {
      title: 'Authentication status',
      description:
        'Report whether the user is signed in to Microsoft, which account is cached, and where the ' +
        'token cache lives. Safe to call at any time; never triggers a sign-in prompt.',
      inputSchema: z.object({}),
      annotations: { readOnlyHint: true, openWorldHint: true }
    },
    handler(async () => {
      const account = await auth.getAccount();
      const cache = inspectCache(config.cachePath);
      const live = session.snapshot();

      let tokenUsable = false;
      if (account) {
        try {
          await auth.getAccessToken({ interactive: false });
          tokenUsable = true;
        } catch {
          tokenUsable = false;
        }
      }

      const data = {
        signedIn: tokenUsable,
        account: account?.username ?? null,
        cachedAccountPresent: Boolean(account),
        method: config.allowDeviceCode ? 'device_code' : 'pkce',
        scopes: config.scopes,
        tenant: config.tenant,
        clientId: config.clientId,
        writeEnabled: config.allowWrite,
        tokenCache: cache,
        pendingSignIn: live.state === 'pending' ? live : null,
        lastSignInError: live.state === 'failed' ? live.error : null
      };

      let summary;
      if (tokenUsable) {
        summary = `Signed in as ${data.account}. Scopes: ${config.scopes.join(', ')}.${
          config.allowWrite ? '' : ' Read-only mode.'
        }`;
      } else if (live.state === 'pending') {
        summary = 'A sign-in is in progress. Complete it, then call authStatus again.';
      } else if (account) {
        summary = `An account (${account.username}) is cached but its token could not be renewed. Call authenticate to sign in again.`;
      } else {
        summary = 'Not signed in. Call the authenticate tool to begin.';
      }

      return ok(summary, data);
    })
  );

  server.registerTool(
    'authenticate',
    {
      title: 'Sign in to Microsoft',
      description:
        'Begin sign-in to the Microsoft account that owns the OneNote notebooks. Returns immediately ' +
        'with the instruction the user must follow (a browser is opened, or a device code is shown). ' +
        'It does NOT wait for the user to finish -- poll `authStatus` afterwards to confirm. ' +
        'Always show the user the returned URL and code verbatim.',
      inputSchema: z.object({}),
      annotations: { readOnlyHint: false, openWorldHint: true, idempotentHint: true }
    },
    handler(async () => {
      // Already good? Don't send the user through a redundant prompt.
      if (await auth.hasCachedAccount()) {
        try {
          await auth.getAccessToken({ interactive: false });
          const account = await auth.getAccount();
          return ok(`Already signed in as ${account.username}. No action needed.`, {
            signedIn: true,
            account: account.username,
            actionRequired: false
          });
        } catch {
          log('Cached account could not be renewed silently; starting interactive sign-in.');
        }
      }

      const state = await session.begin();

      if (state.state === 'failed') {
        return fail(`Sign-in could not be started.\n\n${state.error}`);
      }

      if (state.prompt?.kind === 'device_code') {
        const { verificationUri, userCode, expiresInSeconds } = state.prompt;
        const minutes = expiresInSeconds ? Math.round(expiresInSeconds / 60) : 15;
        return ok(
          [
            'To finish signing in to OneNote:',
            '',
            `1. Open ${verificationUri}`,
            `2. Enter the code: ${userCode}`,
            '3. Sign in with the Microsoft account that owns your notebooks.',
            '',
            `The code expires in about ${minutes} minutes. Once you have signed in, call authStatus to confirm.`
          ].join('\n'),
          {
            actionRequired: true,
            method: 'device_code',
            verificationUri,
            userCode,
            expiresInSeconds
          }
        );
      }

      if (state.prompt?.kind === 'browser') {
        return ok(
          [
            'A browser window was opened on the machine running this server. Complete the Microsoft',
            'sign-in there, then call authStatus to confirm.',
            '',
            'If no window appeared, open this URL manually:',
            state.prompt.url
          ].join('\n'),
          { actionRequired: true, method: 'pkce', url: state.prompt.url }
        );
      }

      return ok(
        'Sign-in has started but no prompt was produced yet. Check the server log, then call authStatus.',
        { actionRequired: true, method: state.method }
      );
    })
  );

  server.registerTool(
    'signOut',
    {
      title: 'Sign out',
      description:
        'Remove the cached Microsoft account and delete the local token cache file. The user must ' +
        'authenticate again afterwards.',
      inputSchema: z.object({}),
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true }
    },
    handler(async () => {
      const result = await auth.signOut();
      return ok(
        result.signedOut
          ? `Signed out. Token cache removed from ${result.cachePath}.`
          : 'No account was signed in. Token cache removed if it existed.',
        result
      );
    })
  );

  // ------------------------------------------------------------------ reads


  server.registerTool(
    'getOutline',
    {
      title: 'Get the full OneNote outline',
      description:
        'START HERE for anything structural. Returns every notebook, section group, and section ' +
        'in ONE call, with IDs and readable paths like "Learn / Cooking". Use this instead of ' +
        'chaining listNotebooks then listSections — it is one request instead of several, and it ' +
        'correctly includes sections nested inside section groups.',
      inputSchema: z.object({
        includePageCounts: z
          .boolean()
          .optional()
          .describe('Also count pages per section. Slightly slower. Defaults to false.')
      }),
      annotations: { readOnlyHint: true, openWorldHint: true }
    },
    handler(async ({ includePageCounts }) => {
      const outline = await onenote.getOutline({
        includePageCounts: includePageCounts ?? false
      });

      const lines = [];
      for (const notebook of outline.notebooks) {
        lines.push(`${notebook.name}${notebook.isDefault ? '  (default)' : ''}`);
        if (!notebook.sections.length) {
          lines.push('    (no sections)');
          continue;
        }
        const byGroup = new Map();
        for (const section of notebook.sections) {
          const key = section.group ?? '';
          if (!byGroup.has(key)) byGroup.set(key, []);
          byGroup.get(key).push(section);
        }
        for (const [group, items] of [...byGroup.entries()].sort()) {
          if (group) lines.push(`    ${group}/`);
          for (const s of items) {
            const indent = group ? '      ' : '    ';
            const count = s.pages !== undefined ? `  (${s.pages} pages)` : '';
            lines.push(`${indent}${s.name}${count}`);
          }
        }
      }

      const { notebooks, sectionGroups, sections } = outline.totals;

      // The tree is already spelled out in the text above. structuredContent
      // therefore carries only the name->ID mapping, which the prose omits.
      // Repeating the whole tree in both fields roughly doubled this tool's cost.
      const ids = {};
      for (const notebook of outline.notebooks) {
        for (const section of notebook.sections) ids[section.path] = section.id;
      }

      return ok(
        `${notebooks} notebook(s), ${sectionGroups} section group(s), ${sections} section(s):\n\n` +
          lines.join('\n') +
          '\n\nPass any of these section names straight to other tools — no ID lookup needed.',
        { sectionIds: ids, totals: outline.totals }
      );
    })
  );

  server.registerTool(
    'listNotebooks',
    {
      title: 'List notebooks',
      description:
        'List notebooks only. Prefer `getOutline` — it returns notebooks AND their sections in ' +
        'the same single call.',
      inputSchema: z.object({ limit: Limit.optional() }),
      annotations: { readOnlyHint: true, openWorldHint: true }
    },
    handler(async ({ limit }) => {
      const notebooks = await onenote.listNotebooks({ limit: limit ?? 50 });
      const summary = notebooks.length
        ? `${notebooks.length} notebook(s):\n${notebooks
            .map((n) => `- ${n.displayName}${n.isDefault ? ' (default)' : ''}`)
            .join('\n')}`
        : 'No notebooks found for this account.';
      return ok(summary, {
        notebooks: notebooks.map((n) => ({
          id: n.id,
          name: n.displayName,
          isDefault: n.isDefault || undefined
        })),
        count: notebooks.length
      });
    })
  );


  server.registerTool(
    'listSections',
    {
      title: 'List sections',
      description:
        'List sections, including those nested inside section groups. Prefer `getOutline` unless ' +
        'you specifically want a flat list.',
      inputSchema: z.object({
        notebook: z
          .string()
          .optional()
          .describe('Notebook name or ID to scope to. Omit for all notebooks.'),
        limit: Limit.optional()
      }),
      annotations: { readOnlyHint: true, openWorldHint: true }
    },
    handler(async ({ notebook, limit }) => {
      let notebookId;
      if (notebook) {
        notebookId = (await onenote.resolveNotebook(notebook)).notebook.id;
      }

      const sections = await onenote.listSections({ notebookId, limit: limit ?? 200 });
      const summary = sections.length
        ? `${sections.length} section(s):\n${sections
            .map(
              (s) =>
                `- ${s.displayName}` +
                `${s.sectionGroupName ? ` (in ${s.sectionGroupName})` : ''}` +
                `${s.notebookName ? ` [${s.notebookName}]` : ''}`
            )
            .join('\n')}`
        : notebook
          ? 'That notebook has no sections.'
          : 'No sections found.';
      return ok(summary, {
        sections: sections.map((s2) => ({
          id: s2.id,
          name: s2.displayName,
          group: s2.sectionGroupName ?? undefined,
          notebook: s2.notebookName ?? undefined
        })),
        count: sections.length
      });
    })
  );

  server.registerTool(
    'listPages',
    {
      title: 'List pages',
      description:
        'List pages, most recently modified first. Pass a section NAME to scope to one section; ' +
        'omit it to list across all sections.',
      inputSchema: z.object({
        section: SectionRef.optional(),
        limit: Limit.optional(),
        orderBy: z
          .enum(['lastModifiedDateTime desc', 'lastModifiedDateTime', 'createdDateTime desc', 'title'])
          .optional()
          .describe('Sort order for results. Defaults to most recently modified first.')
      }),
      annotations: { readOnlyHint: true, openWorldHint: true }
    },
    handler(async ({ section, limit, orderBy }) => {
      const pages = await onenote.listPages({
        section,
        limit: limit ?? DEFAULT_PAGE_LIMIT,
        orderBy: orderBy ?? 'lastModifiedDateTime desc'
      });
      // IDs are deliberately absent from the prose: they are ~50 characters
      // each, the model does not read them, and they are already in
      // structuredContent for when a follow-up call needs one.
      const summary = pages.length
        ? `${pages.length} page(s):\n${pages
            .map(
              (p) =>
                `- ${p.title || '(untitled)'}` +
                `  (${[p.sectionName, shortDate(p.lastModifiedDateTime)].filter(Boolean).join(', ')})`
            )
            .join('\n')}`
        : 'No pages found.';
      return ok(summary, { pages: pages.map(leanPage), count: pages.length });
    })
  );

  server.registerTool(
    'getPage',
    {
      title: 'Get page content',
      description:
        'Read one page by ID. Returns readable text by default; pass format="html" for the raw ' +
        'XHTML, and includeIds=true if you intend to modify the page afterwards.',
      inputSchema: z.object({
        page: PageRef,
        section: SectionRef.optional().describe(
          'Optional: narrow to one section when several pages share a title.'
        ),
        format: z
          .enum(['text', 'html'])
          .optional()
          .describe('"text" for readable Markdown-ish text (default), "html" for raw OneNote XHTML.'),
        includeIds: z
          .boolean()
          .optional()
          .describe('Include OneNote element IDs, required for later PATCH operations.'),
        maxLength: z
          .number()
          .int()
          .min(200)
          .max(200000)
          .optional()
          .describe(
            'Truncate content to this many characters. Defaults to 12000, which covers almost ' +
              'every real page. Raise it only if a page is reported as truncated and you need the rest.'
          )
      }),
      annotations: { readOnlyHint: true, openWorldHint: true }
    },
    handler(async ({ page, section, format, includeIds, maxLength }) => {
      const meta = (await onenote.resolvePage(page, { section })).page;
      const { content, format: actual } = await onenote.getPageContent(meta.id, {
        format: format ?? 'text',
        includeIds: includeIds ?? false
      });

      const cap = maxLength ?? DEFAULT_PAGE_CHARS;
      const truncated = content.length > cap;
      const body = truncated ? `${content.slice(0, cap).trimEnd()}\n\n[...truncated]` : content;

      const header =
        `# ${meta.title || '(untitled)'}` +
        `${meta.sectionName ? `  (${meta.sectionName})` : ''}` +
        `\n_Modified ${shortDate(meta.lastModifiedDateTime)}_\n\n`;

      // The body appears ONLY here, not repeated in structuredContent. Sending
      // it in both fields doubled the cost of every page read.
      return ok(
        header +
          body +
          (truncated
            ? `\n\n(Showing ${cap.toLocaleString()} of ${content.length.toLocaleString()} characters. ` +
              'Call again with a larger maxLength for the rest.)'
            : ''),
        {
          page: { id: meta.id, title: meta.title, section: meta.sectionName ?? undefined },
          format: actual,
          chars: content.length,
          truncated
        }
      );
    })
  );

  server.registerTool(
    'searchPages',
    {
      title: 'Search pages',
      description:
        'Find pages by title (fast — one request). Set searchContent=true to also scan page ' +
        'bodies. OneNote provides no server-side full-text search, so that fetches candidate ' +
        'pages; it runs them in parallel and stops early once enough matches are found, but it ' +
        'is still much slower than a title search. Pass `section` to narrow it dramatically.',
      inputSchema: z.object({
        query: z.string().min(1).describe('Text to search for.'),
        searchContent: z
          .boolean()
          .optional()
          .describe('Also search page body text, not just titles. Slower. Defaults to false.'),
        section: SectionRef.optional().describe(
          'Limit the search to one section, given by name ("Ideas"), path ("Learn / Cooking"), ' +
            'or ID. Strongly recommended alongside searchContent, as it avoids fetching ' +
            'unrelated pages.'
        ),
        limit: z.number().int().min(1).max(100).optional(),
        scanLimit: z
          .number()
          .int()
          .min(1)
          .max(200)
          .optional()
          .describe(
            'Maximum pages to fetch when searchContent is true. Defaults to 30. Raise only if a ' +
              'search genuinely came back empty; each extra page costs time.'
          )
      }),
      annotations: { readOnlyHint: true, openWorldHint: true }
    },
    handler(async ({ query, searchContent, limit, scanLimit, section }) => {
      const results = await onenote.searchPages(query, {
        limit: limit ?? 25,
        searchContent: searchContent ?? false,
        scanLimit: scanLimit ?? 30,
        section
      });

      const stats = results._stats ?? { scanned: 0, elapsedMs: 0 };
      const cost = stats.scanned
        ? `\n\n(Read ${stats.scanned} page(s) in ${(stats.elapsedMs / 1000).toFixed(1)}s. ` +
          'Narrow with `section` to make this faster.)'
        : '';

      const summary = results.length
        ? `${results.length} match(es) for "${query}":\n${results
            .map(
              (r) =>
                `- ${r.title || '(untitled)'} (${r.matchedIn})${
                  r.excerpt ? `\n    ${r.excerpt}` : ''
                }`
            )
            .join('\n')}${cost}`
        : `No pages matched "${query}".${
            searchContent
              ? ' Try raising scanLimit, or check a different section.'
              : ' Try searchContent=true to search inside pages.'
          }${cost}`;

      return ok(summary, {
        results: results.map((r) => ({ ...leanPage(r), matchedIn: r.matchedIn })),
        count: results.length,
        query,
        stats
      });
    })
  );

  // ----------------------------------------------------------------- writes

  server.registerTool(
    'createPage',
    {
      title: 'Create page',
      description:
        'Create a new page. Pass the section by NAME — no lookup needed first. If a default ' +
        'section is configured, `section` can be omitted entirely.',
      inputSchema: z.object({
        section: SectionRef.optional().describe(
          'Section name, path, or ID. Optional if a default section is configured.'
        ),
        title: z.string().min(1).max(255).describe('Page title.'),
        content: z.string().describe('Page body. Plain text unless isHtml is true.'),
        isHtml: z
          .boolean()
          .optional()
          .describe('Treat content as an XHTML fragment instead of plain text.')
      }),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true }
    },
    handler(async ({ section, title, content, isHtml }) => {
      const page = await onenote.createPage({
        section,
        title,
        content: content ?? '',
        isHtml: isHtml ?? false
      });
      return ok(
        `Created "${page.title}" in ${page.sectionName}` +
          `${page.notebookName ? ` (${page.notebookName})` : ''}` +
          `${page.webUrl ? `\n${page.webUrl}` : ''}`,
        { page }
      );
    })
  );

  server.registerTool(
    'appendToPage',
    {
      title: 'Append to page',
      description: 'Add content to the end (or start) of an existing page without replacing it.',
      inputSchema: z.object({
        page: PageRef,
        content: z.string().min(1).describe('Content to add. Plain text unless isHtml is true.'),
        isHtml: z.boolean().optional(),
        position: z
          .enum(['append', 'prepend'])
          .optional()
          .describe('Where to insert relative to existing body content. Defaults to append.')
      }),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true }
    },
    handler(async ({ page, content, isHtml, position }) => {
      const target = (await onenote.resolvePage(page)).page;
      const result = await onenote.appendToPage({
        pageId: target.id,
        content,
        isHtml: isHtml ?? false,
        position: position ?? 'append'
      });
      return ok(
        `Content ${result.position === 'prepend' ? 'prepended to' : 'appended to'} ` +
          `"${target.title}"${target.sectionName ? ` (${target.sectionName})` : ''}.`,
        { ...result, title: target.title }
      );
    })
  );

  server.registerTool(
    'createSection',
    {
      title: 'Create section',
      description: 'Create a new section inside a notebook.',
      inputSchema: z.object({
        notebookId: z.string().min(1).describe('Notebook ID or name.'),
        displayName: z.string().min(1).max(255).describe('Name for the new section.')
      }),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true }
    },
    handler(async ({ notebookId, displayName }) => {
      const section = await onenote.createSection({ notebookId, displayName });
      return ok(`Created section "${section.displayName}" [${section.id}]`, { section });
    })
  );

  server.registerTool(
    'deletePage',
    {
      title: 'Delete page',
      description:
        'Permanently delete a page. This cannot be undone -- confirm the page title with the user ' +
        'before calling.',
      inputSchema: z.object({
        page: PageRef,
        confirmTitle: z
          .string()
          .min(1)
          .describe(
            'The exact current title of the page, as a safeguard. The call fails if it does not match.'
          )
      }),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: true }
    },
    handler(async ({ page, confirmTitle }) => {
      const meta = (await onenote.resolvePage(page)).page;
      const actual = (meta.title || '').trim();
      if (actual.toLowerCase() !== confirmTitle.trim().toLowerCase()) {
        return fail(
          `Refusing to delete. confirmTitle was "${confirmTitle}" but the page is titled "${actual}". ` +
            'Re-read the page and confirm the exact title with the user.'
        );
      }
      const result = await onenote.deletePage(meta.id);
      return ok(`Deleted page "${actual}".`, result);
    })
  );

  return server;
}

// ------------------------------------------------------------------- start up

log(`OneNote MCP server v${version} ready (client ${config.clientId}, tenant ${config.tenant}).`);
if (!(await auth.hasCachedAccount())) {
  log('No cached Microsoft account. Use the "authenticate" tool or run `npm run auth`.');
}
if (config.allowDeviceCode) {
  log(
    'Device code flow is enabled. Note: new Microsoft Entra tenants block this flow by default ' +
      'since 2026-07-01. Unset ONENOTE_ALLOW_DEVICE_CODE to use the browser-based PKCE flow.'
  );
}

// One factory, both protocol eras: modern stateless 2026-07-28 clients and
// 2025-era clients that still open with `initialize`.
const handle = serveStdio(createServer, {
  legacy: 'serve',
  onerror: (error) => log(`Transport error: ${error.message}`)
});

for (const signal of ['SIGINT', 'SIGTERM']) {
  process.on(signal, () => {
    handle.close().finally(() => process.exit(0));
  });
}

/**
 * Minimal .env reader.
 *
 * Avoids a dotenv dependency and, unlike dotenv's default, never overwrites a
 * variable that the host process already set.
 */
async function loadDotEnvIfPresent() {
  if (process.env.ONENOTE_SKIP_DOTENV) return;
  const { readFile } = await import('node:fs/promises');
  try {
    const raw = await readFile(path.join(__dirname, '.env'), 'utf8');
    for (const line of raw.split('\n')) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith('#')) continue;
      const eq = trimmed.indexOf('=');
      if (eq === -1) continue;
      const key = trimmed.slice(0, eq).trim();
      if (!key || key in process.env) continue;
      let value = trimmed.slice(eq + 1).trim();
      if (
        (value.startsWith('"') && value.endsWith('"')) ||
        (value.startsWith("'") && value.endsWith("'"))
      ) {
        value = value.slice(1, -1);
      }
      process.env[key] = value;
    }
  } catch (error) {
    if (error.code !== 'ENOENT') log(`Could not read .env: ${error.message}`);
  }
}
