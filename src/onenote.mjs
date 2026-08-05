/**
 * Microsoft Graph OneNote client.
 *
 * Uses the platform `fetch` (Node >= 18) rather than a Graph SDK. The previous
 * implementation depended on @microsoft/microsoft-graph-client for list calls
 * but then bypassed it with raw fetch for page content, so it carried two HTTP
 * stacks to do one job. Dropping both shrinks the dependency tree that anyone
 * auditing this repo has to trust.
 *
 * Every request goes through `#request`, which:
 *   - attaches a freshly-resolved bearer token,
 *   - retries once on 401 after forcing a token refresh (the old version stored
 *     a bare access token with no refresh path, so it began failing about an
 *     hour after sign-in with an opaque error -- danosb/onenote-mcp#3),
 *   - honours Retry-After on 429/503, which OneNote applies aggressively,
 *   - and converts Graph's error envelope into a readable message.
 */

const USER_AGENT = 'onenote-mcp';

/**
 * How many page fetches run at once during a content search.
 *
 * OneNote allows 120 requests per minute per user per app. Six concurrent
 * fetches at roughly 300ms each is about 20 requests/second in bursts, which
 * finishes a 50-page scan in a few seconds while leaving ample headroom.
 */
const SEARCH_CONCURRENCY = 6;

/**
 * Page text is cached only for the length of one conversation turn or two.
 *
 * A follow-up question ("now find X in those same notes") would otherwise
 * re-download everything. Kept deliberately short, and in memory only, so
 * nothing is stale for long and nothing is ever written to disk.
 */
const CONTENT_CACHE_TTL_MS = 60_000;
const CONTENT_CACHE_MAX_ENTRIES = 200;

export class GraphError extends Error {
  constructor(message, { status, code, requestId, cause } = {}) {
    super(message);
    this.name = 'GraphError';
    this.status = status;
    this.code = code;
    this.requestId = requestId;
    if (cause) this.cause = cause;
  }
}

export class OneNoteClient {
  /**
   * @param {import('./auth.mjs').OneNoteAuth} auth
   * @param {object} config
   * @param {(msg: string) => void} [log]
   */
  constructor(auth, config, log = () => {}) {
    this.auth = auth;
    this.config = config;
    this.log = log;
    /** @type {Map<string, {text: string, at: number}>} in-memory only, never persisted */
    this._contentCache = new Map();
  }

  #cacheGet(key) {
    const hit = this._contentCache.get(key);
    if (!hit) return null;
    if (Date.now() - hit.at > CONTENT_CACHE_TTL_MS) {
      this._contentCache.delete(key);
      return null;
    }
    return hit.text;
  }

  #cacheSet(key, text) {
    // Simple FIFO eviction; this only ever holds one conversation's worth.
    if (this._contentCache.size >= CONTENT_CACHE_MAX_ENTRIES) {
      this._contentCache.delete(this._contentCache.keys().next().value);
    }
    this._contentCache.set(key, { text, at: Date.now() });
  }

  // ---------------------------------------------------------------- transport

  async #request(
    pathOrUrl,
    { method = 'GET', headers = {}, body, raw = false, retry = true, forceRefresh = false } = {}
  ) {
    const url = pathOrUrl.startsWith('http')
      ? pathOrUrl
      : `${this.config.graphBaseUrl}${pathOrUrl}`;

    const token = await this.auth.getAccessToken({ interactive: false, forceRefresh });

    const response = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        'User-Agent': USER_AGENT,
        ...headers
      },
      body
    });

    if (response.status === 401 && retry) {
      // Graph rejected a token MSAL still considered valid (revoked session,
      // password change, new Conditional Access policy). Redeem the refresh
      // token and try exactly once more before surfacing an error.
      this.log('Graph returned 401; forcing token refresh and retrying once.');
      return this.#request(pathOrUrl, {
        method,
        headers,
        body,
        raw,
        retry: false,
        forceRefresh: true
      });
    }

    if ((response.status === 429 || response.status === 503) && retry) {
      const wait = Number(response.headers.get('retry-after') || 2);
      const ms = Math.min(Math.max(wait, 1), 30) * 1000;
      this.log(`Graph returned ${response.status}; retrying after ${ms}ms.`);
      await new Promise((resolve) => setTimeout(resolve, ms));
      return this.#request(pathOrUrl, { method, headers, body, raw, retry: false });
    }

    if (!response.ok) {
      throw await toGraphError(response);
    }

    if (raw) return response.text();
    if (response.status === 204) return null;
    return response.json();
  }

  /**
   * Follow @odata.nextLink until `limit` items are collected.
   * The original code read only the first page, silently hiding notes for anyone
   * with more than ~20 pages.
   */
  async #collect(path, { limit = 100 } = {}) {
    const items = [];
    let next = path;

    while (next && items.length < limit) {
      const payload = await this.#request(next);
      const batch = payload?.value ?? [];
      items.push(...batch);
      next = payload?.['@odata.nextLink'] ?? null;
      if (!batch.length) break;
    }

    return items.slice(0, limit);
  }

  #assertWritable() {
    if (!this.config.allowWrite) {
      throw new GraphError(
        'This server is running read-only (ONENOTE_ALLOW_WRITE=false). Write tools are disabled.',
        { status: 403, code: 'writeDisabled' }
      );
    }
  }

  // ------------------------------------------------------------------- reads

  async whoAmI() {
    // Deliberately NOT a Graph /me call. Reading /me requires the User.Read
    // permission, which would add a third line ("View your basic profile") to
    // the consent screen purely so we could print a name. MSAL already holds
    // the signed-in identity from the sign-in itself, so we use that: same
    // information, one fewer permission, no network request.
    const account = await this.auth.getAccount();
    if (!account) {
      throw new GraphError('Not signed in.', { status: 401 });
    }
    return {
      displayName: account.name ?? account.username ?? 'Signed-in user',
      userPrincipalName: account.username ?? null,
      id: account.homeAccountId ?? null,
      source: 'local sign-in token (no Microsoft Graph call made)'
    };
  }

  async listNotebooks({ limit = 50 } = {}) {
    const notebooks = await this.#collect(
      '/me/onenote/notebooks?$select=id,displayName,createdDateTime,lastModifiedDateTime,isDefault,links&$orderby=displayName',
      { limit }
    );
    return notebooks.map(shapeNotebook);
  }

  async getNotebook(notebookId) {
    const nb = await this.#request(
      `/me/onenote/notebooks/${encodeURIComponent(notebookId)}?$select=id,displayName,createdDateTime,lastModifiedDateTime,isDefault,links`
    );
    return shapeNotebook(nb);
  }

  /**
   * List sections, optionally scoped to one notebook.
   *
   * IMPORTANT: this deliberately does NOT use
   * `/me/onenote/notebooks/{id}/sections`. That endpoint returns only sections
   * sitting directly under the notebook -- any section inside a *section group*
   * is silently omitted. In a real notebook that hid 7 of 9 sections, which made
   * the assistant appear unable to find notes that plainly existed.
   *
   * Instead we read the flat `/me/onenote/sections` collection, which includes
   * every section regardless of nesting, and filter by parent notebook
   * ourselves. One request, complete results.
   *
   * @param {string} [notebookId] When omitted, returns sections across all notebooks.
   */
  async listSections({ notebookId, limit = 200 } = {}) {
    const sections = await this.#collect(
      '/me/onenote/sections' +
        '?$select=id,displayName,createdDateTime,lastModifiedDateTime' +
        '&$expand=parentNotebook($select=id,displayName),parentSectionGroup($select=id,displayName)' +
        '&$orderby=displayName',
      { limit: Math.max(limit, 200) }
    );

    const shaped = sections.map(shapeSection);
    const scoped = notebookId ? shaped.filter((s) => s.notebookId === notebookId) : shaped;
    return scoped.slice(0, limit);
  }

  /** Section groups, so the outline can show real nesting. */
  async listSectionGroups({ limit = 200 } = {}) {
    const groups = await this.#collect(
      '/me/onenote/sectionGroups' +
        '?$select=id,displayName' +
        '&$expand=parentNotebook($select=id,displayName),parentSectionGroup($select=id,displayName)' +
        '&$orderby=displayName',
      { limit }
    );
    return groups.map((g) => ({
      id: g.id,
      displayName: g.displayName,
      notebookId: g.parentNotebook?.id ?? null,
      notebookName: g.parentNotebook?.displayName ?? null,
      parentGroupId: g.parentSectionGroup?.id ?? null
    }));
  }

  /**
   * The whole notebook/section-group/section tree in one call.
   *
   * Locating anything used to cost four round-trips (notebooks -> sections ->
   * pages -> page), each needing its own approval click. This collapses
   * discovery into a single request pair so the assistant can go straight to
   * the thing the user named.
   */
  async getOutline({ includePageCounts = false } = {}) {
    const [notebooks, groups, sections] = await Promise.all([
      this.listNotebooks({ limit: 100 }),
      this.listSectionGroups({ limit: 200 }),
      this.listSections({ limit: 400 })
    ]);

    let pageCounts = null;
    if (includePageCounts) {
      const pages = await this.#collect(
        '/me/onenote/pages?$select=id&$expand=parentSection($select=id)&$top=100',
        { limit: 1000 }
      );
      pageCounts = new Map();
      for (const page of pages) {
        const sectionId = page.parentSection?.id;
        if (sectionId) pageCounts.set(sectionId, (pageCounts.get(sectionId) ?? 0) + 1);
      }
    }

    const groupsById = new Map(groups.map((g) => [g.id, g]));

    /** Human-readable path such as "Work / Projects / Q3 Planning". */
    const pathFor = (section) => {
      const parts = [];
      let cursor = section.sectionGroupId;
      let guard = 0;
      while (cursor && guard++ < 20) {
        const group = groupsById.get(cursor);
        if (!group) break;
        parts.unshift(group.displayName);
        cursor = group.parentGroupId;
      }
      return [section.notebookName, ...parts, section.displayName].filter(Boolean).join(' / ');
    };

    const tree = notebooks.map((notebook) => ({
      id: notebook.id,
      name: notebook.displayName,
      isDefault: notebook.isDefault,
      sections: sections
        .filter((s) => s.notebookId === notebook.id)
        .map((s) => ({
          id: s.id,
          name: s.displayName,
          group: s.sectionGroupName ?? null,
          path: pathFor(s),
          lastModified: s.lastModifiedDateTime,
          ...(pageCounts ? { pages: pageCounts.get(s.id) ?? 0 } : {})
        }))
    }));

    return {
      notebooks: tree,
      totals: {
        notebooks: notebooks.length,
        sectionGroups: groups.length,
        sections: sections.length
      }
    };
  }

  // ------------------------------------------------------------- resolution

  /**
   * Turn a human name into a section, so callers never have to hold a GUID.
   *
   * Accepts an ID, an exact name, a "Notebook / Section" path, or a partial
   * name. Ambiguity is reported rather than guessed at, because silently
   * writing to the wrong section is worse than asking.
   *
   * @returns {Promise<{section: object, matched: string}>}
   */
  async resolveSection(nameOrId, { sections } = {}) {
    const query = String(nameOrId ?? '').trim();
    if (!query) throw new GraphError('A section name or ID is required.', { status: 400 });

    const all = sections ?? (await this.listSections({ limit: 400 }));

    // An ID looks nothing like a name, so test that first and cheaply.
    const byId = all.find((s) => s.id === query);
    if (byId) return { section: byId, matched: 'id' };

    const lower = query.toLowerCase();

    const exact = all.filter((s) => (s.displayName || '').toLowerCase() === lower);
    if (exact.length === 1) return { section: exact[0], matched: 'name' };
    if (exact.length > 1) throw ambiguous(query, exact);

    // "Notebook / Section" or "Group / Section"
    if (query.includes('/')) {
      const wanted = query.split('/').map((p) => p.trim().toLowerCase()).filter(Boolean);
      const leaf = wanted[wanted.length - 1];
      const byPath = all.filter((s) => {
        if ((s.displayName || '').toLowerCase() !== leaf) return false;
        return wanted.slice(0, -1).every(
          (part) =>
            (s.notebookName || '').toLowerCase().includes(part) ||
            (s.sectionGroupName || '').toLowerCase().includes(part)
        );
      });
      if (byPath.length === 1) return { section: byPath[0], matched: 'path' };
      if (byPath.length > 1) throw ambiguous(query, byPath);
    }

    const partial = all.filter((s) => (s.displayName || '').toLowerCase().includes(lower));
    if (partial.length === 1) return { section: partial[0], matched: 'partial' };
    if (partial.length > 1) throw ambiguous(query, partial);

    throw new GraphError(
      `No section matches "${query}".\n\nAvailable sections:\n` +
        all
          .slice(0, 40)
          .map((s) => `  - ${s.displayName}${s.notebookName ? ` (${s.notebookName})` : ''}`)
          .join('\n'),
      { status: 404 }
    );
  }

  /**
   * Turn a page title into a page.
   *
   * This exists because of a real failure: page IDs were removed from tool
   * output text to save tokens, on the assumption that callers would read them
   * from `structuredContent`. MCP clients generally surface only the text to the
   * model, so the assistant could see "AI Search" in search results and had no
   * ID to pass to getPage -- it would send the title, which was rejected.
   *
   * Rather than push ~50-character IDs back into every line of prose, pages are
   * now addressable the same way sections are: by name.
   *
   * @param {string} titleOrId
   * @param {object} [options]
   * @param {string} [options.section] Narrow the search to one section.
   */
  async resolvePage(titleOrId, { section } = {}) {
    const query = String(titleOrId ?? '').trim();
    if (!query) throw new GraphError('A page title or ID is required.', { status: 400 });

    // OneNote page IDs have a distinctive shape; try an exact ID hit first so a
    // genuine ID never pays for a listing.
    if (/^[0-9a-zA-Z]+-[0-9A-Fa-f!]/.test(query) && query.includes('!')) {
      try {
        const direct = await this.getPageMetadata(query);
        if (direct?.id) return { page: direct, matched: 'id' };
      } catch {
        // Not an ID after all; fall through to title matching.
      }
    }

    const pages = await this.listPages({ section, limit: 200 });

    const byId = pages.find((p) => p.id === query);
    if (byId) return { page: byId, matched: 'id' };

    const lower = query.toLowerCase();

    const exact = pages.filter((p) => (p.title || '').toLowerCase() === lower);
    if (exact.length === 1) return { page: exact[0], matched: 'title' };
    if (exact.length > 1) throw ambiguousPage(query, exact);

    const partial = pages.filter((p) => (p.title || '').toLowerCase().includes(lower));
    if (partial.length === 1) return { page: partial[0], matched: 'partial' };
    if (partial.length > 1) throw ambiguousPage(query, partial);

    throw new GraphError(
      `No page matches "${query}"${section ? ` in section "${section}"` : ''}.\n\n` +
        'Use searchPages to find it, or listPages to browse. Recent pages:\n' +
        pages
          .slice(0, 15)
          .map((p) => `  - ${p.title || '(untitled)'}${p.sectionName ? ` (${p.sectionName})` : ''}`)
          .join('\n'),
      { status: 404 }
    );
  }

  /** Same idea for notebooks. */
  async resolveNotebook(nameOrId) {
    const query = String(nameOrId ?? '').trim();
    if (!query) throw new GraphError('A notebook name or ID is required.', { status: 400 });

    const all = await this.listNotebooks({ limit: 100 });

    const byId = all.find((n) => n.id === query);
    if (byId) return { notebook: byId, matched: 'id' };

    const lower = query.toLowerCase();
    const exact = all.filter((n) => (n.displayName || '').toLowerCase() === lower);
    if (exact.length === 1) return { notebook: exact[0], matched: 'name' };

    const partial = all.filter((n) => (n.displayName || '').toLowerCase().includes(lower));
    if (partial.length === 1) return { notebook: partial[0], matched: 'partial' };
    if (partial.length > 1) {
      throw new GraphError(
        `"${query}" matches several notebooks: ${partial.map((n) => n.displayName).join(', ')}. ` +
          'Be more specific.',
        { status: 400 }
      );
    }

    throw new GraphError(
      `No notebook matches "${query}". Available: ${all.map((n) => n.displayName).join(', ')}`,
      { status: 404 }
    );
  }

  /**
   * @param {object} [options]
   * @param {string} [options.section] Section name, path, or ID. Omit for all sections.
   * @param {string} [options.sectionId] Explicit ID (takes precedence).
   */
  async listPages({ section, sectionId, limit = 100, orderBy = 'lastModifiedDateTime desc' } = {}) {
    let resolvedId = sectionId;
    if (!resolvedId && section) {
      resolvedId = (await this.resolveSection(section)).section.id;
    }

    const base = resolvedId
      ? `/me/onenote/sections/${encodeURIComponent(resolvedId)}/pages`
      : '/me/onenote/pages';
    const pages = await this.#collect(
      `${base}?$select=id,title,createdDateTime,lastModifiedDateTime,contentUrl,links&$orderby=${encodeURIComponent(orderBy)}`,
      { limit }
    );
    return pages.map(shapePage);
  }

  async getPageMetadata(pageId) {
    const page = await this.#request(
      `/me/onenote/pages/${encodeURIComponent(pageId)}?$select=id,title,createdDateTime,lastModifiedDateTime,contentUrl,links&$expand=parentSection($select=id,displayName)`
    );
    return shapePage(page);
  }

  /**
   * Fetch a page's content.
   *
   * @param {string} pageId
   * @param {object} [options]
   * @param {'text'|'html'} [options.format='text']
   * @param {boolean} [options.includeIds=false] Ask Graph for element IDs, which
   *   are required if you intend to PATCH the page afterwards.
   */
  async getPageContent(pageId, { format = 'text', includeIds = false } = {}) {
    const cacheKey = `${pageId}:${format}:${includeIds ? 1 : 0}`;
    const cached = this.#cacheGet(cacheKey);
    if (cached !== null) return { format, content: cached, cached: true };

    const query = includeIds ? '?includeIDs=true' : '';
    const html = await this.#request(
      `/me/onenote/pages/${encodeURIComponent(pageId)}/content${query}`,
      { raw: true, headers: { Accept: 'text/html' } }
    );

    if (format === 'html') {
      this.#cacheSet(cacheKey, html);
      return { format: 'html', content: html };
    }

    const { htmlToText } = await import('./html.mjs');
    const text = htmlToText(html);
    this.#cacheSet(cacheKey, text);
    return { format: 'text', content: text };
  }

  /**
   * Search pages.
   *
   * Graph's OneNote endpoint has no server-side full-text `$search`, so title
   * filtering is done with `$filter contains(...)` where possible and content
   * matching requires fetching page bodies. `searchContent` is therefore
   * explicitly opt-in and capped, because it costs one request per page.
   */
  async searchPages(query, { limit = 25, searchContent = false, scanLimit = 50, section } = {}) {
    const term = String(query || '').trim();
    if (!term) throw new GraphError('A non-empty search query is required.', { status: 400 });

    // Scoping the search to one section makes content search dramatically
    // cheaper, since it only fetches pages the user actually cares about.
    const pages = await this.listPages({
      section,
      limit: Math.max(limit, scanLimit)
    });
    const lower = term.toLowerCase();

    const titleMatches = pages.filter((p) => (p.title || '').toLowerCase().includes(lower));
    const results = titleMatches.map((p) => ({ ...p, matchedIn: 'title' }));

    let scanned = 0;
    let elapsedMs = 0;

    if (searchContent) {
      // OneNote offers no server-side full-text search: `$search` is unsupported
      // on /me/onenote/pages, and the Microsoft Search API returns nothing for
      // personal accounts. So matching content means fetching pages.
      //
      // The first implementation awaited each page in turn, which meant ~50
      // sequential round-trips and 15-25 seconds of waiting. Running a small
      // pool of concurrent fetches instead cuts that by roughly the pool size,
      // while staying well under OneNote's 120-requests-per-minute per-user cap.
      const started = Date.now();
      const seen = new Set(results.map((r) => r.id));

      // Pages arrive newest-first, and recent notes are what people usually
      // mean, so early exit tends to trigger quickly.
      const candidates = pages.filter((p) => !seen.has(p.id)).slice(0, scanLimit);

      const matches = await this.#scanConcurrently(candidates, lower, term, limit - results.length);
      results.push(...matches.found);
      scanned = matches.scanned;
      elapsedMs = Date.now() - started;
    }

    return Object.assign(results.slice(0, limit), {
      _stats: { scanned, elapsedMs }
    });
  }

  /**
   * Fetch and scan pages through a bounded worker pool.
   *
   * Stops launching work as soon as enough matches exist, so a query whose hits
   * are all recent finishes almost immediately regardless of scanLimit.
   */
  async #scanConcurrently(candidates, lowerTerm, originalTerm, wanted) {
    const found = [];
    let cursor = 0;
    let scanned = 0;
    let stop = false;

    const worker = async () => {
      while (!stop) {
        const index = cursor++;
        if (index >= candidates.length) return;

        const page = candidates[index];
        try {
          const { content } = await this.getPageContent(page.id, { format: 'text' });
          scanned += 1;

          const at = content.toLowerCase().indexOf(lowerTerm);
          if (at !== -1) {
            found.push({
              ...page,
              matchedIn: 'content',
              excerpt: excerptAround(content, at, originalTerm.length)
            });
            if (found.length >= wanted) stop = true;
          }
        } catch (error) {
          this.log(`Skipping "${page.title}" during content search: ${error.message}`);
        }
      }
    };

    const poolSize = Math.min(SEARCH_CONCURRENCY, Math.max(1, candidates.length));
    await Promise.all(Array.from({ length: poolSize }, worker));

    // Preserve the newest-first ordering the caller expects; the pool finishes
    // out of order.
    found.sort(
      (a, b) => new Date(b.lastModifiedDateTime ?? 0) - new Date(a.lastModifiedDateTime ?? 0)
    );

    return { found, scanned };
  }

  // ------------------------------------------------------------------ writes

  /**
   * Create a page in a section.
   *
   * @param {object} args
   * @param {string} args.sectionId
   * @param {string} args.title
   * @param {string} args.content Plain text unless `isHtml` is set.
   * @param {boolean} [args.isHtml=false]
   */
  async createPage({ section, sectionId, title, content, isHtml = false }) {
    this.#assertWritable();

    // Accept a name, a path, or an ID. Falls back to the pinned default section
    // so "make me a note" works with no discovery at all.
    const target = sectionId ?? section ?? this.config.defaultSection;
    if (!target) {
      throw new GraphError(
        'No section given. Pass a section name (e.g. "Ideas" or "Learn / Cooking"), ' +
          'or set a default section in the extension settings.',
        { status: 400 }
      );
    }
    const resolved = await this.resolveSection(target);
    const targetId = resolved.section.id;

    const { buildPageHtml } = await import('./html.mjs');
    const html = buildPageHtml(title, content, { isHtml });

    const created = await this.#request(
      `/me/onenote/sections/${encodeURIComponent(targetId)}/pages`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/xhtml+xml' },
        body: html
      }
    );

    return {
      ...shapePage(created),
      sectionId: targetId,
      sectionName: resolved.section.displayName,
      notebookName: resolved.section.notebookName
    };
  }

  /**
   * Append to an existing page's body.
   *
   * @param {object} args
   * @param {string} args.pageId
   * @param {string} args.content
   * @param {boolean} [args.isHtml=false]
   * @param {'append'|'prepend'} [args.position='append']
   */
  async appendToPage({ pageId, content, isHtml = false, position = 'append' }) {
    this.#assertWritable();
    if (!pageId) throw new GraphError('pageId is required.', { status: 400 });

    const { escapeHtml } = await import('./html.mjs');
    const fragment = isHtml
      ? content
      : String(content || '')
          .split(/\n{2,}/)
          .map((para) => `<p>${escapeHtml(para).replace(/\n/g, '<br />')}</p>`)
          .join('');

    await this.#request(`/me/onenote/pages/${encodeURIComponent(pageId)}/content`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify([
        {
          target: 'body',
          action: position === 'prepend' ? 'prepend' : 'append',
          content: fragment
        }
      ])
    });

    return { pageId, position, updated: true };
  }

  async deletePage(pageId) {
    this.#assertWritable();
    if (!pageId) throw new GraphError('pageId is required.', { status: 400 });
    await this.#request(`/me/onenote/pages/${encodeURIComponent(pageId)}`, { method: 'DELETE' });
    return { pageId, deleted: true };
  }

  async createSection({ notebookId, displayName }) {
    this.#assertWritable();
    if (!notebookId) throw new GraphError('notebookId is required.', { status: 400 });
    const section = await this.#request(
      `/me/onenote/notebooks/${encodeURIComponent(notebookId)}/sections`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ displayName })
      }
    );
    return shapeSection(section);
  }
}

// ------------------------------------------------------------------- shaping

/**
 * Response shaping keeps tool output small and predictable. The old server
 * returned `JSON.stringify(response.value)` verbatim, which pushed large blobs
 * of OData metadata and self-links into the model's context on every call.
 */

function shapeNotebook(nb) {
  if (!nb) return null;
  return {
    id: nb.id,
    displayName: nb.displayName,
    isDefault: nb.isDefault ?? null,
    createdDateTime: nb.createdDateTime,
    lastModifiedDateTime: nb.lastModifiedDateTime,
    webUrl: nb.links?.oneNoteWebUrl?.href ?? null
  };
}

function shapeSection(section) {
  if (!section) return null;
  return {
    id: section.id,
    displayName: section.displayName,
    notebookId: section.parentNotebook?.id ?? null,
    notebookName: section.parentNotebook?.displayName ?? null,
    sectionGroupId: section.parentSectionGroup?.id ?? null,
    sectionGroupName: section.parentSectionGroup?.displayName ?? null,
    createdDateTime: section.createdDateTime,
    lastModifiedDateTime: section.lastModifiedDateTime
  };
}

function ambiguousPage(query, matches) {
  return new GraphError(
    `"${query}" matches ${matches.length} pages:\n` +
      matches
        .slice(0, 15)
        .map(
          (p) =>
            `  - ${p.title || '(untitled)'}` +
            `${p.sectionName ? ` (in ${p.sectionName})` : ''}` +
            `${p.lastModifiedDateTime ? ` — modified ${p.lastModifiedDateTime.slice(0, 10)}` : ''}`
        )
        .join('\n') +
      '\n\nNarrow it with the `section` argument, or give the full exact title.',
    { status: 400, code: 'ambiguousPage' }
  );
}

/** Ambiguity is surfaced with the candidates, so the user can just pick one. */
function ambiguous(query, matches) {
  return new GraphError(
    `"${query}" matches ${matches.length} sections:\n` +
      matches
        .map(
          (s) =>
            `  - ${s.displayName}` +
            `${s.sectionGroupName ? ` (in ${s.sectionGroupName})` : ''}` +
            `${s.notebookName ? ` [${s.notebookName}]` : ''}`
        )
        .join('\n') +
      '\n\nUse a fuller path like "Notebook / Section", or pass the section ID.',
    { status: 400, code: 'ambiguousSection' }
  );
}

function shapePage(page) {
  if (!page) return null;
  return {
    id: page.id,
    title: page.title,
    sectionId: page.parentSection?.id ?? null,
    sectionName: page.parentSection?.displayName ?? null,
    createdDateTime: page.createdDateTime,
    lastModifiedDateTime: page.lastModifiedDateTime,
    webUrl: page.links?.oneNoteWebUrl?.href ?? null
  };
}

function excerptAround(text, index, matchLength, radius = 120) {
  const start = Math.max(0, index - radius);
  const end = Math.min(text.length, index + matchLength + radius);
  return `${start > 0 ? '...' : ''}${text.slice(start, end).replace(/\s+/g, ' ').trim()}${
    end < text.length ? '...' : ''
  }`;
}

async function toGraphError(response) {
  const requestId = response.headers.get('request-id') || undefined;
  let code;
  let message;

  try {
    const text = await response.text();
    try {
      const parsed = JSON.parse(text);
      code = parsed?.error?.code;
      message = parsed?.error?.message;
    } catch {
      message = text.slice(0, 500);
    }
  } catch {
    /* body already consumed or empty */
  }

  const detail = message || response.statusText || 'Unknown error';

  if (response.status === 401) {
    return new GraphError(
      `Microsoft Graph rejected the access token (401). Re-authenticate: run \`npm run auth\`, or call the \`authenticate\` tool. Detail: ${detail}`,
      { status: 401, code, requestId }
    );
  }

  if (response.status === 403) {
    return new GraphError(
      `Microsoft Graph denied the request (403). The signed-in account may lack the required scope, ` +
        `or your app registration was not consented for it. Detail: ${detail}`,
      { status: 403, code, requestId }
    );
  }

  if (response.status === 404) {
    return new GraphError(
      `Not found (404). The notebook, section, or page ID may be stale -- list them again to get current IDs. Detail: ${detail}`,
      { status: 404, code, requestId }
    );
  }

  return new GraphError(`Microsoft Graph error ${response.status}: ${detail}`, {
    status: response.status,
    code,
    requestId
  });
}
