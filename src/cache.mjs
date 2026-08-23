/**
 * On-disk cache of already-extracted page text.
 *
 * WHY THIS EXISTS
 *
 * Reading a page costs the same three things every time: a download from
 * Microsoft, an HTML-to-text extraction, and -- by far the largest -- the
 * tokens spent putting that text in front of the model. The first two are pure
 * waste when the page has not changed since you last read it, which is the
 * common case: notes are read far more often than they are edited.
 *
 * So the extracted text is kept, and reused whenever OneNote confirms the page
 * is untouched.
 *
 * WHOSE DATA THIS IS
 *
 * Yours, on your own computer, and nowhere else. Specifically:
 *
 *   - The file lives in your OS config directory, mode 0600 inside a 0700
 *     directory -- the same protection as your sign-in token, readable only by
 *     your user account.
 *   - It is scoped to the signed-in Microsoft account. The filename is a hash
 *     of the account ID, so two accounts on one machine can never read each
 *     other's cache, and the filename itself does not spell out your address.
 *   - It is NEVER transmitted. Nothing in this file reads the network. There is
 *     no server to send it to; see PRIVACY.md.
 *   - Signing out deletes it.
 *
 * WHAT MAKES IT SAFE TO REUSE
 *
 * Entries are validated against `lastModifiedDateTime`, the timestamp Microsoft
 * already returns in the page listing the caller just made. If the timestamp
 * differs by even a second, the entry is discarded and the page is re-fetched.
 * A cached answer is therefore never staler than the metadata the caller is
 * holding, which is why there is no expiry clock: time is the wrong test, and
 * an expiry would both serve stale text inside the window and discard perfectly
 * good text outside it.
 */

import { createHash, randomBytes } from 'node:crypto';
import fsp from 'node:fs/promises';
import path from 'node:path';

/**
 * Caps. Generous enough to cover the notebooks anyone actually reads from,
 * small enough that this never becomes a surprise on someone's disk.
 */
const MAX_ENTRIES = 500;
const MAX_TOTAL_CHARS = 4_000_000; // ~4 MB of text, roughly 1,000 typical pages
const MAX_ENTRY_CHARS = 200_000; // a single enormous page should not evict everything

/** Writes are batched: a 50-page search should produce one file write, not 50. */
const FLUSH_DEBOUNCE_MS = 500;

/** Bumped if the stored shape ever changes, so old files are ignored not misread. */
const FORMAT = 1;

/** Cache disabled entirely -- every method is a no-op that costs nothing. */
class NullCache {
  get enabled() {
    return false;
  }
  async get() {
    return null;
  }
  async set() {}
  async flush() {}
  async close() {}
  async clear() {
    return { removed: 0 };
  }
  async stats() {
    return { enabled: false, entries: 0, bytes: 0, path: null };
  }
}

export class PageCache {
  /**
   * @param {object} options
   * @param {string} options.dir          directory to store cache files in
   * @param {() => Promise<string|null>} options.accountId resolves the signed-in account
   * @param {(msg: string) => void} [options.log]
   */
  constructor({ dir, accountId, log = () => {} }) {
    this.dir = dir;
    this.getAccountId = accountId;
    this.log = log;

    /** @type {Map<string, {m: string, t: string, a: number}>|null} */
    this._entries = null;
    this._file = null;
    this._dirty = false;
    this._timer = null;
    this._loading = null;
    this._closed = false;

    /**
     * Unique per instance, not just per process. Two clients in one process
     * (a test, or a host that opens a second connection) would otherwise pick
     * the same temp filename and race each other's rename. On POSIX that is
     * merely sloppy; on Windows renaming a file another handle is writing
     * fails outright with EPERM.
     */
    this._tmpSuffix = `${process.pid}.${randomBytes(4).toString('hex')}`;
  }

  get enabled() {
    return true;
  }

  /**
   * Resolve the per-account file path, and load it once.
   *
   * Any failure here -- no account yet, unreadable file, corrupt JSON -- ends
   * with an empty cache rather than an error. A cache is an optimisation; it
   * must never be the reason a read fails.
   */
  async #load() {
    if (this._entries) return this._entries;
    if (this._loading) return this._loading;

    this._loading = (async () => {
      const account = await this.getAccountId().catch(() => null);
      if (!account) {
        // Not signed in yet. Stay in memory only rather than writing to a
        // shared filename that a later account could pick up.
        this._entries = new Map();
        return this._entries;
      }

      // Hash rather than store the account ID: the cache is scoped per account
      // without the filename disclosing which account it belongs to.
      const key = createHash('sha256').update(String(account)).digest('hex').slice(0, 16);
      this._file = path.join(this.dir, `pages-${key}.json`);

      try {
        const raw = await fsp.readFile(this._file, 'utf8');
        const parsed = JSON.parse(raw);
        this._entries =
          parsed?.format === FORMAT && parsed.entries
            ? new Map(Object.entries(parsed.entries))
            : new Map();
      } catch {
        this._entries = new Map();
      }

      return this._entries;
    })();

    try {
      return await this._loading;
    } finally {
      this._loading = null;
    }
  }

  /**
   * Return cached text for a page, but only if OneNote agrees it is unchanged.
   *
   * `lastModified` is required. Without it there is nothing to validate against,
   * and returning unvalidated text would be a correctness bug -- so a caller
   * that cannot supply it simply gets a miss.
   */
  async get(pageId, { lastModified, variant = 'text' } = {}) {
    if (!lastModified) return null;

    const entries = await this.#load();
    const key = `${pageId}:${variant}`;
    const hit = entries.get(key);
    if (!hit) return null;

    if (hit.m !== lastModified) {
      // The page was edited. Drop it rather than let a stale copy linger.
      entries.delete(key);
      this.#markDirty();
      return null;
    }

    // Touch in memory only. This improves eviction order, but it is not worth
    // a disk write on every read -- and a write per read multiplies the chance
    // of colliding with anything else touching the file.
    hit.a = Date.now();
    return hit.t;
  }

  async set(pageId, { lastModified, variant = 'text', text } = {}) {
    if (!lastModified || typeof text !== 'string') return;
    if (text.length > MAX_ENTRY_CHARS) return;

    const entries = await this.#load();
    entries.set(`${pageId}:${variant}`, { m: lastModified, t: text, a: Date.now() });
    this.#evict(entries);
    this.#markDirty();
  }

  /** Least-recently-used eviction, on both count and total size. */
  #evict(entries) {
    let total = 0;
    for (const entry of entries.values()) total += entry.t.length;
    if (entries.size <= MAX_ENTRIES && total <= MAX_TOTAL_CHARS) return;

    const byAge = [...entries.entries()].sort((a, b) => a[1].a - b[1].a);
    for (const [key, entry] of byAge) {
      if (entries.size <= MAX_ENTRIES && total <= MAX_TOTAL_CHARS) break;
      entries.delete(key);
      total -= entry.t.length;
    }
  }

  #markDirty() {
    if (this._closed) return;
    this._dirty = true;
    if (this._timer) return;
    this._timer = setTimeout(() => {
      this._timer = null;
      this.flush().catch(() => {});
    }, FLUSH_DEBOUNCE_MS);
    // Never hold the process open for a cache write; `beforeExit` covers the
    // case where everything else finished first.
    this._timer.unref?.();
  }

  /**
   * Write the cache out. Atomic (temp file + rename) and 0600, matching how the
   * token cache is written, so a crash mid-write cannot leave a corrupt file
   * and no other user on the machine can read it.
   */
  async flush() {
    // Cancel any pending debounce first. Otherwise a manual flush leaves a
    // timer armed that fires later, after the caller believed it was done.
    if (this._timer) {
      clearTimeout(this._timer);
      this._timer = null;
    }

    if (!this._dirty || !this._file || !this._entries) return;
    this._dirty = false;

    const payload = JSON.stringify({
      format: FORMAT,
      entries: Object.fromEntries(this._entries)
    });

    const tmp = `${this._file}.${this._tmpSuffix}.tmp`;
    try {
      await fsp.mkdir(this.dir, { recursive: true, mode: 0o700 });
      await fsp.writeFile(tmp, payload, { mode: 0o600 });
      await fsp.rename(tmp, this._file);
      await fsp.chmod(this._file, 0o600).catch(() => {});
    } catch (error) {
      this.log(`Could not write page cache: ${error.message}`);
      await fsp.rm(tmp, { force: true }).catch(() => {});
    }
  }

  /**
   * Write anything outstanding and stop scheduling further writes.
   *
   * Without this, a cache whose owner has gone away can still fire a queued
   * write half a second later and recreate a directory that was just deleted.
   */
  async close() {
    if (this._timer) {
      clearTimeout(this._timer);
      this._timer = null;
    }
    await this.flush();
    this._closed = true;
  }

  /** Delete everything for the signed-in account. Used by sign-out and the CLI. */
  async clear() {
    const entries = await this.#load();
    const removed = entries.size;

    entries.clear();
    this._dirty = false;
    if (this._timer) {
      clearTimeout(this._timer);
      this._timer = null;
    }
    if (this._file) await fsp.rm(this._file, { force: true }).catch(() => {});

    return { removed };
  }

  async stats() {
    const entries = await this.#load();
    let bytes = 0;
    for (const entry of entries.values()) bytes += entry.t.length;
    return { enabled: true, entries: entries.size, bytes, path: this._file };
  }
}

/**
 * Build the cache the configuration asks for.
 *
 * `off` and `memory` both return a cache that never touches the disk -- the
 * 60-second in-memory cache inside the Graph client is unaffected either way,
 * so `memory` keeps same-conversation speed while writing nothing down.
 */
export function createPageCache(config, { accountId, log } = {}) {
  if (config.cacheMode !== 'disk') return new NullCache();

  return new PageCache({
    dir: path.join(path.dirname(config.cachePath), 'cache'),
    accountId,
    log
  });
}

/**
 * Remove every account's cached pages, without needing to be signed in.
 *
 * Sign-out uses this: by the time the tokens are gone the account can no longer
 * be resolved, so the file has to be findable by pattern instead.
 */
export async function purgeAllCaches(config) {
  const dir = path.join(path.dirname(config.cachePath), 'cache');
  let files = [];
  try {
    files = await fsp.readdir(dir);
  } catch {
    return { removed: 0 };
  }

  let removed = 0;
  for (const file of files) {
    // Both finished caches and any half-written temp file from an interrupted
    // flush. Leaving a .tmp behind would mean fragments of note text surviving
    // a sign-out, and would stop the directory being removed.
    const isCache = /^pages-[0-9a-f]{16}\.json$/.test(file);
    const isTemp = /^pages-[0-9a-f]{16}\.json\.[0-9]+\.[0-9a-f]+\.tmp$/.test(file);
    if (!isCache && !isTemp) continue;

    await fsp.rm(path.join(dir, file), { force: true, maxRetries: 3, retryDelay: 50 }).catch(
      () => {}
    );
    if (isCache) removed += 1;
  }

  // Best effort: on Windows the directory handle can linger a moment after the
  // files are gone. The files are what matter, so a failure here is not one.
  await fsp.rmdir(dir).catch(() => {});
  return { removed };
}
