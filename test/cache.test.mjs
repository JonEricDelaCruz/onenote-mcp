/**
 * Tests for the local page cache.
 *
 * The correctness bar here is higher than for most of this project, because a
 * cache bug is not a broken feature -- it is either a wrong answer presented
 * confidently, or someone's notes somewhere they did not agree to. So these
 * cover three things in order of severity:
 *
 *   1. It never serves text for a page that changed.
 *   2. It never lets one account read another account's notes.
 *   3. It never writes anything readable by other users, and never sends
 *      anything anywhere.
 */

import { test, describe, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fsp from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';

import { PageCache, createPageCache, purgeAllCaches } from '../src/cache.mjs';
import { OneNoteClient } from '../src/onenote.mjs';

let dir;

before(async () => {
  dir = await fsp.mkdtemp(path.join(os.tmpdir(), 'onenote-cache-test-'));
});

after(async () => {
  await fsp.rm(dir, { recursive: true, force: true });
});

const makeCache = (account = 'account-a', where = dir) =>
  new PageCache({ dir: where, accountId: async () => account });

describe('freshness', () => {
  test('a page returns its stored text when unchanged', async () => {
    const cache = makeCache();
    await cache.set('p1', { lastModified: '2026-08-01T10:00:00Z', text: 'hello notes' });

    assert.equal(await cache.get('p1', { lastModified: '2026-08-01T10:00:00Z' }), 'hello notes');
  });

  test('an edited page is a miss, not a stale hit', async () => {
    const cache = makeCache();
    await cache.set('p1', { lastModified: '2026-08-01T10:00:00Z', text: 'old text' });

    // This is the bug that would matter most: answering from a copy of a page
    // the user has since rewritten.
    assert.equal(await cache.get('p1', { lastModified: '2026-08-02T09:00:00Z' }), null);
  });

  test('an edited page is also evicted, not left to linger', async () => {
    const cache = makeCache();
    await cache.set('p1', { lastModified: 'A', text: 'old' });
    await cache.get('p1', { lastModified: 'B' });

    // Even asking with the original timestamp must not resurrect it.
    assert.equal(await cache.get('p1', { lastModified: 'A' }), null);
  });

  test('without a timestamp to check against, nothing is served', async () => {
    const cache = makeCache();
    await cache.set('p1', { lastModified: 'A', text: 'text' });

    assert.equal(await cache.get('p1', {}), null);
    assert.equal(await cache.get('p1', { lastModified: null }), null);
  });

  test('a page stored without a timestamp is not stored at all', async () => {
    const cache = makeCache();
    await cache.set('p2', { text: 'unverifiable' });

    const stats = await cache.stats();
    assert.equal(stats.entries, 0, 'unverifiable text must never enter the cache');
  });

  test('variants do not bleed into one another', async () => {
    const cache = makeCache();
    await cache.set('p1', { lastModified: 'A', variant: 'text:0:1', text: 'with attachments' });
    await cache.set('p1', { lastModified: 'A', variant: 'text:0:0', text: 'without' });

    assert.equal(await cache.get('p1', { lastModified: 'A', variant: 'text:0:1' }), 'with attachments');
    assert.equal(await cache.get('p1', { lastModified: 'A', variant: 'text:0:0' }), 'without');
  });
});

describe('account isolation', () => {
  test('two accounts on one machine cannot read each other', async () => {
    const a = makeCache('account-a');
    await a.set('shared-page-id', { lastModified: 'A', text: "account A's private note" });
    await a.flush();

    const b = makeCache('account-b');
    const leaked = await b.get('shared-page-id', { lastModified: 'A' });

    assert.equal(leaked, null, 'one account must never see another account’s notes');
  });

  test('the filename does not disclose the account', async () => {
    const cache = makeCache('jon.eric@example.com');
    await cache.set('p1', { lastModified: 'A', text: 'x' });
    await cache.flush();

    const files = await fsp.readdir(dir);
    const mine = files.filter((f) => f.startsWith('pages-'));

    assert.ok(mine.length > 0);
    for (const file of mine) {
      assert.doesNotMatch(file, /jon|eric|example|@/i, 'filename leaks the identity');
      assert.match(file, /^pages-[0-9a-f]{16}\.json$/);
    }
  });

  test('not being signed in means nothing is written', async () => {
    const scratch = await fsp.mkdtemp(path.join(os.tmpdir(), 'onenote-anon-'));
    const cache = new PageCache({ dir: scratch, accountId: async () => null });

    await cache.set('p1', { lastModified: 'A', text: 'x' });
    await cache.flush();

    const files = await fsp.readdir(scratch).catch(() => []);
    assert.equal(files.length, 0, 'no account means no file on disk');
    await fsp.rm(scratch, { recursive: true, force: true });
  });
});

describe('what lands on disk', () => {
  test('the file is readable only by its owner', async (t) => {
    if (process.platform === 'win32') return t.skip('POSIX permissions');

    const scratch = await fsp.mkdtemp(path.join(os.tmpdir(), 'onenote-perm-'));
    const cache = new PageCache({ dir: path.join(scratch, 'cache'), accountId: async () => 'a' });
    await cache.set('p1', { lastModified: 'A', text: 'private' });
    await cache.flush();

    const { path: file } = await cache.stats();
    const fileMode = (await fsp.stat(file)).mode & 0o777;
    const dirMode = (await fsp.stat(path.dirname(file))).mode & 0o777;

    assert.equal(fileMode, 0o600, 'note text must not be world-readable');
    assert.equal(dirMode, 0o700);
    await fsp.rm(scratch, { recursive: true, force: true });
  });

  test('the cache module contains no network code whatsoever', async () => {
    // The strongest claim in PRIVACY.md is that this file cannot phone home.
    // That claim should fail a test, not just a code review, if it stops
    // being true.
    const source = await fsp.readFile(new URL('../src/cache.mjs', import.meta.url), 'utf8');

    for (const forbidden of ['fetch(', 'http://', 'https://', 'net.', 'XMLHttpRequest', 'WebSocket']) {
      assert.ok(
        !source.includes(forbidden),
        `src/cache.mjs must contain no network capability, found: ${forbidden}`
      );
    }
  });

  test('survives a corrupt or truncated file instead of failing the read', async () => {
    const scratch = await fsp.mkdtemp(path.join(os.tmpdir(), 'onenote-corrupt-'));
    const cache = new PageCache({ dir: scratch, accountId: async () => 'a' });

    await cache.set('p1', { lastModified: 'A', text: 'x' });
    await cache.flush();
    const { path: file } = await cache.stats();
    await fsp.writeFile(file, '{ this is not json');

    const fresh = new PageCache({ dir: scratch, accountId: async () => 'a' });
    assert.equal(await fresh.get('p1', { lastModified: 'A' }), null, 'a bad file is a miss');
    await fresh.set('p1', { lastModified: 'A', text: 'recovered' });
    assert.equal(await fresh.get('p1', { lastModified: 'A' }), 'recovered');

    await fsp.rm(scratch, { recursive: true, force: true });
  });

  test('a file from a future format version is ignored, not misread', async () => {
    const scratch = await fsp.mkdtemp(path.join(os.tmpdir(), 'onenote-fmt-'));
    const cache = new PageCache({ dir: scratch, accountId: async () => 'a' });
    await cache.set('p1', { lastModified: 'A', text: 'x' });
    await cache.flush();

    const { path: file } = await cache.stats();
    const parsed = JSON.parse(await fsp.readFile(file, 'utf8'));
    await fsp.writeFile(file, JSON.stringify({ ...parsed, format: 99 }));

    const fresh = new PageCache({ dir: scratch, accountId: async () => 'a' });
    assert.equal(await fresh.get('p1', { lastModified: 'A' }), null);
    await fsp.rm(scratch, { recursive: true, force: true });
  });
});

describe('bounded growth', () => {
  test('an oversized page is skipped rather than allowed to dominate', async () => {
    const cache = makeCache();
    await cache.set('huge', { lastModified: 'A', text: 'x'.repeat(300_000) });

    assert.equal(await cache.get('huge', { lastModified: 'A' }), null);
  });

  test('the entry count stays capped, evicting least recently used', async () => {
    const scratch = await fsp.mkdtemp(path.join(os.tmpdir(), 'onenote-evict-'));
    const cache = new PageCache({ dir: scratch, accountId: async () => 'a' });

    for (let i = 0; i < 600; i += 1) {
      await cache.set(`p${i}`, { lastModified: 'A', text: `page ${i}` });
    }

    const stats = await cache.stats();
    assert.ok(stats.entries <= 500, `expected <= 500 entries, got ${stats.entries}`);
    assert.equal(await cache.get('p599', { lastModified: 'A' }), 'page 599', 'newest survives');
    assert.equal(await cache.get('p0', { lastModified: 'A' }), null, 'oldest evicted');

    await fsp.rm(scratch, { recursive: true, force: true });
  });
});

describe('erasing it', () => {
  test('clear empties the cache and removes the file', async () => {
    const scratch = await fsp.mkdtemp(path.join(os.tmpdir(), 'onenote-clear-'));
    const cache = new PageCache({ dir: scratch, accountId: async () => 'a' });
    await cache.set('p1', { lastModified: 'A', text: 'x' });
    await cache.flush();

    const { path: file } = await cache.stats();
    const { removed } = await cache.clear();

    assert.equal(removed, 1);
    assert.equal(await cache.get('p1', { lastModified: 'A' }), null);
    await assert.rejects(fsp.stat(file), 'the file itself must be gone');
    await fsp.rm(scratch, { recursive: true, force: true });
  });

  test('signing out purges every account on the machine', async () => {
    const scratch = await fsp.mkdtemp(path.join(os.tmpdir(), 'onenote-purge-'));
    const cacheDir = path.join(scratch, 'cache');
    const config = { cachePath: path.join(scratch, 'token-cache.json'), cacheMode: 'disk' };

    for (const account of ['a', 'b', 'c']) {
      const cache = new PageCache({ dir: cacheDir, accountId: async () => account });
      await cache.set('p1', { lastModified: 'A', text: 'x' });
      await cache.flush();
    }

    const { removed } = await purgeAllCaches(config);
    assert.equal(removed, 3, 'every account’s cache must go, not just the current one');
    await assert.rejects(fsp.stat(cacheDir), 'the cache directory should be gone too');

    await fsp.rm(scratch, { recursive: true, force: true });
  });

  test('purging when nothing was ever cached is not an error', async () => {
    const scratch = await fsp.mkdtemp(path.join(os.tmpdir(), 'onenote-nothing-'));
    const result = await purgeAllCaches({ cachePath: path.join(scratch, 'token-cache.json') });

    assert.equal(result.removed, 0);
    await fsp.rm(scratch, { recursive: true, force: true });
  });
});

describe('the off switches actually switch it off', () => {
  const accountId = async () => 'a';

  test('memory mode writes nothing to disk', async () => {
    const scratch = await fsp.mkdtemp(path.join(os.tmpdir(), 'onenote-mem-'));
    const config = { cacheMode: 'memory', cachePath: path.join(scratch, 'token-cache.json') };
    const cache = createPageCache(config, { accountId });

    await cache.set('p1', { lastModified: 'A', text: 'secret' });
    await cache.flush();

    assert.equal(cache.enabled, false);
    assert.equal(await cache.get('p1', { lastModified: 'A' }), null);
    await assert.rejects(fsp.stat(path.join(scratch, 'cache')), 'nothing may be written');
    await fsp.rm(scratch, { recursive: true, force: true });
  });

  test('off mode writes nothing to disk', async () => {
    const scratch = await fsp.mkdtemp(path.join(os.tmpdir(), 'onenote-off-'));
    const config = { cacheMode: 'off', cachePath: path.join(scratch, 'token-cache.json') };
    const cache = createPageCache(config, { accountId });

    await cache.set('p1', { lastModified: 'A', text: 'secret' });
    await cache.flush();

    assert.equal(await cache.get('p1', { lastModified: 'A' }), null);
    await assert.rejects(fsp.stat(path.join(scratch, 'cache')), 'nothing may be written');
    await fsp.rm(scratch, { recursive: true, force: true });
  });

  test('a client built without config caches nothing', async () => {
    // Guards against a partially-constructed client silently writing notes to
    // disk because a config field happened to be undefined.
    const cache = createPageCache({}, { accountId });
    assert.equal(cache.enabled, false);
  });

  test('disk mode is the default and does write', async () => {
    const scratch = await fsp.mkdtemp(path.join(os.tmpdir(), 'onenote-disk-'));
    const config = { cacheMode: 'disk', cachePath: path.join(scratch, 'token-cache.json') };
    const cache = createPageCache(config, { accountId });

    await cache.set('p1', { lastModified: 'A', text: 'remembered' });
    await cache.flush();

    assert.equal(await cache.get('p1', { lastModified: 'A' }), 'remembered');
    await fsp.rm(scratch, { recursive: true, force: true });
  });
});

describe('persistence across sessions', () => {
  test('a new process reads what the previous one stored', async () => {
    const scratch = await fsp.mkdtemp(path.join(os.tmpdir(), 'onenote-persist-'));

    const first = new PageCache({ dir: scratch, accountId: async () => 'a' });
    await first.set('p1', { lastModified: 'A', text: 'written last session' });
    await first.flush();

    // Whole point of the feature: quitting the app must not throw the work away.
    const second = new PageCache({ dir: scratch, accountId: async () => 'a' });
    assert.equal(await second.get('p1', { lastModified: 'A' }), 'written last session');

    await fsp.rm(scratch, { recursive: true, force: true });
  });
});

describe('end to end: the client really does skip the work', () => {
  /** A client whose Graph calls are counted, backed by a real temp cache dir. */
  const cachingClient = (scratch) => {
    const calls = { content: 0 };

    globalThis.fetch = async (url) => {
      if (String(url).includes('/content')) calls.content += 1;
      return {
        ok: true,
        status: 200,
        headers: { get: () => 'text/html' },
        text: async () => '<body><h1>Budget</h1><p>The number is 42.</p></body>',
        json: async () => ({ value: [] })
      };
    };

    const client = new OneNoteClient(
      { getAccessToken: async () => 'token', getAccount: async () => ({ homeAccountId: 'acct-1' }) },
      {
        graphBaseUrl: 'https://g/v1.0',
        cacheMode: 'disk',
        cachePath: path.join(scratch, 'token-cache.json')
      },
      () => {}
    );

    return { client, calls };
  };

  test('a later session reads the page without downloading it again', async (t) => {
    const real = globalThis.fetch;
    const scratch = await fsp.mkdtemp(path.join(os.tmpdir(), 'onenote-e2e-'));
    t.after(async () => {
      globalThis.fetch = real;
      await fsp.rm(scratch, { recursive: true, force: true });
    });

    const first = cachingClient(scratch);
    const a = await first.client.getPageContent('p1', {
      lastModified: '2026-08-01T10:00:00Z',
      readAttachments: false
    });
    await first.client.flushCache();

    assert.match(a.content, /42/);
    assert.equal(first.calls.content, 1);

    // A brand new client, as if the AI app had been quit and reopened. Its
    // 60-second memory cache is empty, so anything it returns came off disk.
    const second = cachingClient(scratch);
    const b = await second.client.getPageContent('p1', {
      lastModified: '2026-08-01T10:00:00Z',
      readAttachments: false
    });

    assert.equal(b.content, a.content, 'same text, without paying for it again');
    assert.equal(b.cached, true);
    assert.equal(second.calls.content, 0, 'no download should have happened at all');
  });

  test('editing the page forces a fresh download', async (t) => {
    const real = globalThis.fetch;
    const scratch = await fsp.mkdtemp(path.join(os.tmpdir(), 'onenote-e2e-stale-'));
    t.after(async () => {
      globalThis.fetch = real;
      await fsp.rm(scratch, { recursive: true, force: true });
    });

    const first = cachingClient(scratch);
    await first.client.getPageContent('p1', {
      lastModified: '2026-08-01T10:00:00Z',
      readAttachments: false
    });
    await first.client.flushCache();

    const second = cachingClient(scratch);
    await second.client.getPageContent('p1', {
      lastModified: '2026-08-05T12:00:00Z', // user edited it since
      readAttachments: false
    });

    assert.equal(second.calls.content, 1, 'an edited page must be re-fetched');
  });

  test('html reads are never persisted', async (t) => {
    const real = globalThis.fetch;
    const scratch = await fsp.mkdtemp(path.join(os.tmpdir(), 'onenote-e2e-html-'));
    t.after(async () => {
      globalThis.fetch = real;
      await fsp.rm(scratch, { recursive: true, force: true });
    });

    const { client } = cachingClient(scratch);
    await client.getPageContent('p1', { format: 'html', lastModified: 'A' });
    await client.flushCache();

    const stats = await client.pageCache.stats();
    assert.equal(stats.entries, 0, 'raw HTML is large and re-read by nobody');
  });
});
