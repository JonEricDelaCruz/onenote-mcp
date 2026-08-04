/**
 * Content-search performance behaviour.
 *
 * OneNote has no server-side full-text search, so matching page bodies means
 * fetching pages. The first implementation awaited each one in turn: ~50
 * sequential round-trips, 15-25 seconds of waiting. These tests pin the three
 * properties that fixed it, because a regression would be invisible in output
 * and only show up as "the tool feels slow".
 */

import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { OneNoteClient } from '../src/onenote.mjs';

const LATENCY_MS = 40;

/**
 * Client with the network faked out, so we can observe concurrency and counts
 * without touching Microsoft.
 */
function makeClient({ pages, matchOn = () => false, latency = LATENCY_MS } = {}) {
  const client = new OneNoteClient(
    { getAccessToken: async () => 'token', getAccount: async () => ({ username: 'x' }) },
    { graphBaseUrl: 'https://graph.invalid/v1.0', allowWrite: true },
    () => {}
  );

  const state = { fetches: 0, concurrent: 0, peakConcurrent: 0 };

  client.listPages = async () => pages;

  client.getPageContent = async (pageId) => {
    state.fetches += 1;
    state.concurrent += 1;
    state.peakConcurrent = Math.max(state.peakConcurrent, state.concurrent);
    try {
      await new Promise((resolve) => setTimeout(resolve, latency));
      return { format: 'text', content: matchOn(pageId) ? 'contains the needle here' : 'nothing' };
    } finally {
      state.concurrent -= 1;
    }
  };

  return { client, state };
}

const makePages = (n) =>
  Array.from({ length: n }, (_, i) => ({
    id: `page-${i}`,
    title: `Page ${i}`,
    lastModifiedDateTime: new Date(Date.now() - i * 60_000).toISOString()
  }));

describe('content search concurrency', () => {
  test('fetches pages in parallel, not one at a time', async () => {
    const { client, state } = makeClient({ pages: makePages(24) });

    const started = Date.now();
    await client.searchPages('needle', { searchContent: true, scanLimit: 24, limit: 25 });
    const elapsed = Date.now() - started;

    assert.ok(state.peakConcurrent > 1, 'requests must overlap');
    assert.ok(
      elapsed < 24 * LATENCY_MS * 0.6,
      `expected parallel speedup, took ${elapsed}ms vs ${24 * LATENCY_MS}ms sequential`
    );
  });

  test('keeps concurrency bounded, to stay under OneNote rate limits', async () => {
    const { client, state } = makeClient({ pages: makePages(60) });
    await client.searchPages('needle', { searchContent: true, scanLimit: 60, limit: 25 });

    assert.ok(
      state.peakConcurrent <= 8,
      `concurrency should stay small, peaked at ${state.peakConcurrent}`
    );
  });

  test('stops early once enough matches are found', async () => {
    // Every page matches, so the limit should be hit almost immediately.
    const { client, state } = makeClient({ pages: makePages(100), matchOn: () => true });

    const results = await client.searchPages('needle', {
      searchContent: true,
      scanLimit: 100,
      limit: 3
    });

    assert.equal(results.length, 3);
    assert.ok(
      state.fetches < 30,
      `should not scan everything after satisfying the limit; fetched ${state.fetches}`
    );
  });

  test('honours scanLimit as an upper bound', async () => {
    const { client, state } = makeClient({ pages: makePages(100) });
    await client.searchPages('needle', { searchContent: true, scanLimit: 10, limit: 25 });
    assert.ok(state.fetches <= 10, `fetched ${state.fetches}, expected at most 10`);
  });

  test('reports how much work it did', async () => {
    const { client } = makeClient({ pages: makePages(12) });
    const results = await client.searchPages('needle', {
      searchContent: true,
      scanLimit: 12,
      limit: 25
    });

    assert.ok(results._stats, 'stats should be attached so the cost is visible');
    assert.equal(results._stats.scanned, 12);
    assert.ok(results._stats.elapsedMs >= 0);
  });

  test('a failing page does not abort the whole search', async () => {
    const { client } = makeClient({ pages: makePages(10), matchOn: (id) => id === 'page-7' });
    const original = client.getPageContent;
    client.getPageContent = async (id) => {
      if (id === 'page-3') throw new Error('boom');
      return original(id);
    };

    const results = await client.searchPages('needle', {
      searchContent: true,
      scanLimit: 10,
      limit: 25
    });

    assert.equal(results.length, 1);
    assert.equal(results[0].id, 'page-7');
  });

  test('title search makes no page fetches at all', async () => {
    const pages = makePages(50);
    pages[4].title = 'Budget planning';
    const { client, state } = makeClient({ pages });

    const results = await client.searchPages('budget', { searchContent: false });

    assert.equal(state.fetches, 0, 'title search must not download page bodies');
    assert.equal(results.length, 1);
    assert.equal(results[0].matchedIn, 'title');
  });

  test('results come back newest first', async () => {
    const { client } = makeClient({ pages: makePages(12), matchOn: () => true });
    const results = await client.searchPages('needle', {
      searchContent: true,
      scanLimit: 12,
      limit: 12
    });

    const times = results.map((r) => new Date(r.lastModifiedDateTime).getTime());
    assert.deepEqual(times, [...times].sort((a, b) => b - a), 'pool finishes out of order');
  });
});

describe('content cache', () => {
  /**
   * Exercises the real caching path by faking only the HTTP boundary, and
   * restoring it afterwards so no other test is affected.
   */
  test('a repeated read does not hit the network twice', async (t) => {
    const realFetch = globalThis.fetch;
    t.after(() => {
      globalThis.fetch = realFetch;
    });

    let networkCalls = 0;
    globalThis.fetch = async () => {
      networkCalls += 1;
      return {
        ok: true,
        status: 200,
        headers: { get: () => 'text/html' },
        text: async () => '<html><body><p>hello world</p></body></html>'
      };
    };

    const client = new OneNoteClient(
      { getAccessToken: async () => 'token' },
      { graphBaseUrl: 'https://graph.invalid/v1.0', allowWrite: true },
      () => {}
    );

    const first = await client.getPageContent('p1', { format: 'text' });
    const second = await client.getPageContent('p1', { format: 'text' });

    assert.match(first.content, /hello world/);
    assert.notEqual(first.cached, true, 'first read is a real fetch');
    assert.equal(second.cached, true, 'second read should come from cache');
    assert.equal(networkCalls, 1, 'expected exactly one network fetch');
  });

  test('a different page is not served from another page cache entry', async (t) => {
    const realFetch = globalThis.fetch;
    t.after(() => {
      globalThis.fetch = realFetch;
    });

    const bodies = { p1: 'first page', p2: 'second page' };
    globalThis.fetch = async (url) => {
      const id = url.includes('p2') ? 'p2' : 'p1';
      return {
        ok: true,
        status: 200,
        headers: { get: () => 'text/html' },
        text: async () => `<html><body><p>${bodies[id]}</p></body></html>`
      };
    };

    const client = new OneNoteClient(
      { getAccessToken: async () => 'token' },
      { graphBaseUrl: 'https://graph.invalid/v1.0', allowWrite: true },
      () => {}
    );

    assert.match((await client.getPageContent('p1')).content, /first page/);
    assert.match((await client.getPageContent('p2')).content, /second page/);
  });
});
