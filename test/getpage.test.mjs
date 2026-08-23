/**
 * Regressions from real use, both found by actually reading a note rather than
 * by any unit test:
 *
 * 1. getPage returned the body only in `content` and metadata only in
 *    `structuredContent`. A client that surfaced structuredContent showed the
 *    model `{title, chars, truncated}`, so a 19,000-character page looked empty.
 *
 * 2. Listing pages across all sections fails on accounts with many sections;
 *    OneNote returns "maximum sections exceeded" instead of paginating.
 */

import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { converse as rpcConverse } from './rpc-client.mjs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { OneNoteClient } from '../src/onenote.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SERVER = path.join(__dirname, '..', 'onenote-mcp.mjs');

const META = {
  'io.modelcontextprotocol/protocolVersion': '2026-07-28',
  'io.modelcontextprotocol/clientInfo': { name: 'test', version: '1.0' },
  'io.modelcontextprotocol/clientCapabilities': {}
};

function converse(messages) {
  return rpcConverse(SERVER, messages, {
    env: {
      ...process.env,
      ONENOTE_CLIENT_ID: '00000000-0000-0000-0000-000000000001',
      ONENOTE_SKIP_DOTENV: '1',
      ONENOTE_TOKEN_CACHE: path.join(__dirname, '..', '.test-cache', 'tc.json')
    }
  }).then((r) => r.responses);
}

describe('getPage returns readable content', () => {
  test('does not hide the body behind structuredContent', async () => {
    const responses = await converse([
      { jsonrpc: '2.0', id: 1, method: 'tools/list', params: { _meta: META } }
    ]);
    const getPage = responses
      .find((r) => r.id === 1)
      .result.tools.find((t) => t.name === 'getPage');

    // The description must promise text, since text is now the only carrier.
    assert.match(getPage.description, /read/i);
    assert.ok(getPage.inputSchema.properties.page);
  });

  test('the handler emits body text and no metadata-only structuredContent', async () => {
    // Read the source rather than the wire, because reaching the wire needs a
    // real Microsoft account. The invariant is what matters: getPage must not
    // return a structuredContent object that omits the body.
    const { readFileSync } = await import('node:fs');
    const src = readFileSync(SERVER, 'utf8');
    const start = src.indexOf("'getPage'");
    const end = src.indexOf("'searchPages'", start);
    const block = src.slice(start, end);

    // Match the property assignment, not the comment explaining why it is absent.
    const codeOnly = block
      .split('\n')
      .filter((line) => !line.trim().startsWith('//'))
      .join('\n');

    assert.ok(
      !/structuredContent\s*:/.test(codeOnly),
      'getPage must not return structuredContent; the body lives in the text'
    );
    assert.ok(/type: 'text'/.test(codeOnly), 'getPage must return a text block');
    assert.ok(/body/.test(codeOnly), 'the page body must appear in the returned text');
  });
});

describe('accounts with many sections', () => {
  function clientWith({ globalFails }) {
    const client = new OneNoteClient(
      { getAccessToken: async () => 't' },
      { graphBaseUrl: 'https://graph.invalid/v1.0', allowWrite: true },
      () => {}
    );
    const calls = { global: 0, perSection: 0 };

    client.listSections = async () => [
      { id: 's1', displayName: 'Reading' },
      { id: 's2', displayName: 'Topics' }
    ];

    // Stand in for the private #collect by intercepting fetch.
    globalThis.fetch = async (url) => {
      if (url.includes('/me/onenote/pages')) {
        calls.global += 1;
        if (globalFails) {
          return {
            ok: false,
            status: 400,
            headers: { get: () => 'application/json' },
            text: async () =>
              JSON.stringify({
                error: { code: '30104', message: 'Maximum sections exceeded' }
              })
          };
        }
      } else {
        calls.perSection += 1;
      }
      return {
        ok: true,
        status: 200,
        headers: { get: () => 'application/json' },
        text: async () => '',
        json: async () => ({
          value: [
            {
              id: `p-${calls.perSection}`,
              title: `Page ${calls.perSection}`,
              lastModifiedDateTime: '2026-08-01T00:00:00Z'
            }
          ]
        })
      };
    };

    return { client, calls };
  }

  test('falls back to per-section reads when the global query is refused', async (t) => {
    const realFetch = globalThis.fetch;
    t.after(() => {
      globalThis.fetch = realFetch;
    });

    const { client, calls } = clientWith({ globalFails: true });
    const pages = await client.listPages({ limit: 10 });

    assert.ok(calls.global >= 1, 'should attempt the cheap global query first');
    assert.ok(calls.perSection >= 2, 'should then read each section');
    assert.ok(pages.length > 0, 'should still return pages');
    assert.ok(
      pages.every((p) => p.sectionName),
      'fallback results should carry their section name'
    );
  });

  test('uses the cheap path when the account allows it', async (t) => {
    const realFetch = globalThis.fetch;
    t.after(() => {
      globalThis.fetch = realFetch;
    });

    const { client, calls } = clientWith({ globalFails: false });
    await client.listPages({ limit: 10 });

    assert.equal(calls.perSection, 0, 'no per-section reads needed when global works');
  });
});
