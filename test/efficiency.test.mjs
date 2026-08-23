/**
 * Efficiency guards.
 *
 * Every regression this file catches is invisible in output: the tool still
 * works, it just costs more time, bandwidth, or tokens than it should. Those
 * are exactly the regressions that survive code review.
 */

import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { listTools as rpcListTools } from './rpc-client.mjs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { OneNoteClient } from '../src/onenote.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SERVER = path.join(__dirname, '..', 'onenote-mcp.mjs');

const META = {
  'io.modelcontextprotocol/protocolVersion': '2026-07-28',
  'io.modelcontextprotocol/clientInfo': { name: 't', version: '1' },
  'io.modelcontextprotocol/clientCapabilities': {}
};

function listTools() {
  return rpcListTools(SERVER, {
    ...process.env,
    ONENOTE_CLIENT_ID: '00000000-0000-0000-0000-000000000001',
    ONENOTE_SKIP_DOTENV: '1'
  });
}

/** Fake transport that counts requests by kind. */
function countingClient() {
  const calls = { content: 0, resources: 0, pages: 0, other: 0 };

  globalThis.fetch = async (url) => {
    if (url.includes('/content')) calls.content += 1;
    else if (url.includes('/resources/')) calls.resources += 1;
    else if (url.includes('/pages')) calls.pages += 1;
    else calls.other += 1;

    return {
      ok: true,
      status: 200,
      headers: { get: (h) => (h === 'content-type' ? 'text/html' : null) },
      text: async () =>
        '<body><p>Some text</p>' +
        '<img src="https://g/v1.0/me/onenote/resources/1/$value" />' +
        '<img src="https://g/v1.0/me/onenote/resources/2/$value" />' +
        '<img src="https://g/v1.0/me/onenote/resources/3/$value" /></body>',
      arrayBuffer: async () =>
        new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]).buffer,
      json: async () => ({ value: [] })
    };
  };

  const client = new OneNoteClient(
    { getAccessToken: async () => 'token' },
    { graphBaseUrl: 'https://g/v1.0', allowWrite: true },
    () => {}
  );
  return { client, calls };
}

describe('network efficiency', () => {
  test('reading a page with images downloads the page once, not twice', async (t) => {
    const real = globalThis.fetch;
    t.after(() => {
      globalThis.fetch = real;
    });

    // getPage reads text and then images. Before the HTML was cached, each half
    // fetched /content independently, doubling latency and bandwidth on the
    // single most common operation.
    const { client, calls } = countingClient();
    await client.getPageContent('p1', { readAttachments: false });
    await client.getPageImages('p1');

    assert.equal(calls.content, 1, `expected 1 /content request, made ${calls.content}`);
  });

  test('repeat reads inside the cache window make no request at all', async (t) => {
    const real = globalThis.fetch;
    t.after(() => {
      globalThis.fetch = real;
    });

    const { client, calls } = countingClient();
    await client.getPageContent('p1', { readAttachments: false });
    await client.getPageContent('p1', { readAttachments: false });
    await client.getPageContent('p1', { readAttachments: false });

    assert.equal(calls.content, 1, 'three reads of one page should hit the network once');
  });

  test('image count is capped, so an image-heavy page cannot flood the context', async (t) => {
    const real = globalThis.fetch;
    t.after(() => {
      globalThis.fetch = real;
    });

    const { client, calls } = countingClient();
    const { images, total } = await client.getPageImages('p1');

    assert.equal(total, 3, 'should report every image found');
    assert.ok(images.length <= 2, `default cap should be 2, attached ${images.length}`);
    assert.ok(calls.resources <= 2, 'must not download images it will not send');
  });
});

describe('token efficiency', () => {
  test('the tool schema stays within budget', async () => {
    const tools = await listTools();
    const size = JSON.stringify(tools).length;
    // Paid once per session, on every conversation.
    assert.ok(size < 14000, `tool schemas grew to ${size} chars; trim descriptions`);
  });

  test('the tool surface stays small', async () => {
    const tools = await listTools();
    assert.ok(tools.length <= 13, `${tools.length} tools; each one costs schema tokens`);
  });

  test('expensive defaults are conservative', async () => {
    const tools = await listTools();
    const getPage = tools.find((t) => t.name === 'getPage');

    // Images cost 1,300-4,000 tokens each, often more than the page text.
    assert.match(
      getPage.inputSchema.properties.maxImages.description,
      /defaults to 2/i,
      'the image default must stay small and be documented'
    );
    assert.match(
      getPage.inputSchema.properties.maxLength.description,
      /12000|defaults/i,
      'page text must have a documented cap'
    );
  });

  test('list results carry no bulky per-item fields', async () => {
    // webUrl is ~120 characters per page and was the single largest waste in
    // list output before it was dropped.
    const { readFileSync } = await import('node:fs');
    const src = readFileSync(SERVER, 'utf8');
    const lean = src.slice(src.indexOf('function leanPage'), src.indexOf('function fail'));

    assert.ok(!/webUrl/.test(lean), 'leanPage must not include webUrl');
    assert.ok(!/createdDateTime/.test(lean), 'leanPage must not include createdDateTime');
  });
});

describe('startup cost', () => {
  test('the server answers quickly from cold', async () => {
    const started = Date.now();
    await listTools();
    const elapsed = Date.now() - started;

    // Claude Desktop launches this on every app start; a slow boot is felt.
    assert.ok(elapsed < 5000, `cold start took ${elapsed}ms`);
  });
});

describe('images are offered, not charged for', () => {
  /**
   * Sending images on every read taxed the many questions that never needed
   * them: a screenshot costs more than the page it sits on. The policy is that
   * images are sent when a page is mostly picture, and otherwise announced so
   * the user can ask.
   */
  const THIN = 400;

  const decide = ({ mode, textLength, includeImages, format = 'text' }) => {
    const isText = format === 'text';
    const thin = textLength < THIN;
    const wants =
      isText &&
      (includeImages === true ||
        (includeImages !== false && (mode === 'always' || (mode === 'auto' && thin))));
    const announces = !wants && isText && includeImages !== false && mode !== 'never';
    return { wants, announces };
  };

  test('a page that is mostly picture sends its images', () => {
    // Its text cannot answer anything, so the image IS the content.
    assert.equal(decide({ mode: 'auto', textLength: 120 }).wants, true);
  });

  test('a page with real prose only announces them', () => {
    const r = decide({ mode: 'auto', textLength: 5000 });
    assert.equal(r.wants, false, 'should not spend tokens unprompted');
    assert.equal(r.announces, true, 'but must tell the reader they exist');
  });

  test('an explicit request always wins, even against the never setting', () => {
    assert.equal(decide({ mode: 'auto', textLength: 5000, includeImages: true }).wants, true);
    assert.equal(decide({ mode: 'never', textLength: 120, includeImages: true }).wants, true);
  });

  test('an explicit refusal always wins, and stays quiet', () => {
    const r = decide({ mode: 'auto', textLength: 120, includeImages: false });
    assert.equal(r.wants, false);
    assert.equal(r.announces, false, 'no nagging after the user said no');
  });

  test('the never setting neither sends nor announces', () => {
    const r = decide({ mode: 'never', textLength: 120 });
    assert.equal(r.wants, false);
    assert.equal(r.announces, false);
  });

  test('html format never carries images', () => {
    // Raw XHTML is for machine use; attaching pictures to it helps nobody.
    assert.equal(decide({ mode: 'always', textLength: 120, format: 'html' }).wants, false);
  });

  test('the image mode setting is exposed to users', async () => {
    const { readFileSync } = await import('node:fs');
    const manifest = JSON.parse(readFileSync(path.join(__dirname, '..', 'manifest.json'), 'utf8'));
    const setting = manifest.user_config.image_mode;

    assert.ok(setting, 'users need a way to change this without editing code');
    assert.equal(setting.default, 'auto');
    assert.match(setting.description, /auto|always|never/i);
  });

  test('counting images does not download them', async (t) => {
    const real = globalThis.fetch;
    t.after(() => {
      globalThis.fetch = real;
    });

    const { client, calls } = countingClient();
    await client.getPageContent('p1', { readAttachments: false });
    const { imageCount } = await client.countPageImages('p1');

    assert.equal(imageCount, 3);
    assert.equal(calls.resources, 0, 'counting must not fetch image bytes');
    assert.equal(calls.content, 1, 'and must reuse the cached page HTML');
  });
});

describe('portability guards', () => {
  /**
   * These are the mistakes that pass on macOS and Linux and fail only on
   * Windows, which means they are found by CI rather than by writing them --
   * an expensive way to learn. Catch them locally instead.
   */
  const sourceFiles = async () => {
    const { readdir, readFile } = await import('node:fs/promises');
    const root = path.join(__dirname, '..');
    const out = [];

    for (const dir of ['.', 'src', 'test', 'scripts']) {
      const full = path.join(root, dir);
      for (const name of await readdir(full)) {
        if (!name.endsWith('.mjs')) continue;
        out.push({ name: path.join(dir, name), text: await readFile(path.join(full, name), 'utf8') });
      }
    }
    return out;
  };

  test('no file URL is converted with .pathname', async () => {
    // `new URL(...).pathname` yields "/C:/Users/..." on Windows, which cannot
    // be spawned, imported, or opened. fileURLToPath() is the correct tool.
    for (const file of await sourceFiles()) {
      // Comments are stripped first: prose explaining the hazard should not
      // count as committing it.
      const code = file.text.replace(/\/\/.*$/gm, '');
      assert.ok(
        !/\)\s*\.pathname/.test(code),
        `${file.name} reads .pathname off a URL; use fileURLToPath() instead`
      );
    }
  });

  test('no test waits a fixed time for a spawned server', async () => {
    // Sleeping for a guessed number of milliseconds is a bet on how fast the
    // machine is. CI runners are slow, so the bet loses there and nowhere else.
    for (const file of await sourceFiles()) {
      if (!file.name.startsWith('test')) continue;
      assert.ok(
        !/setTimeout\(\(\) => child\.stdin\.end\(\)/.test(file.text),
        `${file.name} closes stdin on a timer; wait for the response instead`
      );
    }
  });
});
