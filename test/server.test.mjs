/**
 * End-to-end protocol tests: spawn the real server over stdio and speak
 * JSON-RPC to it.
 *
 * These lock in the behaviours that were previously broken:
 *   - the server starts at all (v1 imported a local SDK checkout that did not
 *     exist, so `npm start` failed outright)
 *   - every tool exposes a real input schema (v1 exposed none, then read a
 *     `random_string` placeholder argument)
 *   - both the 2026-07-28 stateless era and the 2025-era `initialize` handshake
 *     are served from the same binary
 *   - a missing client ID fails fast with an explanatory message
 */

import { test, describe, after } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SERVER = path.join(__dirname, '..', 'onenote-mcp.mjs');

const MODERN_META = {
  'io.modelcontextprotocol/protocolVersion': '2026-07-28',
  'io.modelcontextprotocol/clientInfo': { name: 'test', version: '1.0' },
  'io.modelcontextprotocol/clientCapabilities': {}
};

const ENV = {
  ...process.env,
  // A syntactically valid GUID that is not the borrowed Microsoft one. No real
  // network call is made: these tests only exercise discovery, never sign-in.
  ONENOTE_CLIENT_ID: '00000000-0000-0000-0000-000000000001',
  ONENOTE_ALLOW_DEVICE_CODE: '',
  ONENOTE_TOKEN_CACHE: path.join(__dirname, '..', '.test-cache', 'token-cache.json'),
  // A developer's own .env must not change what these tests observe.
  ONENOTE_SKIP_DOTENV: '1'
};

/** Send a batch of JSON-RPC messages, collect stdout responses, then exit. */
function converse(messages, { env = ENV, timeoutMs = 15000 } = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [SERVER], {
      env,
      stdio: ['pipe', 'pipe', 'pipe']
    });

    let stdout = '';
    let stderr = '';
    const timer = setTimeout(() => {
      child.kill('SIGKILL');
      reject(new Error(`Timed out after ${timeoutMs}ms.\nstderr:\n${stderr}`));
    }, timeoutMs);

    child.stdout.on('data', (chunk) => {
      stdout += chunk;
    });
    child.stderr.on('data', (chunk) => {
      stderr += chunk;
    });

    child.on('error', reject);
    child.on('close', (code) => {
      clearTimeout(timer);
      const responses = stdout
        .split('\n')
        .filter((line) => line.trim())
        .map((line) => {
          try {
            return JSON.parse(line);
          } catch {
            return { unparsed: line };
          }
        });
      resolve({ responses, stderr, stdout, code });
    });

    for (const message of messages) {
      child.stdin.write(`${JSON.stringify(message)}\n`);
    }

    // Give the server a moment to answer, then close the input side.
    setTimeout(() => child.stdin.end(), 1200);
  });
}

after(async () => {
  const { rm } = await import('node:fs/promises');
  await rm(path.join(__dirname, '..', '.test-cache'), { recursive: true, force: true });
});

describe('startup', () => {
  test('behaves correctly when no application ID is supplied', async () => {
    // Once a maintainer fills in BUNDLED_CLIENT_ID, a bare launch must succeed
    // (that is the whole point of bundling it). Before that, it must fail with
    // instructions. Assert whichever applies so this test survives the switch.
    const { BUNDLED_CLIENT_ID } = await import('../src/config.mjs');
    const env = { ...ENV };
    delete env.ONENOTE_CLIENT_ID;

    const { stderr, code, responses } = await converse(
      [{ jsonrpc: '2.0', id: 1, method: 'tools/list', params: { _meta: MODERN_META } }],
      { env }
    );

    if (BUNDLED_CLIENT_ID.trim()) {
      assert.equal(code, 0, 'a bundled ID should let the server start with no configuration');
      assert.ok(
        responses.find((r) => r.id === 1)?.result?.tools?.length,
        'server should serve tools using the bundled ID'
      );
    } else {
      assert.equal(code, 1, 'should exit non-zero');
      assert.match(stderr, /No Microsoft application ID is configured/);
      assert.match(stderr, /entra\.microsoft\.com/, 'should say where to register an app');
    }
  });

  test('rejects the borrowed Microsoft first-party client ID', async () => {
    const env = { ...ENV, ONENOTE_CLIENT_ID: '14d82eec-204b-4c2f-b7e8-296a70dab67e' };
    const { stderr, code } = await converse([], { env });
    assert.equal(code, 1);
    assert.match(stderr, /belongs to Microsoft/i);
  });

  test('rejects an application ID that is not a GUID', async () => {
    const env = { ...ENV, ONENOTE_CLIENT_ID: 'my-app-id' };
    const { stderr, code } = await converse([], { env });
    assert.equal(code, 1);
    assert.match(stderr, /does not look like a Microsoft application ID/);
  });

  test('never writes diagnostics to stdout', async () => {
    // stdout is the JSON-RPC channel; a stray console.log corrupts the stream.
    const { stdout } = await converse([
      { jsonrpc: '2.0', id: 1, method: 'tools/list', params: { _meta: MODERN_META } }
    ]);
    for (const line of stdout.split('\n').filter((l) => l.trim())) {
      assert.doesNotThrow(() => JSON.parse(line), `non-JSON on stdout: ${line}`);
    }
  });
});

describe('2026-07-28 stateless era', () => {
  test('answers tools/list with no handshake', async () => {
    const { responses } = await converse([
      { jsonrpc: '2.0', id: 1, method: 'tools/list', params: { _meta: MODERN_META } }
    ]);
    const result = responses.find((r) => r.id === 1)?.result;
    assert.ok(result, 'expected a result for tools/list');
    assert.ok(Array.isArray(result.tools) && result.tools.length > 0);
  });

  test('returns cache hints on list results (SEP-2549)', async () => {
    const { responses } = await converse([
      { jsonrpc: '2.0', id: 1, method: 'tools/list', params: { _meta: MODERN_META } }
    ]);
    const result = responses.find((r) => r.id === 1)?.result;
    assert.ok('ttlMs' in result, 'modern era should carry ttlMs');
    assert.ok('cacheScope' in result, 'modern era should carry cacheScope');
  });

  test('every tool declares an object input schema', async () => {
    const { responses } = await converse([
      { jsonrpc: '2.0', id: 1, method: 'tools/list', params: { _meta: MODERN_META } }
    ]);
    const { tools } = responses.find((r) => r.id === 1).result;

    for (const tool of tools) {
      assert.ok(tool.inputSchema, `${tool.name} has no inputSchema`);
      assert.equal(tool.inputSchema.type, 'object', `${tool.name} schema is not an object`);
      assert.ok(tool.description, `${tool.name} has no description`);
    }
  });

  test('no tool exposes the legacy random_string placeholder', async () => {
    const { responses } = await converse([
      { jsonrpc: '2.0', id: 1, method: 'tools/list', params: { _meta: MODERN_META } }
    ]);
    const { tools } = responses.find((r) => r.id === 1).result;
    for (const tool of tools) {
      const props = Object.keys(tool.inputSchema.properties ?? {});
      assert.ok(
        !props.includes('random_string'),
        `${tool.name} still exposes the random_string placeholder`
      );
    }
  });

  test('tools that need an ID actually require it', async () => {
    const { responses } = await converse([
      { jsonrpc: '2.0', id: 1, method: 'tools/list', params: { _meta: MODERN_META } }
    ]);
    const { tools } = responses.find((r) => r.id === 1).result;
    const byName = Object.fromEntries(tools.map((t) => [t.name, t]));

    assert.deepEqual(byName.getPage.inputSchema.required, ['pageId']);
    assert.ok(byName.searchPages.inputSchema.required.includes('query'));
    assert.ok(byName.createPage.inputSchema.required.includes('title'));
  });

  test('sections can be addressed by name, not just ID', async () => {
    // Requiring a GUID forced a discovery round-trip before every action, which
    // is what made the assistant feel slow and click-heavy.
    const { responses } = await converse([
      { jsonrpc: '2.0', id: 1, method: 'tools/list', params: { _meta: MODERN_META } }
    ]);
    const { tools } = responses.find((r) => r.id === 1).result;
    const byName = Object.fromEntries(tools.map((t) => [t.name, t]));

    for (const tool of ['listPages', 'createPage', 'searchPages']) {
      const props = byName[tool].inputSchema.properties;
      assert.ok(props.section, `${tool} should accept a section name`);
      assert.match(
        props.section.description,
        /name/i,
        `${tool}'s section argument should document name matching`
      );
    }

    // createPage must not demand a section at all, so a configured default works.
    assert.ok(
      !byName.createPage.inputSchema.required.includes('section'),
      'createPage should allow the default section to supply it'
    );

    assert.ok(byName.listSections.inputSchema.properties.notebook);
  });

  test('getOutline exists and needs no arguments', async () => {
    const { responses } = await converse([
      { jsonrpc: '2.0', id: 1, method: 'tools/list', params: { _meta: MODERN_META } }
    ]);
    const { tools } = responses.find((r) => r.id === 1).result;
    const outline = tools.find((t) => t.name === 'getOutline');

    assert.ok(outline, 'getOutline should be registered');
    assert.equal(outline.annotations?.readOnlyHint, true);
    assert.deepEqual(outline.inputSchema.required ?? [], []);
    assert.match(
      outline.description,
      /one call/i,
      'description should steer the model away from crawling'
    );
  });

  test('server instructions discourage hierarchy crawling', async () => {
    const { responses } = await converse([
      {
        jsonrpc: '2.0',
        id: 1,
        method: 'initialize',
        params: {
          protocolVersion: '2025-06-18',
          capabilities: {},
          clientInfo: { name: 'c', version: '1' }
        }
      }
    ]);
    const instructions = responses.find((r) => r.id === 1)?.result?.instructions ?? '';
    assert.match(instructions, /getOutline/);
    assert.match(instructions, /section group/i);
  });

  test('destructive tools are annotated as such', async () => {
    const { responses } = await converse([
      { jsonrpc: '2.0', id: 1, method: 'tools/list', params: { _meta: MODERN_META } }
    ]);
    const { tools } = responses.find((r) => r.id === 1).result;
    const byName = Object.fromEntries(tools.map((t) => [t.name, t]));

    assert.equal(byName.deletePage.annotations?.destructiveHint, true);
    assert.equal(byName.listNotebooks.annotations?.readOnlyHint, true);
    assert.equal(byName.getPage.annotations?.readOnlyHint, true);
    assert.equal(byName.getOutline.annotations?.readOnlyHint, true);
  });

  test('deletePage requires a title confirmation', async () => {
    const { responses } = await converse([
      { jsonrpc: '2.0', id: 1, method: 'tools/list', params: { _meta: MODERN_META } }
    ]);
    const { tools } = responses.find((r) => r.id === 1).result;
    const del = tools.find((t) => t.name === 'deletePage');
    assert.ok(
      del.inputSchema.required.includes('confirmTitle'),
      'deletePage must require confirmTitle as a guard'
    );
  });
});

describe('2025-era compatibility', () => {
  test('serves clients that open with initialize', async () => {
    const { responses } = await converse([
      {
        jsonrpc: '2.0',
        id: 1,
        method: 'initialize',
        params: {
          protocolVersion: '2025-06-18',
          capabilities: {},
          clientInfo: { name: 'legacy-client', version: '1.0' }
        }
      },
      { jsonrpc: '2.0', method: 'notifications/initialized' },
      { jsonrpc: '2.0', id: 2, method: 'tools/list' }
    ]);

    const init = responses.find((r) => r.id === 1)?.result;
    assert.ok(init, 'initialize should be answered');
    assert.equal(init.serverInfo.name, 'onenote');

    const list = responses.find((r) => r.id === 2)?.result;
    assert.ok(Array.isArray(list?.tools) && list.tools.length > 0);
  });
});

describe('unauthenticated behaviour', () => {
  test('data tools fail with actionable guidance, not a hang or a stack trace', async () => {
    const { responses } = await converse([
      {
        jsonrpc: '2.0',
        id: 1,
        method: 'tools/call',
        params: { name: 'listNotebooks', arguments: {}, _meta: MODERN_META }
      }
    ]);

    const result = responses.find((r) => r.id === 1)?.result;
    assert.ok(result, 'tool call should return a result, not a protocol error');
    assert.equal(result.isError, true);
    const text = result.content.map((c) => c.text).join('\n');
    assert.match(text, /authenticate|sign in/i, 'should tell the user how to fix it');
  });

  test('authStatus reports not-signed-in without prompting', async () => {
    const { responses } = await converse([
      {
        jsonrpc: '2.0',
        id: 1,
        method: 'tools/call',
        params: { name: 'authStatus', arguments: {}, _meta: MODERN_META }
      }
    ]);

    const result = responses.find((r) => r.id === 1)?.result;
    assert.ok(result, 'authStatus should always answer');
    assert.notEqual(result.isError, true, 'authStatus must not error when signed out');
    assert.equal(result.structuredContent.signedIn, false);
    assert.equal(result.structuredContent.cachedAccountPresent, false);
  });

  test('input validation rejects a bad argument type', async () => {
    const { responses } = await converse([
      {
        jsonrpc: '2.0',
        id: 1,
        method: 'tools/call',
        params: { name: 'getPage', arguments: { pageId: 12345 }, _meta: MODERN_META }
      }
    ]);

    const message = responses.find((r) => r.id === 1);
    const failed =
      message?.error !== undefined || message?.result?.isError === true;
    assert.ok(failed, 'a numeric pageId should be rejected by schema validation');
  });
});

describe('token efficiency', () => {
  test('no redundant tools are exposed', async () => {
    // Every tool costs schema tokens in every session. getNotebook and whoAmI
    // were fully covered by getOutline and authStatus respectively.
    const { responses } = await converse([
      { jsonrpc: '2.0', id: 1, method: 'tools/list', params: { _meta: MODERN_META } }
    ]);
    const names = responses.find((r) => r.id === 1).result.tools.map((t) => t.name);

    assert.ok(!names.includes('getNotebook'), 'getNotebook duplicates getOutline');
    assert.ok(!names.includes('whoAmI'), 'whoAmI duplicates authStatus');
    assert.ok(names.length <= 13, `tool surface should stay small, got ${names.length}`);
  });

  test('the whole tool schema stays within a sane budget', async () => {
    const { responses } = await converse([
      { jsonrpc: '2.0', id: 1, method: 'tools/list', params: { _meta: MODERN_META } }
    ]);
    const size = JSON.stringify(responses.find((r) => r.id === 1).result.tools).length;

    // Sent once per session. A regression here is easy to introduce by writing
    // an over-long description and impossible to notice by eye.
    assert.ok(size < 12000, `tool schemas grew to ${size} chars; trim descriptions`);
  });

  test('getPage caps content by default and says so', async () => {
    const { responses } = await converse([
      { jsonrpc: '2.0', id: 1, method: 'tools/list', params: { _meta: MODERN_META } }
    ]);
    const getPage = responses.find((r) => r.id === 1).result.tools.find((t) => t.name === 'getPage');

    assert.match(
      getPage.inputSchema.properties.maxLength.description,
      /default/i,
      'the cap should be documented so the model knows it can ask for more'
    );
  });
});
