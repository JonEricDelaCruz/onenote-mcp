/**
 * Tests for editing MCP client configuration files.
 *
 * This code writes to files the user did not create and may have other servers
 * in. Clobbering a Claude Desktop config would break every other integration
 * they have, so the merge behaviour is worth pinning down precisely.
 */

import { test, describe, beforeEach, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, writeFileSync, readFileSync, readdirSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { writeServerEntry, removeServerEntry, CLIENTS, configPathFor } from '../src/clients.mjs';

let dir;
let configPath;

const ENTRY = { command: 'node', args: ['/path/to/onenote-mcp.mjs'], env: {} };

beforeEach(() => {
  dir = mkdtempSync(path.join(tmpdir(), 'onenote-clients-'));
  configPath = path.join(dir, 'config.json');
});

afterEach(() => {
  rmSync(dir, { recursive: true, force: true });
});

const read = () => JSON.parse(readFileSync(configPath, 'utf8'));

describe('writeServerEntry', () => {
  test('creates the file when it does not exist', async () => {
    const result = await writeServerEntry({ configPath, serverName: 'onenote', serverConfig: ENTRY });
    assert.equal(result.existed, false);
    assert.deepEqual(read().mcpServers.onenote, ENTRY);
  });

  test('creates missing parent directories', async () => {
    const nested = path.join(dir, 'a', 'b', 'config.json');
    await writeServerEntry({ configPath: nested, serverName: 'onenote', serverConfig: ENTRY });
    assert.ok(existsSync(nested));
  });

  test('preserves other MCP servers', async () => {
    writeFileSync(
      configPath,
      JSON.stringify({ mcpServers: { github: { command: 'gh-mcp' }, slack: { command: 'slack-mcp' } } })
    );

    const result = await writeServerEntry({ configPath, serverName: 'onenote', serverConfig: ENTRY });

    const config = read();
    assert.deepEqual(config.mcpServers.github, { command: 'gh-mcp' });
    assert.deepEqual(config.mcpServers.slack, { command: 'slack-mcp' });
    assert.deepEqual(config.mcpServers.onenote, ENTRY);
    assert.deepEqual(result.otherServers.sort(), ['github', 'slack']);
  });

  test('preserves unrelated top-level keys', async () => {
    writeFileSync(configPath, JSON.stringify({ theme: 'dark', globalShortcut: 'Cmd+K', mcpServers: {} }));
    await writeServerEntry({ configPath, serverName: 'onenote', serverConfig: ENTRY });

    const config = read();
    assert.equal(config.theme, 'dark');
    assert.equal(config.globalShortcut, 'Cmd+K');
  });

  test('updates an existing entry rather than duplicating it', async () => {
    writeFileSync(configPath, JSON.stringify({ mcpServers: { onenote: { command: 'old' } } }));
    const result = await writeServerEntry({ configPath, serverName: 'onenote', serverConfig: ENTRY });

    assert.equal(result.existed, true);
    assert.deepEqual(read().mcpServers.onenote, ENTRY);
    assert.equal(Object.keys(read().mcpServers).length, 1);
  });

  test('backs up an existing file before modifying it', async () => {
    writeFileSync(configPath, JSON.stringify({ mcpServers: { github: { command: 'gh' } } }));
    const result = await writeServerEntry({ configPath, serverName: 'onenote', serverConfig: ENTRY });

    assert.ok(result.backupPath, 'a backup path should be reported');
    const backup = JSON.parse(readFileSync(result.backupPath, 'utf8'));
    assert.deepEqual(backup.mcpServers, { github: { command: 'gh' } });
  });

  test('does not create a backup when there was no file', async () => {
    const result = await writeServerEntry({ configPath, serverName: 'onenote', serverConfig: ENTRY });
    assert.equal(result.backupPath, null);
  });

  test('dry run reports changes without touching disk', async () => {
    writeFileSync(configPath, JSON.stringify({ mcpServers: { github: { command: 'gh' } } }));
    const before = readFileSync(configPath, 'utf8');

    const result = await writeServerEntry({
      configPath,
      serverName: 'onenote',
      serverConfig: ENTRY,
      dryRun: true
    });

    assert.equal(result.dryRun, true);
    assert.match(result.preview, /onenote/);
    assert.equal(readFileSync(configPath, 'utf8'), before, 'file must be unchanged');
    assert.equal(readdirSync(dir).length, 1, 'no backup should be written on a dry run');
  });

  test('refuses to touch a file containing invalid JSON', async () => {
    writeFileSync(configPath, '{ this is not json');
    await assert.rejects(
      () => writeServerEntry({ configPath, serverName: 'onenote', serverConfig: ENTRY }),
      /invalid JSON/
    );
    assert.equal(readFileSync(configPath, 'utf8'), '{ this is not json', 'must be left alone');
  });

  test('treats an empty file as an empty config', async () => {
    writeFileSync(configPath, '   \n');
    await writeServerEntry({ configPath, serverName: 'onenote', serverConfig: ENTRY });
    assert.deepEqual(read().mcpServers.onenote, ENTRY);
  });

  test('writes valid, readable JSON', async () => {
    await writeServerEntry({ configPath, serverName: 'onenote', serverConfig: ENTRY });
    const raw = readFileSync(configPath, 'utf8');
    assert.doesNotThrow(() => JSON.parse(raw));
    assert.match(raw, /\n$/, 'should end with a newline');
    assert.match(raw, /\n {2}"mcpServers"/, 'should be indented for humans');
  });
});

describe('removeServerEntry', () => {
  test('removes only our entry', async () => {
    writeFileSync(
      configPath,
      JSON.stringify({ mcpServers: { onenote: ENTRY, github: { command: 'gh' } } })
    );

    const result = await removeServerEntry({ configPath, serverName: 'onenote' });
    assert.equal(result.removed, true);

    const config = read();
    assert.equal(config.mcpServers.onenote, undefined);
    assert.deepEqual(config.mcpServers.github, { command: 'gh' });
  });

  test('is a no-op when the entry is absent', async () => {
    writeFileSync(configPath, JSON.stringify({ mcpServers: { github: { command: 'gh' } } }));
    const result = await removeServerEntry({ configPath, serverName: 'onenote' });
    assert.equal(result.removed, false);
    assert.deepEqual(read().mcpServers.github, { command: 'gh' });
  });

  test('is a no-op when the file does not exist', async () => {
    const result = await removeServerEntry({ configPath, serverName: 'onenote' });
    assert.equal(result.removed, false);
  });
});

describe('client locations', () => {
  test('every known client has a path on every supported platform', () => {
    for (const [key, client] of Object.entries(CLIENTS)) {
      for (const platform of ['darwin', 'win32', 'linux']) {
        assert.ok(client.paths[platform], `${key} has no path for ${platform}`);
        assert.ok(path.isAbsolute(client.paths[platform]), `${key}/${platform} path is not absolute`);
      }
    }
  });

  test('Claude Desktop uses the documented per-OS locations', () => {
    // Assert on path SEGMENTS, not separators. These paths are built with
    // path.join, which emits backslashes when the test itself runs on Windows —
    // so a regex containing "/" passes on macOS and Linux and fails on Windows.
    // That is exactly how this suite went red on windows-latest while staying
    // green everywhere else.
    const segmentsOf = (p) => p.split(/[/\\]/);

    const mac = segmentsOf(configPathFor('claude-desktop', 'darwin'));
    assert.ok(mac.includes('Library'), 'macOS path should live under Library');
    assert.ok(mac.includes('Application Support'), 'macOS path should use Application Support');
    assert.ok(mac.includes('Claude'), 'macOS path should be inside a Claude folder');
    assert.equal(mac[mac.length - 1], 'claude_desktop_config.json');

    const win = segmentsOf(configPathFor('claude-desktop', 'win32'));
    assert.ok(win.includes('Claude'));
    assert.equal(win[win.length - 1], 'claude_desktop_config.json');

    const linux = segmentsOf(configPathFor('claude-desktop', 'linux'));
    assert.equal(linux[linux.length - 1], 'claude_desktop_config.json');
  });

  test('unknown clients resolve to null rather than throwing', () => {
    assert.equal(configPathFor('not-a-real-client'), null);
  });
});
