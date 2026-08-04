/**
 * MCP client configuration files.
 *
 * Editing these by hand is where most people get stuck: the path is buried in a
 * per-OS application-support directory, the file may not exist yet, and a single
 * misplaced comma silently breaks every other server the user had configured.
 *
 * So we locate the file, merge into it rather than overwrite, and always take a
 * timestamped backup first.
 */

import fs from 'node:fs';
import fsp from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';

const home = os.homedir();

/** Per-OS config locations for the clients that use the standard mcpServers shape. */
export const CLIENTS = {
  'claude-desktop': {
    label: 'Claude Desktop',
    paths: {
      darwin: path.join(home, 'Library', 'Application Support', 'Claude', 'claude_desktop_config.json'),
      win32: path.join(process.env.APPDATA || path.join(home, 'AppData', 'Roaming'), 'Claude', 'claude_desktop_config.json'),
      linux: path.join(process.env.XDG_CONFIG_HOME || path.join(home, '.config'), 'Claude', 'claude_desktop_config.json')
    },
    restartHint: 'Quit Claude Desktop completely and reopen it.'
  },
  cursor: {
    label: 'Cursor',
    paths: {
      darwin: path.join(home, '.cursor', 'mcp.json'),
      win32: path.join(home, '.cursor', 'mcp.json'),
      linux: path.join(home, '.cursor', 'mcp.json')
    },
    restartHint: 'Restart Cursor.'
  },
  windsurf: {
    label: 'Windsurf',
    paths: {
      darwin: path.join(home, '.codeium', 'windsurf', 'mcp_config.json'),
      win32: path.join(home, '.codeium', 'windsurf', 'mcp_config.json'),
      linux: path.join(home, '.codeium', 'windsurf', 'mcp_config.json')
    },
    restartHint: 'Restart Windsurf.'
  }
};

export function configPathFor(clientKey, platform = process.platform) {
  return CLIENTS[clientKey]?.paths[platform] ?? null;
}

/** Which known clients appear to be installed, based on their config directory existing. */
export function detectInstalledClients(platform = process.platform) {
  const found = [];
  for (const [key, client] of Object.entries(CLIENTS)) {
    const configPath = client.paths[platform];
    if (!configPath) continue;
    // The config file itself may not exist yet even when the app does, so test
    // the containing directory.
    if (fs.existsSync(configPath) || fs.existsSync(path.dirname(configPath))) {
      found.push({ key, label: client.label, configPath, exists: fs.existsSync(configPath) });
    }
  }
  return found;
}

async function readConfig(configPath) {
  try {
    const raw = await fsp.readFile(configPath, 'utf8');
    if (!raw.trim()) return {};
    return JSON.parse(raw);
  } catch (error) {
    if (error.code === 'ENOENT') return {};
    if (error instanceof SyntaxError) {
      throw new Error(
        `${configPath} contains invalid JSON, so it cannot be edited safely.\n` +
          `Fix or remove the file, then try again. (${error.message})`
      );
    }
    throw error;
  }
}

/**
 * Add or update this server's entry in a client's config.
 *
 * @param {object} args
 * @param {string} args.configPath
 * @param {string} args.serverName Key under mcpServers.
 * @param {object} args.serverConfig `{ command, args, env }`
 * @param {boolean} [args.dryRun] Report what would change without writing.
 */
export async function writeServerEntry({ configPath, serverName, serverConfig, dryRun = false }) {
  const config = await readConfig(configPath);
  const existed = Boolean(config.mcpServers?.[serverName]);
  const otherServers = Object.keys(config.mcpServers ?? {}).filter((k) => k !== serverName);

  const updated = {
    ...config,
    mcpServers: {
      ...(config.mcpServers ?? {}),
      [serverName]: serverConfig
    }
  };

  const serialized = `${JSON.stringify(updated, null, 2)}\n`;

  if (dryRun) {
    return { dryRun: true, existed, otherServers, configPath, preview: serialized };
  }

  // Back up anything we are about to modify, so a bad merge is always recoverable.
  let backupPath = null;
  if (fs.existsSync(configPath)) {
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    backupPath = `${configPath}.backup-${stamp}`;
    await fsp.copyFile(configPath, backupPath);
  }

  await fsp.mkdir(path.dirname(configPath), { recursive: true });

  // Write to a temp file and rename, so an interrupted write cannot leave the
  // user with a truncated config and no working MCP servers at all.
  const tmp = `${configPath}.${process.pid}.tmp`;
  await fsp.writeFile(tmp, serialized, 'utf8');
  await fsp.rename(tmp, configPath);

  return { dryRun: false, existed, otherServers, configPath, backupPath };
}

/** Remove this server's entry, leaving every other server untouched. */
export async function removeServerEntry({ configPath, serverName }) {
  const config = await readConfig(configPath);
  if (!config.mcpServers?.[serverName]) return { removed: false, configPath };

  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  const backupPath = `${configPath}.backup-${stamp}`;
  await fsp.copyFile(configPath, backupPath);

  delete config.mcpServers[serverName];
  await fsp.writeFile(configPath, `${JSON.stringify(config, null, 2)}\n`, 'utf8');
  return { removed: true, configPath, backupPath };
}
