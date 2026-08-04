#!/usr/bin/env node
/**
 * OneNote CLI.
 *
 * Replaces the ten near-identical helper scripts the project used to ship, each
 * of which re-implemented token reading and Graph setup by copy-paste. This
 * shares exactly the same auth and Graph modules as the MCP server, so if
 * `onenote-cli auth` works then the MCP server will authenticate too -- which
 * makes this the reliable way to sign in for MCP clients that cannot surface an
 * interactive prompt (danosb/onenote-mcp#1, #6).
 *
 * Usage:
 *   onenote-cli auth                       Sign in
 *   onenote-cli status                     Show auth + config state
 *   onenote-cli signout                    Clear cached credentials
 *   onenote-cli notebooks                  List notebooks
 *   onenote-cli sections [notebookId]      List sections
 *   onenote-cli pages [sectionId]          List pages
 *   onenote-cli read <pageId> [--html]     Print one page
 *   onenote-cli search <query> [--content] Search pages
 *   onenote-cli dump [--limit N]           Print every page's text
 *   onenote-cli create <sectionId> <title> [--file f | --text t]
 */

import { readFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

import { loadConfig, ConfigError, BUNDLED_CLIENT_ID } from './src/config.mjs';
import { OneNoteAuth, inspectCache } from './src/auth.mjs';
import { OneNoteClient } from './src/onenote.mjs';
import {
  CLIENTS,
  configPathFor,
  detectInstalledClients,
  writeServerEntry,
  removeServerEntry
} from './src/clients.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const { version } = JSON.parse(readFileSync(path.join(__dirname, 'package.json'), 'utf8'));

const USAGE = `onenote-cli v${version}

  setup                             Sign in and configure your AI app (start here)
  doctor                            Diagnose problems and print exact fixes
  uninstall                         Remove this server from your AI app's config

  auth                              Sign in to Microsoft
  status                            Show authentication and configuration state
  signout                           Remove cached credentials
  whoami                            Show the signed-in account

  notebooks                         List notebooks
  sections [notebookId]             List sections (all, or in one notebook)
  pages [sectionId]                 List pages (all, or in one section)
  read <pageId> [--html]            Print a page
  search <query> [--content]        Search page titles, optionally bodies
  dump [--limit N]                  Print the text of every page
  create <sectionId> <title>        Create a page
        [--text "..." | --file path]

Options:
  --limit N       Cap results (default 50)
  --json          Emit JSON instead of formatted text
  --client NAME   Target a specific app: claude-desktop, cursor, windsurf
  --dry-run       Show what setup would change, without changing it
  --help          Show this message

Configuration comes from environment variables or .env -- see .env.example.
`;

/** ANSI styling, suppressed when not a TTY or when NO_COLOR is set. */
const useColor = process.stdout.isTTY && !process.env.NO_COLOR;
const style = {
  bold: (s) => (useColor ? `[1m${s}[0m` : s),
  dim: (s) => (useColor ? `[2m${s}[0m` : s),
  green: (s) => (useColor ? `[32m${s}[0m` : s),
  yellow: (s) => (useColor ? `[33m${s}[0m` : s),
  red: (s) => (useColor ? `[31m${s}[0m` : s)
};

const PASS = () => style.green('  ok  ');
const WARN = () => style.yellow(' warn ');
const FAIL = () => style.red(' fail ');

// --------------------------------------------------------------------- helpers

function parseArgs(argv) {
  const positional = [];
  const flags = {};

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg.startsWith('--')) {
      positional.push(arg);
      continue;
    }
    const key = arg.slice(2);
    const next = argv[i + 1];
    if (next !== undefined && !next.startsWith('--')) {
      flags[key] = next;
      i += 1;
    } else {
      flags[key] = true;
    }
  }

  return { positional, flags };
}

function out(value) {
  process.stdout.write(`${value}\n`);
}

function emit(flags, data, render) {
  if (flags.json) {
    out(JSON.stringify(data, null, 2));
  } else {
    render();
  }
}

function limitOf(flags, fallback = 50) {
  const n = Number(flags.limit);
  return Number.isFinite(n) && n > 0 ? Math.min(n, 200) : fallback;
}

// ---------------------------------------------------------------------- main

async function main() {
  const { positional, flags } = parseArgs(process.argv.slice(2));
  const command = positional[0];

  if (!command || flags.help || command === 'help') {
    out(USAGE);
    return 0;
  }

  await loadDotEnvIfPresent();

  // setup and doctor must run even when configuration is broken -- diagnosing
  // and repairing that state is precisely their job.
  if (command === 'setup') return runSetup(flags);
  if (command === 'doctor') return runDoctor(flags);
  if (command === 'uninstall') return runUninstall(flags);

  let config;
  try {
    config = loadConfig();
  } catch (error) {
    if (error instanceof ConfigError) {
      process.stderr.write(`\n${error.message}\n\n`);
      process.stderr.write(style.dim('Run `onenote-cli doctor` for a full diagnosis.\n\n'));
      return 1;
    }
    throw error;
  }

  const log = (msg) => process.stderr.write(`${msg}\n`);
  const auth = new OneNoteAuth(config, log);
  const client = new OneNoteClient(auth, config, log);

  switch (command) {
    case 'auth': {
      if (await auth.hasCachedAccount()) {
        try {
          await auth.getAccessToken({ interactive: false });
          const account = await auth.getAccount();
          out(`Already signed in as ${account.username}.`);
          out('Run `onenote-cli signout` first if you want to switch accounts.');
          return 0;
        } catch {
          out('Cached credentials could not be renewed. Signing in again...');
        }
      }

      if (config.allowDeviceCode) {
        out('Using device code flow (ONENOTE_ALLOW_DEVICE_CODE is set).');
        out(
          'Note: new Microsoft Entra tenants block this flow by default since 2026-07-01.\n' +
            'If it fails, unset ONENOTE_ALLOW_DEVICE_CODE to use the browser flow.\n'
        );
      }

      const result = await auth.signIn({
        onDeviceCode: (info) => {
          out('');
          out(`  Open:  ${info.verificationUri}`);
          out(`  Code:  ${info.userCode}`);
          out('');
          out('Waiting for you to complete sign-in...');
        }
      });

      out(`\nSigned in as ${result.account?.username ?? 'unknown account'} via ${result.method}.`);
      out(`Credentials cached at ${config.cachePath} (mode 0600).`);
      return 0;
    }

    case 'status': {
      const account = await auth.getAccount();
      let usable = false;
      let reason = null;
      if (account) {
        try {
          await auth.getAccessToken({ interactive: false });
          usable = true;
        } catch (error) {
          reason = error.message;
        }
      }

      const data = {
        signedIn: usable,
        account: account?.username ?? null,
        clientId: config.clientId,
        tenant: config.tenant,
        scopes: config.scopes,
        flow: config.allowDeviceCode ? 'device_code' : 'pkce (authorization code)',
        writeEnabled: config.allowWrite,
        tokenCache: inspectCache(config.cachePath)
      };

      emit(flags, data, () => {
        out(`Signed in:     ${usable ? `yes (${data.account})` : 'no'}`);
        if (reason) out(`Reason:        ${reason}`);
        out(`Client ID:     ${data.clientId}`);
        out(`Tenant:        ${data.tenant}`);
        out(`Scopes:        ${data.scopes.join(', ')}`);
        out(`Flow:          ${data.flow}`);
        out(`Writes:        ${data.writeEnabled ? 'enabled' : 'disabled (read-only)'}`);
        out(
          `Token cache:   ${data.tokenCache.path}${
            data.tokenCache.exists ? ` (mode ${data.tokenCache.mode})` : ' (not created yet)'
          }`
        );
      });
      return usable ? 0 : 1;
    }

    case 'signout': {
      const result = await auth.signOut();
      out(result.signedOut ? 'Signed out and cleared token cache.' : 'No cached account found.');
      return 0;
    }

    case 'whoami': {
      const me = await client.whoAmI();
      emit(flags, me, () => out(`${me.displayName} <${me.userPrincipalName}>`));
      return 0;
    }

    case 'notebooks': {
      const notebooks = await client.listNotebooks({ limit: limitOf(flags) });
      emit(flags, { notebooks }, () => {
        if (!notebooks.length) return out('No notebooks found.');
        for (const nb of notebooks) {
          out(`${nb.displayName}${nb.isDefault ? '  (default)' : ''}`);
          out(`  id: ${nb.id}`);
        }
      });
      return 0;
    }

    case 'sections': {
      const sections = await client.listSections({
        notebookId: positional[1],
        limit: limitOf(flags, 100)
      });
      emit(flags, { sections }, () => {
        if (!sections.length) return out('No sections found.');
        for (const s of sections) {
          out(`${s.displayName}${s.notebookName ? `  (${s.notebookName})` : ''}`);
          out(`  id: ${s.id}`);
        }
      });
      return 0;
    }

    case 'pages': {
      const pages = await client.listPages({
        sectionId: positional[1],
        limit: limitOf(flags, 100)
      });
      emit(flags, { pages }, () => {
        if (!pages.length) return out('No pages found.');
        for (const p of pages) {
          out(`${p.title || '(untitled)'}`);
          out(`  modified: ${p.lastModifiedDateTime}`);
          out(`  id: ${p.id}`);
        }
      });
      return 0;
    }

    case 'read': {
      const pageId = positional[1];
      if (!pageId) {
        process.stderr.write('Usage: onenote-cli read <pageId> [--html]\n');
        return 1;
      }
      const meta = await client.getPageMetadata(pageId);
      const { content, format } = await client.getPageContent(pageId, {
        format: flags.html ? 'html' : 'text'
      });
      emit(flags, { page: meta, format, content }, () => {
        out(`# ${meta.title || '(untitled)'}`);
        out(`_Modified ${meta.lastModifiedDateTime}_\n`);
        out(content);
      });
      return 0;
    }

    case 'search': {
      const query = positional.slice(1).join(' ');
      if (!query) {
        process.stderr.write('Usage: onenote-cli search <query> [--content]\n');
        return 1;
      }
      const results = await client.searchPages(query, {
        limit: limitOf(flags, 25),
        searchContent: Boolean(flags.content)
      });
      emit(flags, { query, results }, () => {
        if (!results.length) {
          out(`No pages matched "${query}".`);
          if (!flags.content) out('Try --content to search page bodies as well.');
          return;
        }
        for (const r of results) {
          out(`${r.title || '(untitled)'}  [${r.matchedIn}]`);
          if (r.excerpt) out(`  ${r.excerpt}`);
          out(`  id: ${r.id}`);
        }
      });
      return 0;
    }

    case 'dump': {
      // Replaces read-all-pages.js / get-all-page-contents.js /
      // get-all-page-contents-full.js, which were three variants of this.
      const pages = await client.listPages({ limit: limitOf(flags, 100) });
      const collected = [];

      for (const page of pages) {
        try {
          const { content } = await client.getPageContent(page.id, {
            format: flags.html ? 'html' : 'text'
          });
          collected.push({ ...page, content });
          if (!flags.json) {
            out('='.repeat(72));
            out(`${page.title || '(untitled)'}`);
            out(`Modified ${page.lastModifiedDateTime}`);
            out('='.repeat(72));
            out('');
            out(content);
            out('');
          }
        } catch (error) {
          process.stderr.write(`Skipped "${page.title}": ${error.message}\n`);
        }
      }

      if (flags.json) out(JSON.stringify({ pages: collected }, null, 2));
      else process.stderr.write(`\nRead ${collected.length} of ${pages.length} page(s).\n`);
      return 0;
    }

    case 'create': {
      const [, sectionId, ...titleParts] = positional;
      const title = titleParts.join(' ');
      if (!sectionId || !title) {
        process.stderr.write(
          'Usage: onenote-cli create <sectionId> <title> [--text "..." | --file path]\n'
        );
        return 1;
      }

      let content = '';
      if (typeof flags.file === 'string') {
        content = await readFile(path.resolve(flags.file), 'utf8');
      } else if (typeof flags.text === 'string') {
        content = flags.text;
      }

      const page = await client.createPage({
        sectionId,
        title,
        content,
        isHtml: Boolean(flags.html)
      });

      emit(flags, { page }, () => {
        out(`Created "${page.title}"`);
        out(`  id: ${page.id}`);
        if (page.webUrl) out(`  url: ${page.webUrl}`);
      });
      return 0;
    }

    default:
      process.stderr.write(`Unknown command "${command}".\n\n${USAGE}`);
      return 1;
  }
}

// ------------------------------------------------------------------ setup ---

/** Ask a question on the terminal. Returns '' when stdin is not interactive. */
async function ask(question, { defaultValue = '' } = {}) {
  if (!process.stdin.isTTY) return defaultValue;
  const { createInterface } = await import('node:readline/promises');
  const rl = createInterface({ input: process.stdin, output: process.stdout });
  try {
    const answer = await rl.question(question);
    return answer.trim() || defaultValue;
  } finally {
    rl.close();
  }
}

async function confirm(question, { defaultYes = true } = {}) {
  const suffix = defaultYes ? ' [Y/n] ' : ' [y/N] ';
  const answer = (await ask(question + suffix, { defaultValue: defaultYes ? 'y' : 'n' })).toLowerCase();
  return answer.startsWith('y');
}

/**
 * End-to-end first-run experience: resolve a client ID, sign in, then write the
 * MCP client config. Replaces the three manual steps people most often get
 * wrong (creating .env, finding the config file, editing JSON correctly).
 */
async function runSetup(flags) {
  out('');
  out(style.bold('  OneNote MCP setup'));
  out(style.dim('  ─────────────────────────────────────────────'));
  out('');

  // --- Step 1: application ID --------------------------------------------
  let clientId = (process.env.ONENOTE_CLIENT_ID || '').trim() || BUNDLED_CLIENT_ID.trim();

  if (clientId) {
    out(`${PASS()} Microsoft application ID found.`);
  } else {
    out(`${WARN()} No Microsoft application ID configured.`);
    out('');
    out('  This build does not ship a default, so you need to create one free');
    out('  Microsoft app registration. It takes about three minutes.');
    out('');
    out(`  ${style.bold('1.')} Open ${style.bold('https://entra.microsoft.com')} and sign in`);
    out(`  ${style.bold('2.')} Search "App registrations" -> "New registration"`);
    out(`  ${style.bold('3.')} Name it anything, e.g. "OneNote for Claude"`);
    out(`  ${style.bold('4.')} Account types: ${style.bold('Accounts in any organizational directory and personal Microsoft accounts')}`);
    out(`  ${style.bold('5.')} Click Register`);
    out(`  ${style.bold('6.')} Left menu -> Authentication -> Add a platform`);
    out(`  ${style.bold('7.')} Choose "Mobile and desktop applications", tick ${style.bold('http://localhost')}, Configure`);
    out(`  ${style.bold('8.')} Left menu -> Overview -> copy the ${style.bold('Application (client) ID')}`);
    out('');
    out(style.dim('  You do NOT need a client secret.'));
    out('');

    if (!process.stdin.isTTY) {
      out(`${FAIL()} Cannot prompt for the ID (not an interactive terminal).`);
      out('  Set ONENOTE_CLIENT_ID and run setup again.');
      return 1;
    }

    clientId = await ask('  Paste the Application (client) ID here: ');
    if (!clientId) {
      out(`${FAIL()} No ID entered. Run setup again when you have it.`);
      return 1;
    }
  }

  // Validate before doing anything irreversible.
  let config;
  try {
    config = loadConfig({ ...process.env, ONENOTE_CLIENT_ID: clientId });
  } catch (error) {
    out('');
    out(`${FAIL()} ${error.message.split('\n')[0]}`);
    out('');
    return 1;
  }

  // Persist to .env so the server picks it up without further configuration.
  if (config.clientIdSource === 'environment' && !BUNDLED_CLIENT_ID.trim()) {
    const envPath = path.join(__dirname, '.env');
    const existing = await readFile(envPath, 'utf8').catch(() => '');
    if (!/^ONENOTE_CLIENT_ID=.+$/m.test(existing)) {
      const next = existing
        ? `${existing.replace(/\n*$/, '\n')}ONENOTE_CLIENT_ID=${clientId}\n`
        : `ONENOTE_CLIENT_ID=${clientId}\n`;
      await (await import('node:fs/promises')).writeFile(envPath, next, { mode: 0o600 });
      out(`${PASS()} Saved to .env`);
    }
  }

  // --- Step 2: sign in ----------------------------------------------------
  out('');
  const log = () => {};
  const auth = new OneNoteAuth(config, log);

  let signedIn = false;
  if (await auth.hasCachedAccount()) {
    try {
      await auth.getAccessToken({ interactive: false });
      const account = await auth.getAccount();
      out(`${PASS()} Already signed in as ${style.bold(account.username)}`);
      signedIn = true;
    } catch {
      /* fall through to interactive sign-in */
    }
  }

  if (!signedIn && flags['dry-run']) {
    out(`${WARN()} Not signed in. Skipping sign-in because this is a dry run.`);
  } else if (!signedIn && !process.stdin.isTTY) {
    // Launching a browser and blocking would hang a non-interactive caller.
    out(`${WARN()} Not signed in, and this is not an interactive terminal.`);
    out('  Run `onenote-cli auth` from a terminal to sign in.');
  } else if (!signedIn) {
    out(style.dim('  Opening your browser to sign in to Microsoft...'));
    try {
      const result = await auth.signIn({
        onDeviceCode: (info) => {
          out('');
          out(`  Open ${style.bold(info.verificationUri)}`);
          out(`  Enter code ${style.bold(info.userCode)}`);
          out('');
        }
      });
      out(`${PASS()} Signed in as ${style.bold(result.account?.username ?? 'your account')}`);
      signedIn = true;
    } catch (error) {
      out(`${FAIL()} Sign-in failed.`);
      out('');
      out(`  ${error.message.split('\n').join('\n  ')}`);
      out('');
      return 1;
    }
  }

  // --- Step 3: verify access ---------------------------------------------
  if (signedIn) {
    try {
      const client = new OneNoteClient(auth, config, log);
      const notebooks = await client.listNotebooks({ limit: 5 });
      out(`${PASS()} Found ${notebooks.length} notebook(s)`);
    } catch (error) {
      out(`${WARN()} Signed in, but could not list notebooks: ${error.message.split('\n')[0]}`);
    }
  }

  // --- Step 4: configure the AI app --------------------------------------
  out('');
  const detected = detectInstalledClients();
  const targets = flags.client
    ? [{ key: flags.client, label: CLIENTS[flags.client]?.label ?? flags.client, configPath: configPathFor(flags.client) }]
    : detected;

  if (!targets.length) {
    out(`${WARN()} No supported AI app detected (Claude Desktop, Cursor, Windsurf).`);
    out('  Install one, then run `onenote-cli setup` again.');
    return 0;
  }

  const serverConfig = {
    command: process.execPath,
    args: [path.join(__dirname, 'onenote-mcp.mjs')],
    env: BUNDLED_CLIENT_ID.trim() ? {} : { ONENOTE_CLIENT_ID: config.clientId }
  };

  for (const target of targets) {
    if (!target.configPath) {
      out(`${FAIL()} Unknown app "${target.key}".`);
      continue;
    }

    if (flags['dry-run']) {
      const result = await writeServerEntry({
        configPath: target.configPath,
        serverName: 'onenote',
        serverConfig,
        dryRun: true
      });
      out(`${WARN()} Dry run -- ${target.label} would be updated at:`);
      out(`  ${result.configPath}`);
      if (result.otherServers.length) {
        out(style.dim(`  Preserving ${result.otherServers.length} other server(s): ${result.otherServers.join(', ')}`));
      }
      continue;
    }

    try {
      const result = await writeServerEntry({
        configPath: target.configPath,
        serverName: 'onenote',
        serverConfig
      });
      out(`${PASS()} ${result.existed ? 'Updated' : 'Configured'} ${style.bold(target.label)}`);
      out(style.dim(`  ${result.configPath}`));
      if (result.backupPath) out(style.dim(`  Backup: ${path.basename(result.backupPath)}`));
      if (result.otherServers.length) {
        out(style.dim(`  Left ${result.otherServers.length} other server(s) untouched`));
      }
    } catch (error) {
      out(`${FAIL()} Could not configure ${target.label}: ${error.message}`);
    }
  }

  out('');
  out(style.bold('  Done.'));
  for (const target of targets) {
    const hint = CLIENTS[target.key]?.restartHint;
    if (hint) out(`  ${hint}`);
  }
  out('');
  out(style.dim('  Then try asking: "List my OneNote notebooks"'));
  out('');
  return 0;
}

// ----------------------------------------------------------------- doctor ---

/** Diagnose every layer between the user and their notes, and say how to fix it. */
async function runDoctor(flags) {
  const checks = [];
  const add = (name, status, detail, fix) => checks.push({ name, status, detail, fix });

  // Runtime
  const major = Number(process.versions.node.split('.')[0]);
  if (major >= 20) add('Node.js', 'pass', `v${process.versions.node}`);
  else add('Node.js', 'fail', `v${process.versions.node}`, 'Install Node.js 20 or newer from https://nodejs.org');

  // Configuration
  let config = null;
  try {
    config = loadConfig();
    add(
      'Application ID',
      'pass',
      `${config.clientId} (${config.clientIdSource === 'bundled' ? 'built in' : 'from environment'})`
    );
  } catch (error) {
    add('Application ID', 'fail', error.message.split('\n')[0], 'Run `onenote-cli setup`');
  }

  // Credentials
  if (config) {
    const auth = new OneNoteAuth(config, () => {});
    const cache = inspectCache(config.cachePath);

    if (!cache.exists) {
      add('Credentials', 'fail', 'no saved sign-in', 'Run `onenote-cli auth`');
    } else if (cache.mode !== '600') {
      add('Credentials', 'warn', `file mode is ${cache.mode}, expected 600`, `chmod 600 "${cache.path}"`);
    } else {
      add('Credentials', 'pass', `cached, mode ${cache.mode}`);
    }

    if (cache.exists) {
      try {
        await auth.getAccessToken({ interactive: false });
        const account = await auth.getAccount();
        add('Microsoft sign-in', 'pass', account?.username ?? 'valid token');
      } catch (error) {
        add('Microsoft sign-in', 'fail', error.message.split('\n')[0], 'Run `onenote-cli auth`');
      }

      // Reachability
      try {
        const client = new OneNoteClient(auth, config, () => {});
        const notebooks = await client.listNotebooks({ limit: 5 });
        add('OneNote access', 'pass', `${notebooks.length} notebook(s) visible`);
      } catch (error) {
        add('OneNote access', 'fail', error.message.split('\n')[0], 'Check scopes and consent, then re-run auth');
      }
    }

    if (config.allowDeviceCode) {
      add(
        'Sign-in method',
        'warn',
        'device code flow enabled',
        'New Microsoft tenants block this by default since 2026-07-01. Unset ONENOTE_ALLOW_DEVICE_CODE.'
      );
    } else {
      add('Sign-in method', 'pass', 'browser (PKCE)');
    }

    add('Write access', config.allowWrite ? 'pass' : 'warn', config.allowWrite ? 'enabled' : 'read-only mode');
  }

  // AI app wiring
  const detected = detectInstalledClients();
  if (!detected.length) {
    add('AI app', 'warn', 'none detected', 'Install Claude Desktop, Cursor, or Windsurf');
  } else {
    for (const client of detected) {
      let configured = false;
      try {
        const raw = await readFile(client.configPath, 'utf8');
        configured = Boolean(JSON.parse(raw)?.mcpServers?.onenote);
      } catch {
        configured = false;
      }
      add(
        client.label,
        configured ? 'pass' : 'warn',
        configured ? 'configured' : 'installed but not configured',
        configured ? undefined : 'Run `onenote-cli setup`'
      );
    }
  }

  if (flags.json) {
    out(JSON.stringify({ checks }, null, 2));
    return checks.some((c) => c.status === 'fail') ? 1 : 0;
  }

  out('');
  out(style.bold('  OneNote MCP diagnostics'));
  out(style.dim('  ─────────────────────────────────────────────'));
  for (const check of checks) {
    const badge = check.status === 'pass' ? PASS() : check.status === 'warn' ? WARN() : FAIL();
    out(`${badge} ${check.name.padEnd(18)} ${style.dim(check.detail ?? '')}`);
    if (check.fix) out(`        ${style.dim('->')} ${check.fix}`);
  }

  const failed = checks.filter((c) => c.status === 'fail');
  out('');
  if (failed.length) {
    out(`  ${style.red(`${failed.length} problem(s) found.`)} Fix the items marked -> above.`);
  } else {
    out(`  ${style.green('Everything looks good.')}`);
  }
  out('');
  return failed.length ? 1 : 0;
}

// -------------------------------------------------------------- uninstall ---

async function runUninstall(flags) {
  const detected = flags.client
    ? [{ key: flags.client, label: CLIENTS[flags.client]?.label ?? flags.client, configPath: configPathFor(flags.client) }]
    : detectInstalledClients();

  let removed = 0;
  for (const target of detected) {
    if (!target.configPath) continue;
    const result = await removeServerEntry({ configPath: target.configPath, serverName: 'onenote' });
    if (result.removed) {
      out(`${PASS()} Removed from ${target.label}`);
      out(style.dim(`  Backup: ${path.basename(result.backupPath)}`));
      removed += 1;
    }
  }

  if (!removed) out(`${WARN()} No configured AI app found.`);

  if (await confirm('  Also delete your saved Microsoft sign-in?', { defaultYes: false })) {
    try {
      const config = loadConfig();
      const auth = new OneNoteAuth(config, () => {});
      await auth.signOut();
      out(`${PASS()} Signed out and cleared credentials`);
    } catch {
      out(`${WARN()} Could not clear credentials (already gone?)`);
    }
  }

  return 0;
}

/** Same permissive .env handling as the server: never override the real environment. */
async function loadDotEnvIfPresent() {
  if (process.env.ONENOTE_SKIP_DOTENV) return;
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
    if (error.code !== 'ENOENT') process.stderr.write(`Could not read .env: ${error.message}\n`);
  }
}

main()
  .then((code) => process.exit(code ?? 0))
  .catch((error) => {
    process.stderr.write(`\n${error.message}\n`);
    if (process.env.ONENOTE_DEBUG) process.stderr.write(`${error.stack}\n`);
    process.exit(1);
  });
