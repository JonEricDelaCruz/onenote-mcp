/**
 * Configuration, loaded from environment (optionally via a .env file).
 */

import path from 'node:path';
import os from 'node:os';

// ===========================================================================
//  >>> PASTE YOUR MICROSOFT APPLICATION (CLIENT) ID BETWEEN THE QUOTES <<<
//
//  Leave it empty and the tool still works -- it will simply ask each user to
//  register their own Microsoft app first (see SETUP-FOR-JON.md). Fill it in
//  and that step disappears for everyone.
//
//  It looks like: 3f9a21c4-8e7b-4d12-9c55-1a2b3c4d5e6f
//
//  This value is NOT a secret. It is a public identifier, exactly like the
//  ones the Azure CLI and GitHub CLI publish in their own source. It is safe
//  in a public repository: this is a "public client" using PKCE with a
//  localhost-only redirect, so a sign-in can only ever deliver a token to the
//  machine the user is sitting at. It grants no access on its own.
// ===========================================================================
export const BUNDLED_CLIENT_ID = '40a14204-42d0-43f4-b88d-ac395fa7827e';

/**
 * The Microsoft first-party client ID that earlier versions of this project
 * borrowed ("Microsoft Graph Command Line Tools"). Rejected on sight: it is not
 * ours to use, tenants frequently block it, and it hides the real consent
 * surface from the person signing in.
 */
const BORROWED_MICROSOFT_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';

/** Shown when no client ID is available from any source. */
const REGISTRATION_HELP = `No Microsoft application ID is configured.

This build does not ship a default, so you will need to create your own free
Microsoft app registration once. It takes about three minutes:

  1. Go to https://entra.microsoft.com and sign in
  2. Search for "App registrations" and click "New registration"
  3. Name it anything, e.g. "OneNote for Claude"
  4. Under "Supported account types", choose:
       Accounts in any organizational directory and personal Microsoft accounts
  5. Click "Register"
  6. Click "Authentication" in the left menu
  7. Click "Add a platform" -> "Mobile and desktop applications"
  8. Tick the http://localhost checkbox, then click "Configure"
  9. Click "Overview" and copy the "Application (client) ID"

Then set it, either by:
  - putting ONENOTE_CLIENT_ID=<the id> in your .env file, or
  - setting it in your MCP client's configuration, or
  - running: npx onenote-cli setup

You do NOT need a client secret.`;

/** Scopes for the user's own notebooks. `.All` variants are opt-in via ONENOTE_SCOPES. */
export const DEFAULT_SCOPES = ['Notes.ReadWrite', 'offline_access'];

const TRUTHY = new Set(['1', 'true', 'yes', 'on']);

function bool(value, fallback = false) {
  if (value === undefined || value === '') return fallback;
  return TRUTHY.has(String(value).trim().toLowerCase());
}

/**
 * Resolve the token cache location.
 *
 * Defaults to the user's config directory rather than the repository working
 * tree, so a cache file can never be committed by accident.
 */
function resolveCachePath(raw) {
  if (raw && raw.trim()) return path.resolve(raw.trim());

  const home = os.homedir();
  if (process.platform === 'win32') {
    const base = process.env.APPDATA || path.join(home, 'AppData', 'Roaming');
    return path.join(base, 'onenote-mcp', 'token-cache.json');
  }
  if (process.platform === 'darwin') {
    return path.join(home, 'Library', 'Application Support', 'onenote-mcp', 'token-cache.json');
  }
  const base = process.env.XDG_CONFIG_HOME || path.join(home, '.config');
  return path.join(base, 'onenote-mcp', 'token-cache.json');
}

export class ConfigError extends Error {
  constructor(message) {
    super(message);
    this.name = 'ConfigError';
  }
}

export function loadConfig(env = process.env) {
  // A user-supplied ID always wins over the bundled one, so anyone can point
  // this at their own registration without editing source.
  const supplied = (env.ONENOTE_CLIENT_ID || '').trim();
  const clientId = supplied || BUNDLED_CLIENT_ID.trim();
  const clientIdSource = supplied ? 'environment' : BUNDLED_CLIENT_ID.trim() ? 'bundled' : 'none';

  if (!clientId) {
    throw new ConfigError(REGISTRATION_HELP);
  }

  if (clientId.toLowerCase() === BORROWED_MICROSOFT_CLIENT_ID) {
    throw new ConfigError(
      'The configured application ID belongs to Microsoft ("Microsoft Graph Command Line\n' +
        'Tools"), not to this project. Many tenants block it, and using it misrepresents this\n' +
        'tool to your identity provider.\n\n' +
        `${REGISTRATION_HELP}`
    );
  }

  // Catch a half-finished setup early, with a clearer message than Entra's.
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(clientId)) {
    throw new ConfigError(
      `"${clientId}" does not look like a Microsoft application ID.\n\n` +
        'It should be a GUID, like 3f9a21c4-8e7b-4d12-9c55-1a2b3c4d5e6f.\n' +
        'Copy the "Application (client) ID" from your app registration Overview page --\n' +
        'not the Object ID, and not the Directory (tenant) ID.'
    );
  }

  const scopes = (env.ONENOTE_SCOPES || '')
    .split(/[,\s]+/)
    .map((s) => s.trim())
    .filter(Boolean);

  const tenant = (env.ONENOTE_TENANT_ID || 'common').trim();

  return {
    clientId,
    clientIdSource,
    tenant,
    authority: `https://login.microsoftonline.com/${tenant}`,
    scopes: scopes.length ? scopes : DEFAULT_SCOPES,
    cachePath: resolveCachePath(env.ONENOTE_TOKEN_CACHE),
    /** Device code flow is blocked by default in new Entra tenants since 2026-07-01. */
    allowDeviceCode: bool(env.ONENOTE_ALLOW_DEVICE_CODE, false),
    /** Fixed loopback port; must match a registered redirect URI if set. */
    redirectPort: env.ONENOTE_REDIRECT_PORT ? Number(env.ONENOTE_REDIRECT_PORT) : undefined,
    /** Tools that modify OneNote are opt-in. */
    allowWrite: bool(env.ONENOTE_ALLOW_WRITE, true),
    /**
     * Optional pinned section (name, "Notebook / Section" path, or ID). When
     * set, writes default here and the assistant can skip discovery entirely.
     */
    defaultSection: (env.ONENOTE_DEFAULT_SECTION || '').trim() || null,
    graphBaseUrl: 'https://graph.microsoft.com/v1.0'
  };
}
