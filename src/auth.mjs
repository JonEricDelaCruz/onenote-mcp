/**
 * Microsoft Entra authentication for OneNote access.
 *
 * Primary flow: authorization code + PKCE over a loopback redirect. This is the
 * flow Microsoft recommends for native/desktop/CLI apps, and it is not affected
 * by the device-code restrictions below.
 *
 * Fallback flow: device code, opt-in only via ONENOTE_ALLOW_DEVICE_CODE=true.
 * Since 2026-07-01 all new Microsoft Entra tenants block device code flow as
 * part of security defaults, and many existing tenants block it via Conditional
 * Access. It remains useful for headless machines whose tenant permits it.
 *
 * Tokens are persisted through MSAL's own cache (which holds the refresh token
 * and handles silent renewal). The cache file is written with 0600 permissions
 * and stored outside the repository.
 */

import fs from 'node:fs';
import fsp from 'node:fs/promises';
import path from 'node:path';
import { PublicClientApplication, LogLevel } from '@azure/msal-node';

/** Written to the browser tab after a successful loopback redirect. */
const SUCCESS_HTML = `<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><title>Signed in</title>
<style>body{font:16px/1.5 system-ui,-apple-system,sans-serif;margin:0;display:grid;
place-items:center;height:100vh;background:#faf9f7;color:#1a1a19}
.card{text-align:center;padding:2.5rem 3rem;background:#fff;border:1px solid #e5e3df;border-radius:12px}
h1{font-size:1.125rem;margin:0 0 .5rem}p{margin:0;color:#6b6a67;font-size:.9375rem}</style>
</head><body><div class="card"><h1>Signed in to OneNote</h1>
<p>You can close this tab and return to your terminal.</p></div></body></html>`;

const FAILURE_HTML = `<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><title>Sign-in failed</title>
<style>body{font:16px/1.5 system-ui,-apple-system,sans-serif;margin:0;display:grid;
place-items:center;height:100vh;background:#faf9f7;color:#1a1a19}
.card{text-align:center;padding:2.5rem 3rem;background:#fff;border:1px solid #e5e3df;border-radius:12px}
h1{font-size:1.125rem;margin:0 0 .5rem}p{margin:0;color:#6b6a67;font-size:.9375rem}</style>
</head><body><div class="card"><h1>Sign-in failed</h1>
<p>Check your terminal for details.</p></div></body></html>`;

export class AuthError extends Error {
  constructor(message, { cause } = {}) {
    super(message);
    this.name = 'AuthError';
    if (cause) this.cause = cause;
  }
}

/**
 * File-backed MSAL cache plugin.
 *
 * Writes are atomic (temp file + rename) and mode 0600. On POSIX the directory
 * is 0700 as well, so the cache is not readable by other local users.
 */
function createCachePlugin(cachePath, log) {
  const dir = path.dirname(cachePath);

  return {
    async beforeCacheAccess(ctx) {
      try {
        const data = await fsp.readFile(cachePath, 'utf8');
        if (data.trim()) ctx.tokenCache.deserialize(data);
      } catch (error) {
        if (error.code !== 'ENOENT') {
          log(`Could not read token cache (${error.code}); starting empty.`);
        }
      }
    },

    async afterCacheAccess(ctx) {
      if (!ctx.cacheHasChanged) return;
      try {
        await fsp.mkdir(dir, { recursive: true, mode: 0o700 });
        const tmp = `${cachePath}.${process.pid}.tmp`;
        await fsp.writeFile(tmp, ctx.tokenCache.serialize(), { mode: 0o600 });
        await fsp.rename(tmp, cachePath);
        // rename preserves the temp file's mode, but be explicit for the case
        // where cachePath already existed with looser permissions.
        await fsp.chmod(cachePath, 0o600).catch(() => {});
      } catch (error) {
        log(`Could not persist token cache: ${error.message}`);
      }
    }
  };
}

export class OneNoteAuth {
  /**
   * @param {object} config Result of loadConfig().
   * @param {(msg: string) => void} [log] Diagnostic sink. Must not write to
   *   stdout when running as a stdio MCP server -- stdout is the JSON-RPC channel.
   */
  constructor(config, log = () => {}) {
    this.config = config;
    this.log = log;

    this.msal = new PublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: config.authority
      },
      cache: {
        cachePlugin: createCachePlugin(config.cachePath, log)
      },
      system: {
        loggerOptions: {
          // MSAL messages can contain PII; only surface warnings and errors,
          // and never enable PII logging.
          piiLoggingEnabled: false,
          logLevel: LogLevel.Warning,
          loggerCallback: (level, message, containsPii) => {
            if (!containsPii) log(`[msal] ${message}`);
          }
        }
      }
    });
  }

  /** Scopes minus OIDC scopes, which MSAL adds implicitly and rejects as input. */
  get #requestScopes() {
    const reserved = new Set(['openid', 'profile', 'offline_access', 'email']);
    return this.config.scopes.filter((s) => !reserved.has(s.toLowerCase()));
  }

  /** @returns {Promise<import('@azure/msal-node').AccountInfo | null>} */
  async getAccount() {
    const accounts = await this.msal.getAllAccounts();
    return accounts.length ? accounts[0] : null;
  }

  /** True if a cached account exists (does not prove the refresh token is still valid). */
  async hasCachedAccount() {
    return (await this.getAccount()) !== null;
  }

  /**
   * Return a valid access token, refreshing silently when possible.
   *
   * @param {object} [options]
   * @param {boolean} [options.interactive=true] Allow an interactive sign-in if
   *   silent acquisition fails. Set false in contexts where opening a browser
   *   would be wrong (e.g. answering a tool call without user intent).
   * @param {boolean} [options.forceRefresh=false] Bypass the cached access token
   *   and redeem the refresh token. Used after Graph rejects a token that MSAL
   *   still considers valid (revoked session, changed password, policy change).
   */
  async getAccessToken({ interactive = true, forceRefresh = false } = {}) {
    const account = await this.getAccount();

    if (account) {
      try {
        const result = await this.msal.acquireTokenSilent({
          account,
          scopes: this.#requestScopes,
          forceRefresh
        });
        if (result?.accessToken) return result.accessToken;
      } catch (error) {
        this.log(`Silent token acquisition failed (${error.errorCode || error.name}); reauthenticating.`);
      }
    }

    if (!interactive) {
      throw new AuthError(
        'Not authenticated. Run `npm run auth` (or call the `authenticate` tool) to sign in.'
      );
    }

    const result = await this.signIn();
    return result.accessToken;
  }

  /**
   * Perform an interactive sign-in.
   *
   * @param {object} [hooks]
   * @param {(info: object) => void} [hooks.onDeviceCode] Called as soon as Entra
   *   issues a device code, before MSAL begins polling. Callers running inside an
   *   MCP tool use this to show the code without waiting for the flow to finish.
   * @param {(url: string) => void} [hooks.onBrowserOpen] Called with the
   *   authorization URL just before the browser is launched.
   * @returns {Promise<{accessToken: string, account: object, method: string}>}
   */
  async signIn(hooks = {}) {
    if (this.config.allowDeviceCode) {
      return this.#signInWithDeviceCode(hooks);
    }
    return this.#signInWithPkce(hooks);
  }

  /** Authorization code + PKCE via a loopback listener. MSAL generates and verifies the PKCE pair. */
  async #signInWithPkce({ onBrowserOpen } = {}) {
    const { default: open } = await import('open');

    try {
      const result = await this.msal.acquireTokenInteractive({
        scopes: this.#requestScopes,
        successTemplate: SUCCESS_HTML,
        errorTemplate: FAILURE_HTML,
        preferredPort: this.config.redirectPort,
        openBrowser: async (url) => {
          this.log('Opening your browser to sign in to Microsoft...');
          this.log(`If it does not open, visit:\n${url}`);
          onBrowserOpen?.(url);
          try {
            await open(url);
          } catch (error) {
            // A headless or locked-down host may have no browser. The URL is
            // already surfaced, so let the user open it themselves.
            this.log(`Could not launch a browser (${error.message}). Open the URL manually.`);
          }
        }
      });

      return { accessToken: result.accessToken, account: result.account, method: 'pkce' };
    } catch (error) {
      throw new AuthError(this.#explain(error), { cause: error });
    }
  }

  /** Device code flow. Opt-in; blocked by default in new tenants since 2026-07-01. */
  async #signInWithDeviceCode({ onDeviceCode } = {}) {
    try {
      const result = await this.msal.acquireTokenByDeviceCode({
        scopes: this.#requestScopes,
        deviceCodeCallback: (info) => {
          this.deviceCodeMessage = info.message;
          this.log(info.message);
          onDeviceCode?.(info);
        }
      });

      if (!result?.accessToken) {
        throw new AuthError('Device code flow completed without returning a token.');
      }
      return { accessToken: result.accessToken, account: result.account, method: 'device_code' };
    } catch (error) {
      if (error instanceof AuthError) throw error;
      throw new AuthError(this.#explain(error), { cause: error });
    }
  }

  /** Remove the cached account and delete the cache file. */
  async signOut() {
    const account = await this.getAccount();
    if (account) {
      const cache = this.msal.getTokenCache();
      await cache.removeAccount(account);
    }
    try {
      await fsp.unlink(this.config.cachePath);
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
    return { signedOut: Boolean(account), cachePath: this.config.cachePath };
  }

  /** Turn MSAL/Entra error codes into something a user can act on. */
  #explain(error) {
    const code = error?.errorCode || '';
    const message = error?.errorMessage || error?.message || String(error);

    if (/AADSTS7000218|invalid_client/i.test(message)) {
      return (
        'Entra rejected the client. Your app registration is likely configured as a ' +
        'confidential/web client. Under Authentication, add a "Mobile and desktop applications" ' +
        'platform with redirect URI http://localhost, and ensure "Allow public client flows" is enabled.\n\n' +
        `Original error: ${message}`
      );
    }

    if (/AADSTS50011|redirect_uri/i.test(message)) {
      return (
        'Redirect URI mismatch. Add http://localhost as a "Mobile and desktop applications" ' +
        'redirect URI on your app registration. If you set ONENOTE_REDIRECT_PORT, register ' +
        `http://localhost:${this.config.redirectPort} exactly.\n\n` +
        `Original error: ${message}`
      );
    }

    if (/AADSTS16000/i.test(message)) {
      return (
        'Microsoft could not complete an interactive sign-in for this account.\n\n' +
        'If the account is a personal Microsoft account (outlook.com, hotmail.com, live.com), ' +
        'it may be signed in to the "Microsoft Services" holding tenant, which has no directory. ' +
        'Sign out of all Microsoft accounts in your browser, then try again in a private window ' +
        'and pick the account that owns your OneNote notebooks.\n\n' +
        `Original error: ${message}`
      );
    }

    if (/AADSTS50020|does not exist in tenant/i.test(message)) {
      return (
        'The account you signed in with is not allowed by this application.\n\n' +
        'This almost always means the app registration\'s "Supported account types" is too narrow ' +
        'for the account you chose. In the Microsoft Entra portal, open your app registration -> ' +
        'Authentication -> Supported account types, and select:\n\n' +
        '  "Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) ' +
        'and personal Microsoft accounts (e.g. Skype, Xbox)"\n\n' +
        'If that option is greyed out, open the Manifest blade instead and set:\n' +
        '  "signInAudience": "AzureADandPersonalMicrosoftAccount"\n\n' +
        'Alternatively, sign in with the same account that created the app registration.\n\n' +
        `Original error: ${message}`
      );
    }

    if (/AADSTS50105/i.test(message)) {
      return (
        'Your tenant requires users to be explicitly assigned to this application. ' +
        'Ask an administrator to assign your account under Enterprise applications -> ' +
        'your app -> Users and groups.\n\n' +
        `Original error: ${message}`
      );
    }

    if (/AADSTS65001|consent/i.test(message)) {
      return (
        'Consent was not granted for the requested scopes ' +
        `(${this.config.scopes.join(', ')}). Approve the consent prompt, or ask an admin to ` +
        'grant consent if your tenant requires admin consent.\n\n' +
        `Original error: ${message}`
      );
    }

    if (/AADSTS700016|application.+not found/i.test(message)) {
      return (
        `No application with client ID ${this.config.clientId} was found in tenant ` +
        `"${this.config.tenant}". Check ONENOTE_CLIENT_ID, and set ONENOTE_TENANT_ID if your ` +
        'app is single-tenant.\n\n' +
        `Original error: ${message}`
      );
    }

    if (/device.?code|AADSTS7000220|authorization_pending|expired_token/i.test(`${code} ${message}`)) {
      return (
        'Device code flow failed. Since 2026-07-01 new Microsoft Entra tenants block this flow ' +
        'by default under security defaults, and many tenants block it via Conditional Access.\n\n' +
        'Unset ONENOTE_ALLOW_DEVICE_CODE to use the browser-based PKCE flow instead.\n\n' +
        `Original error: ${message}`
      );
    }

    return `Authentication failed: ${message}`;
  }
}

/**
 * Non-blocking sign-in coordinator for MCP tool calls.
 *
 * An MCP `authenticate` tool must not block while waiting for a human. Device
 * codes are valid for ~15 minutes and MSAL polls for the entire window, so
 * awaiting the flow inside a tool handler guarantees a client-side timeout --
 * and, because the device code was only ever written to stderr, the user never
 * saw the code they were supposed to enter (danosb/onenote-mcp#6, #1).
 *
 * This class starts the flow in the background and resolves as soon as there is
 * something actionable to *show* the user (the device code, or confirmation
 * that a browser was opened). Callers return that immediately; completion is
 * observed later via `snapshot()`.
 */
export class AuthSession {
  /** @param {OneNoteAuth} auth */
  constructor(auth) {
    this.auth = auth;
    /** @type {'idle'|'pending'|'succeeded'|'failed'} */
    this.state = 'idle';
    this.prompt = null;
    this.error = null;
    this.startedAt = null;
    this.completedAt = null;
    this.method = null;
    this.promise = null;
  }

  get inProgress() {
    return this.state === 'pending';
  }

  /**
   * Begin a sign-in if one is not already running.
   *
   * @returns {Promise<{state: string, prompt: object|null, method: string|null, error: string|null}>}
   *   Resolves once the user has been told what to do -- not once sign-in finishes.
   */
  async begin() {
    if (this.state === 'pending') return this.snapshot();

    this.state = 'pending';
    this.prompt = null;
    this.error = null;
    this.startedAt = new Date().toISOString();
    this.completedAt = null;
    this.method = this.auth.config.allowDeviceCode ? 'device_code' : 'pkce';

    // Resolves as soon as the user-facing instruction is known.
    let announce;
    const announced = new Promise((resolve) => {
      announce = resolve;
    });

    this.promise = (async () => {
      try {
        const result = await this.auth.signIn({
          // Surface the device code the moment Entra issues it, before polling.
          onDeviceCode: (info) => {
            this.prompt = {
              kind: 'device_code',
              verificationUri: info.verificationUri,
              userCode: info.userCode,
              expiresInSeconds: info.expiresIn ?? null,
              message: info.message
            };
            announce();
          },
          // The browser is opened on this machine; tell the user to look at it.
          onBrowserOpen: (url) => {
            this.prompt = {
              kind: 'browser',
              url,
              message:
                'A browser window was opened on the machine running this server. ' +
                'Complete the Microsoft sign-in there.'
            };
            announce();
          }
        });

        this.state = 'succeeded';
        this.method = result.method;
        this.completedAt = new Date().toISOString();
        this.account = result.account?.username ?? null;
      } catch (error) {
        this.state = 'failed';
        this.error = error.message;
        this.completedAt = new Date().toISOString();
      } finally {
        announce();
      }
    })();

    // Do not let an unobserved rejection escape; state carries the failure.
    this.promise.catch(() => {});

    // Bound the wait so a hung network call cannot stall the tool response.
    await Promise.race([announced, delay(20_000)]);
    return this.snapshot();
  }

  snapshot() {
    return {
      state: this.state,
      method: this.method,
      prompt: this.prompt,
      error: this.error,
      startedAt: this.startedAt,
      completedAt: this.completedAt,
      account: this.account ?? null
    };
  }
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms).unref?.() ?? setTimeout(resolve, ms));
}

/** Report cache file location and permissions, for the CLI `status` command. */
export function inspectCache(cachePath) {
  try {
    const stat = fs.statSync(cachePath);
    return {
      exists: true,
      path: cachePath,
      mode: (stat.mode & 0o777).toString(8).padStart(3, '0'),
      modified: stat.mtime.toISOString()
    };
  } catch {
    return { exists: false, path: cachePath };
  }
}
