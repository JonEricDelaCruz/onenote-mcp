import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { loadConfig, ConfigError, DEFAULT_SCOPES, BUNDLED_CLIENT_ID } from '../src/config.mjs';

const BORROWED_MICROSOFT_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
const VALID = '3f9a21c4-8e7b-4d12-9c55-1a2b3c4d5e6f';

describe('application ID resolution', () => {
  test('rejects the borrowed Microsoft first-party client ID', () => {
    assert.throws(
      () => loadConfig({ ONENOTE_CLIENT_ID: BORROWED_MICROSOFT_CLIENT_ID }),
      (error) => error instanceof ConfigError && /belongs to Microsoft/i.test(error.message)
    );
  });

  test('rejects it case-insensitively too', () => {
    assert.throws(
      () => loadConfig({ ONENOTE_CLIENT_ID: BORROWED_MICROSOFT_CLIENT_ID.toUpperCase() }),
      ConfigError
    );
  });

  test('rejects a value that is not a GUID, with a helpful message', () => {
    assert.throws(
      () => loadConfig({ ONENOTE_CLIENT_ID: 'my-cool-app' }),
      (error) =>
        error instanceof ConfigError && /does not look like a Microsoft application ID/.test(error.message)
    );
  });

  test('accepts a well-formed GUID', () => {
    const config = loadConfig({ ONENOTE_CLIENT_ID: VALID });
    assert.equal(config.clientId, VALID);
    assert.equal(config.clientIdSource, 'environment');
  });

  test('an environment value overrides the bundled one', () => {
    const config = loadConfig({ ONENOTE_CLIENT_ID: VALID });
    assert.equal(config.clientId, VALID);
    assert.equal(config.clientIdSource, 'environment');
  });

  test('behaviour with no ID matches whether an ID is bundled', () => {
    // The repository ships with BUNDLED_CLIENT_ID empty so that the tool still
    // works before a maintainer fills it in. Once filled, no-config must
    // silently succeed instead of erroring. Assert whichever applies, so this
    // test keeps passing after the constant is set.
    if (BUNDLED_CLIENT_ID.trim()) {
      const config = loadConfig({});
      assert.equal(config.clientIdSource, 'bundled');
      assert.equal(config.clientId, BUNDLED_CLIENT_ID.trim());
    } else {
      assert.throws(
        () => loadConfig({}),
        (error) =>
          error instanceof ConfigError && /No Microsoft application ID is configured/.test(error.message)
      );
    }
  });

  test('the registration instructions are actionable when no ID exists', () => {
    if (BUNDLED_CLIENT_ID.trim()) return; // not applicable once bundled
    try {
      loadConfig({});
      assert.fail('expected a ConfigError');
    } catch (error) {
      assert.match(error.message, /entra\.microsoft\.com/);
      assert.match(error.message, /App registrations/);
      assert.match(error.message, /http:\/\/localhost/);
      assert.match(error.message, /do NOT need a client secret/i);
    }
  });

  test('a bundled ID is never the borrowed Microsoft one', () => {
    assert.notEqual(BUNDLED_CLIENT_ID.trim().toLowerCase(), BORROWED_MICROSOFT_CLIENT_ID);
  });
});

describe('defaults', () => {
  test('uses least-privilege scopes for the user own notebooks', () => {
    const config = loadConfig({ ONENOTE_CLIENT_ID: VALID });
    assert.deepEqual(config.scopes, DEFAULT_SCOPES);
    assert.ok(!config.scopes.some((s) => s.endsWith('.All')), 'must not request .All by default');
  });

  test('device code flow is off by default', () => {
    assert.equal(loadConfig({ ONENOTE_CLIENT_ID: VALID }).allowDeviceCode, false);
  });

  test('device code flow can be opted into', () => {
    const config = loadConfig({ ONENOTE_CLIENT_ID: VALID, ONENOTE_ALLOW_DEVICE_CODE: 'true' });
    assert.equal(config.allowDeviceCode, true);
  });

  test('token cache defaults outside the repository', () => {
    const config = loadConfig({ ONENOTE_CLIENT_ID: VALID });
    assert.ok(!config.cachePath.includes('onenote-mcp-main'), 'cache must not land in the repo');
    assert.match(config.cachePath, /onenote-mcp/);
  });

  test('writes are enabled by default and can be disabled', () => {
    assert.equal(loadConfig({ ONENOTE_CLIENT_ID: VALID }).allowWrite, true);
    assert.equal(
      loadConfig({ ONENOTE_CLIENT_ID: VALID, ONENOTE_ALLOW_WRITE: 'false' }).allowWrite,
      false
    );
  });

  test('custom scopes override the default', () => {
    const config = loadConfig({
      ONENOTE_CLIENT_ID: VALID,
      ONENOTE_SCOPES: 'Notes.Read offline_access'
    });
    assert.deepEqual(config.scopes, ['Notes.Read', 'offline_access']);
  });

  test('authority reflects the tenant', () => {
    const config = loadConfig({ ONENOTE_CLIENT_ID: VALID, ONENOTE_TENANT_ID: 'consumers' });
    assert.equal(config.authority, 'https://login.microsoftonline.com/consumers');
  });
});
