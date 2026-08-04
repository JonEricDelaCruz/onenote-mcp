/**
 * Regression tests for the non-blocking sign-in coordinator.
 *
 * The bug being pinned down (danosb/onenote-mcp#6, #1): the v1 `authenticate`
 * tool awaited the whole device code flow. MSAL polls for the full ~15 minute
 * code lifetime, so the tool call always hit the client's timeout -- and because
 * the code was only ever written to stderr, the user never saw what to enter.
 *
 * The contract these tests enforce:
 *   1. begin() resolves as soon as the user-facing prompt exists.
 *   2. It resolves well before the underlying flow completes.
 *   3. The prompt actually carries the verification URI and user code.
 *   4. A failure is recorded as state, never thrown as an unhandled rejection.
 */

import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { AuthSession } from '../src/auth.mjs';

/** Stand-in for OneNoteAuth that mimics MSAL's callback-then-poll timing. */
function fakeAuth({ mode = 'device_code', pollMs = 5000, fail = null } = {}) {
  return {
    config: { allowDeviceCode: mode === 'device_code' },
    async signIn({ onDeviceCode, onBrowserOpen }) {
      if (mode === 'device_code') {
        onDeviceCode?.({
          verificationUri: 'https://microsoft.com/devicelogin',
          userCode: 'ABCD-EFGH',
          expiresIn: 900,
          message: 'Go to https://microsoft.com/devicelogin and enter ABCD-EFGH'
        });
      } else {
        onBrowserOpen?.('https://login.microsoftonline.com/common/oauth2/v2.0/authorize?x=1');
      }

      // The long poll while waiting for the human.
      await new Promise((resolve) => setTimeout(resolve, pollMs));

      if (fail) throw new Error(fail);
      return { accessToken: 'token', account: { username: 'user@example.com' }, method: mode };
    }
  };
}

describe('AuthSession: device code', () => {
  test('returns the code without waiting for the flow to finish', async () => {
    const session = new AuthSession(fakeAuth({ pollMs: 5000 }));

    const started = Date.now();
    const state = await session.begin();
    const elapsed = Date.now() - started;

    assert.ok(
      elapsed < 1000,
      `begin() must return promptly, took ${elapsed}ms (the v1 bug was blocking here)`
    );
    assert.equal(state.state, 'pending');
    assert.equal(state.prompt.kind, 'device_code');
  });

  test('prompt carries the URL and code the user needs', async () => {
    const session = new AuthSession(fakeAuth());
    const { prompt } = await session.begin();

    assert.equal(prompt.verificationUri, 'https://microsoft.com/devicelogin');
    assert.equal(prompt.userCode, 'ABCD-EFGH');
    assert.equal(prompt.expiresInSeconds, 900);
  });

  test('reports success once the background flow completes', async () => {
    const session = new AuthSession(fakeAuth({ pollMs: 50 }));
    await session.begin();
    await session.promise;

    const snapshot = session.snapshot();
    assert.equal(snapshot.state, 'succeeded');
    assert.equal(snapshot.account, 'user@example.com');
    assert.ok(snapshot.completedAt);
  });

  test('a second begin() while pending does not start a duplicate flow', async () => {
    let calls = 0;
    const auth = fakeAuth({ pollMs: 3000 });
    const wrapped = {
      config: auth.config,
      signIn: (hooks) => {
        calls += 1;
        return auth.signIn(hooks);
      }
    };

    const session = new AuthSession(wrapped);
    await session.begin();
    await session.begin();

    assert.equal(calls, 1, 'concurrent begin() calls must share one flow');
  });
});

describe('AuthSession: browser (PKCE)', () => {
  test('returns the authorization URL promptly', async () => {
    const session = new AuthSession(fakeAuth({ mode: 'pkce', pollMs: 5000 }));

    const started = Date.now();
    const state = await session.begin();
    const elapsed = Date.now() - started;

    assert.ok(elapsed < 1000, `expected a prompt return, took ${elapsed}ms`);
    assert.equal(state.prompt.kind, 'browser');
    assert.match(state.prompt.url, /^https:\/\/login\.microsoftonline\.com\//);
  });
});

describe('AuthSession: failure handling', () => {
  test('records the error as state instead of throwing', async () => {
    const session = new AuthSession(fakeAuth({ pollMs: 20, fail: 'AADSTS7000218 boom' }));
    await session.begin();
    await session.promise;

    const snapshot = session.snapshot();
    assert.equal(snapshot.state, 'failed');
    assert.match(snapshot.error, /AADSTS7000218/);
  });

  test('a rejected flow does not surface as an unhandled rejection', async () => {
    const seen = [];
    const listener = (error) => seen.push(error);
    process.on('unhandledRejection', listener);

    const session = new AuthSession(fakeAuth({ pollMs: 10, fail: 'nope' }));
    await session.begin();
    await session.promise;
    await new Promise((resolve) => setTimeout(resolve, 50));

    process.off('unhandledRejection', listener);
    assert.equal(seen.length, 0, 'background sign-in failure must be contained');
  });

  test('a hung flow still returns within the announce timeout', async () => {
    // No callback is ever invoked and the flow never settles.
    const session = new AuthSession({
      config: { allowDeviceCode: false },
      signIn: () => new Promise(() => {})
    });

    const state = await Promise.race([
      session.begin(),
      new Promise((resolve) => setTimeout(() => resolve('never-returned'), 22000))
    ]);

    assert.notEqual(state, 'never-returned', 'begin() must not hang forever');
    assert.equal(state.state, 'pending');
  });
});
