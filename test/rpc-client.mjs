/**
 * Talk to a real server process over stdio, without guessing how long it needs.
 *
 * Not a test file (the runner only collects `*.test.mjs`), just shared plumbing.
 *
 * WHY THIS EXISTS
 *
 * Every spawn-based test used to do the same thing: write a request, wait a
 * fixed 900-1200ms, close stdin, and parse whatever had arrived. That works on
 * a developer laptop and fails on CI, because the delay is a bet on how fast
 * the machine is. A GitHub Windows runner is a slow two-core VM, and Node takes
 * far longer to start there than on Linux or macOS. Lose the bet and stdout is
 * still empty, so `JSON.parse(undefined)` throws -- a confusing failure that
 * has nothing to do with the behaviour under test.
 *
 * So this waits for the response instead of for the clock. It resolves the
 * moment the expected number of complete JSON lines has arrived, which also
 * makes the suite faster everywhere: no test sits through a delay it did not
 * need.
 */

import { spawn } from 'node:child_process';

/**
 * @param {string} serverPath   path to the server entry point
 * @param {object[]} messages   JSON-RPC messages to send
 * @param {object} [options]
 * @param {object} [options.env]        environment for the child
 * @param {number} [options.expect]     how many responses to wait for. Defaults
 *   to the number of REQUESTS sent -- messages carrying an `id`. Notifications
 *   deliberately get no reply, so counting raw messages would wait forever.
 * @param {boolean} [options.waitForExit] also wait for the process to exit, so
 *   the caller can assert on its exit code
 * @param {number} [options.timeoutMs]  give up after this long
 * @returns {Promise<{responses: object[], stdout: string, stderr: string, code: number|null}>}
 */
export function converse(serverPath, messages, options = {}) {
  const {
    env = process.env,
    expect = messages.filter((m) => m && m.id !== undefined).length,
    waitForExit = false,
    timeoutMs = 30_000
  } = options;

  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [serverPath], {
      env,
      stdio: ['pipe', 'pipe', 'pipe']
    });

    const responses = [];
    let pending = '';
    let stdout = '';
    let stderr = '';
    let settled = false;

    const settle = (action, value) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      try {
        child.kill('SIGKILL');
      } catch {
        // Already gone; nothing to clean up.
      }
      action(value);
    };

    const result = (code) => ({ responses, stdout, stderr, code });

    const timer = setTimeout(() => {
      settle(
        reject,
        new Error(
          `Timed out after ${timeoutMs}ms waiting for ${expect} response(s)` +
            `${waitForExit ? ' and exit' : ''}; received ${responses.length}.` +
            `\nstderr:\n${stderr}`
        )
      );
    }, timeoutMs);

    child.stdout.on('data', (chunk) => {
      stdout += chunk;
      pending += chunk;

      // Keep the trailing fragment: a chunk boundary can land mid-line, and
      // parsing half a message would look like malformed output.
      const lines = pending.split('\n');
      pending = lines.pop() ?? '';

      for (const line of lines) {
        if (!line.trim()) continue;
        try {
          responses.push(JSON.parse(line));
        } catch {
          // Surfaced rather than swallowed: a test asserting stdout is clean
          // JSON needs to see this.
          responses.push({ unparsed: line });
        }
      }

      maybeFinish();
    });

    function maybeFinish() {
      if (responses.length < expect) return;

      // Some tests need the exit code -- a bad configuration must exit 1. For
      // those, stop writing and let the process end on its own terms.
      if (waitForExit) {
        child.stdin.end();
        return;
      }
      settle(resolve, result(null));
    }

    child.stderr.on('data', (chunk) => {
      stderr += chunk;
    });

    child.on('error', (error) => settle(reject, error));

    // A server that exits on its own (bad config, for instance) is a valid
    // outcome. Resolve with whatever it managed to say, and let the test judge.
    child.on('close', (code) => settle(resolve, result(code)));

    for (const message of messages) {
      child.stdin.write(`${JSON.stringify(message)}\n`);
    }

    // Nothing to wait for: the point of the call is how the process exits.
    if (expect === 0 && waitForExit) child.stdin.end();
  });
}

/** The `_meta` block a 2026-07-28 stateless client sends on every call. */
export const META = {
  'io.modelcontextprotocol/protocolVersion': '2026-07-28',
  'io.modelcontextprotocol/clientInfo': { name: 'test', version: '1' },
  'io.modelcontextprotocol/clientCapabilities': {}
};

/** Convenience: list the server's tools. */
export async function listTools(serverPath, env) {
  const { responses, stderr } = await converse(
    serverPath,
    [{ jsonrpc: '2.0', id: 1, method: 'tools/list', params: { _meta: META } }],
    { env, expect: 1 }
  );

  const tools = responses[0]?.result?.tools;
  if (!tools) {
    throw new Error(
      `Server returned no tool list.\nresponses: ${JSON.stringify(responses)}\nstderr:\n${stderr}`
    );
  }
  return tools;
}
