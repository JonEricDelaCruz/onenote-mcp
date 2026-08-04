#!/usr/bin/env node
/**
 * Portable test runner.
 *
 * `node --test` argument handling differs by version and platform, and the
 * differences are silent until CI turns red:
 *
 *   node --test "test/*.test.mjs"   works on Node 22+, but Node 20 has no glob
 *                                   support and treats it as a literal filename
 *   node --test test/*.test.mjs     relies on the SHELL to expand the glob, which
 *                                   Windows cmd.exe does not do
 *   node --test test/               behaves inconsistently across versions
 *
 * Passing an explicit list of files is the one form every supported version and
 * platform agrees on, so we build that list here.
 */

import { readdirSync } from 'node:fs';
import { spawn, spawnSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const testDir = path.join(root, 'test');

/**
 * `--syntax-only` mode: parse every source file without running it.
 *
 * The previous `npm run check` hardcoded two filenames and used shell `&&`, so
 * it silently skipped everything in src/ and behaved differently on Windows.
 */
if (process.argv.includes('--syntax-only')) {
  const targets = [
    path.join(root, 'onenote-mcp.mjs'),
    path.join(root, 'onenote-cli.mjs'),
    ...readdirSync(path.join(root, 'src'))
      .filter((n) => n.endsWith('.mjs'))
      .map((n) => path.join(root, 'src', n)),
    ...readdirSync(path.join(root, 'scripts'))
      .filter((n) => n.endsWith('.mjs'))
      .map((n) => path.join(root, 'scripts', n))
  ];

  let failed = 0;
  for (const file of targets) {
    const result = spawnSync(process.execPath, ['--check', file], { stdio: 'pipe' });
    const label = path.relative(root, file);
    if (result.status === 0) {
      console.log(`  ok    ${label}`);
    } else {
      failed += 1;
      console.error(`  FAIL  ${label}\n${result.stderr?.toString() ?? ''}`);
    }
  }
  console.log(`\n${targets.length - failed}/${targets.length} files parsed cleanly.`);
  process.exit(failed ? 1 : 0);
}

const files = readdirSync(testDir)
  .filter((name) => name.endsWith('.test.mjs'))
  .sort()
  .map((name) => path.join(testDir, name));

if (!files.length) {
  console.error('No test files found in test/');
  process.exit(1);
}

// Generous timeout: the auth-session suite deliberately exercises a hung
// sign-in, and Windows runners are slower than they look.
const args = ['--test', '--test-timeout=60000', '--test-concurrency=2', ...files];

console.log(`Running ${files.length} test files on Node ${process.version} (${process.platform})\n`);

const child = spawn(process.execPath, args, { stdio: 'inherit', cwd: root });
child.on('exit', (code) => process.exit(code ?? 1));
child.on('error', (error) => {
  console.error(error);
  process.exit(1);
});
