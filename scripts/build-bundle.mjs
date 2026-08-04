#!/usr/bin/env node
/**
 * Build the .mcpb bundle for one-click install in Claude Desktop.
 *
 * A bundle must be self-contained: Claude Desktop unpacks it and runs the entry
 * point with its own bundled Node, without an npm install. So we stage the
 * source plus a production-only dependency tree, then pack.
 *
 * Staging in a temp directory (rather than packing the repo in place) keeps dev
 * dependencies, tests, .env files, and the token cache out of the artifact --
 * see the exclusion assertions at the end.
 */

import { execFileSync } from 'node:child_process';
import {
  mkdtempSync,
  rmSync,
  cpSync,
  existsSync,
  readFileSync,
  writeFileSync,
  mkdirSync,
  readdirSync,
  statSync
} from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const distDir = path.join(root, 'dist');

/** Files and directories that make up the runtime. Everything else is excluded. */
const INCLUDE = ['onenote-mcp.mjs', 'onenote-cli.mjs', 'src', 'manifest.json', 'icon.png', 'README.md', 'LICENSE'];

/** Nothing matching these may ever appear in the artifact. */
const FORBIDDEN = [
  '.env',
  '.access-token.txt',
  'token-cache.json',
  '.git',
  'test',
  '.github',
  'scripts',
  '.test-cache'
];

function run(command, args, cwd) {
  return execFileSync(command, args, { cwd, stdio: 'pipe', encoding: 'utf8' });
}

const pkg = JSON.parse(readFileSync(path.join(root, 'package.json'), 'utf8'));
const manifest = JSON.parse(readFileSync(path.join(root, 'manifest.json'), 'utf8'));

// The two version numbers are user-visible in different places; drift between
// them is confusing, so treat it as a build error.
if (manifest.version !== pkg.version) {
  console.error(
    `Version mismatch: package.json is ${pkg.version} but manifest.json is ${manifest.version}.\n` +
      'Update both, then rebuild.'
  );
  process.exit(1);
}

const stage = mkdtempSync(path.join(tmpdir(), 'onenote-mcpb-'));
console.log(`Staging in ${stage}`);

try {
  // 1. Copy runtime files.
  for (const entry of INCLUDE) {
    const from = path.join(root, entry);
    if (!existsSync(from)) {
      console.error(`Missing required file: ${entry}`);
      process.exit(1);
    }
    cpSync(from, path.join(stage, entry), { recursive: true });
  }

  // 2. A package.json without scripts or devDependencies. Scripts are stripped
  //    so nothing can execute during the consumer-side install.
  writeFileSync(
    path.join(stage, 'package.json'),
    `${JSON.stringify(
      {
        name: pkg.name,
        version: pkg.version,
        description: pkg.description,
        type: 'module',
        main: pkg.main,
        license: pkg.license,
        dependencies: pkg.dependencies
      },
      null,
      2
    )}\n`
  );
  cpSync(path.join(root, 'package-lock.json'), path.join(stage, 'package-lock.json'));

  // 3. Production dependency tree, reproducible from the lockfile, with all
  //    lifecycle scripts disabled.
  console.log('Installing production dependencies...');
  run('npm', ['ci', '--omit=dev', '--ignore-scripts', '--no-audit', '--no-fund'], stage);

  // 4. Pack.
  mkdirSync(distDir, { recursive: true });
  const outFile = path.join(distDir, `onenote-mcp-${pkg.version}.mcpb`);
  console.log('Packing bundle...');
  const packOutput = run('npx', ['--yes', '--package=@anthropic-ai/mcpb', 'mcpb', 'pack', stage, outFile], root);
  console.log(packOutput.trim());

  // 5. Verify nothing sensitive or extraneous made it in.
  //
  //    `mcpb info` truncates its file listing, so checking its output would
  //    silently pass. Unpack the real archive and inspect every entry.
  const verifyDir = mkdtempSync(path.join(tmpdir(), 'onenote-verify-'));
  try {
    run('npx', ['--yes', '--package=@anthropic-ai/mcpb', 'mcpb', 'unpack', outFile, verifyDir], root);

    const entries = [];
    const walk = (dir, prefix = '') => {
      for (const item of readdirSync(dir, { withFileTypes: true })) {
        const rel = prefix ? `${prefix}/${item.name}` : item.name;
        entries.push(rel);
        if (item.isDirectory()) walk(path.join(dir, item.name), rel);
      }
    };
    walk(verifyDir);

    const leaked = entries.filter((entry) => {
      const segments = entry.split('/');
      // Ignore anything inside node_modules: a dependency legitimately having a
      // file called "test" is not our secret leaking.
      if (segments.includes('node_modules')) return false;
      return FORBIDDEN.some((name) => segments.includes(name));
    });

    if (leaked.length) {
      console.error(`\nRefusing to ship. Bundle contains excluded paths:\n  ${leaked.join('\n  ')}`);
      rmSync(outFile, { force: true });
      process.exit(1);
    }

    // The entry point must actually be runnable from the unpacked layout.
    for (const required of ['onenote-mcp.mjs', 'manifest.json', 'icon.png', 'src', 'node_modules']) {
      if (!existsSync(path.join(verifyDir, required))) {
        console.error(`\nRefusing to ship: unpacked bundle is missing ${required}`);
        rmSync(outFile, { force: true });
        process.exit(1);
      }
    }

    const fileCount = entries.length;
    const sizeMb = (statSync(outFile).size / 1024 / 1024).toFixed(1);
    console.log(`\nVerified: ${fileCount} entries, no excluded paths.`);
    console.log(`Built ${path.relative(root, outFile)} (${sizeMb} MB)`);
    console.log('Install by double-clicking it, or drag it onto the Claude Desktop window.');
  } finally {
    rmSync(verifyDir, { recursive: true, force: true });
  }
} finally {
  rmSync(stage, { recursive: true, force: true });
}
