/**
 * Compiles the logic modules and their tests to CommonJS and runs them with
 * Node's built-in test runner. The SPFx HTTP library only loads inside
 * SharePoint, so it is swapped for a small stand-in.
 */
'use strict';

const { execFileSync, spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const out = path.join(root, 'temp', 'tests');

fs.rmSync(out, { recursive: true, force: true });
execFileSync(process.execPath, [path.join(root, 'node_modules', 'typescript', 'bin', 'tsc'), '-p', 'tsconfig.test.json'], {
  cwd: root,
  stdio: 'inherit'
});

const tests = fs
  .readdirSync(path.join(out, 'tests'))
  .filter((f) => f.endsWith('.test.js'))
  .map((f) => path.join(out, 'tests', f));

const result = spawnSync(process.execPath, ['--require', path.join(root, 'tools', 'test-stubs.js'), '--test', ...tests], {
  cwd: root,
  stdio: 'inherit'
});
process.exit(result.status === null ? 1 : result.status);
