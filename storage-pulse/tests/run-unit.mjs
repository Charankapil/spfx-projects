// Bundles each unit test with esbuild and runs them with node:test. The code under test is the
// SPFx build's own output (lib/, ES5), not src/, so the tests see exactly what ships - including
// ES5 quirks such as Error subclasses losing their prototype. Run `gulp bundle` first.
import * as esbuild from 'esbuild';
import { spawnSync } from 'child_process';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { compiledLib, PROJECT, spfxStubs } from './plugins.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
if (!fs.existsSync(path.join(PROJECT, 'lib/webparts/storagePulse/services/StorageScanService.js'))) {
  console.error('lib/ not found: run "npx gulp bundle --ship" in the project folder first.');
  process.exit(1);
}
const out = path.join(here, '.out');
fs.rmSync(out, { recursive: true, force: true });
const tests = fs.readdirSync(path.join(here, 'unit')).filter((f) => f.endsWith('.test.ts'));
await esbuild.build({
  entryPoints: tests.map((t) => path.join(here, 'unit', t)),
  outdir: out,
  outExtension: { '.js': '.mjs' },
  bundle: true,
  platform: 'node',
  format: 'esm',
  target: 'node20',
  nodePaths: [path.join(PROJECT, 'node_modules')],
  plugins: [spfxStubs(), compiledLib()],
  logLevel: 'warning'
});
const result = spawnSync(process.execPath, ['--test', ...tests.map((t) => path.join(out, t.replace(/\.ts$/, '.mjs')))], {
  stdio: 'inherit'
});
process.exit(result.status === null ? 1 : result.status);
