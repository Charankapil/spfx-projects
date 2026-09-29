// Bundles each unit test (TypeScript, straight from src/) with esbuild and runs them with node:test.
import * as esbuild from 'esbuild';
import { spawnSync } from 'child_process';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { PROJECT, spfxStubs } from './plugins.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
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
  plugins: [spfxStubs()],
  logLevel: 'warning'
});
const result = spawnSync(process.execPath, ['--test', ...tests.map((t) => path.join(out, t.replace(/\.ts$/, '.mjs')))], {
  stdio: 'inherit'
});
process.exit(result.status === null ? 1 : result.status);
