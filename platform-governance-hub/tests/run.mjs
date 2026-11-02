// Bundles the TypeScript unit tests with esbuild and runs them with node:test.
import { build } from 'esbuild';
import { readdirSync, mkdirSync, rmSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';

const here = dirname(fileURLToPath(import.meta.url));
const out = join(here, '..', 'temp', 'tests');
rmSync(out, { recursive: true, force: true });
mkdirSync(out, { recursive: true });
const entries = readdirSync(here).filter((f) => f.endsWith('.test.ts')).map((f) => join(here, f));
await build({ entryPoints: entries, bundle: true, platform: 'node', format: 'cjs', outdir: out, logLevel: 'warning', external: ['react', 'react-dom', '@fluentui/react'] });
const files = readdirSync(out).filter((f) => f.endsWith('.test.js')).map((f) => join(out, f));
const r = spawnSync(process.execPath, ['--test', ...files], { stdio: 'inherit' });
process.exit(r.status ?? 1);
