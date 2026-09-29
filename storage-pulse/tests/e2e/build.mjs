import * as esbuild from 'esbuild';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { cssModules, PROJECT, spfxStubs } from '../plugins.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
export async function buildPage(dark) {
  const out = path.join(here, 'out');
  fs.mkdirSync(out, { recursive: true });
  await esbuild.build({
    entryPoints: [path.join(here, 'entry.js')],
    bundle: true,
    outfile: path.join(out, 'bundle.js'),
    format: 'iife',
    nodePaths: [path.join(PROJECT, 'node_modules')],
    define: { 'process.env.NODE_ENV': '"production"' },
    plugins: [spfxStubs(), cssModules()],
    logLevel: 'warning'
  });
  fs.writeFileSync(
    path.join(out, 'index.html'),
    '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">' +
      '<style>body{margin:0;background:#faf9f8;font-family:"Segoe UI",sans-serif}body.dark{background:#0f1426}' +
      '#root{max-width:1180px;margin:24px auto;padding:0 16px}</style></head>' +
      '<body><div id="root"></div><script src="bundle.js"></script></body></html>'
  );
  return path.join(out, 'index.html');
}
