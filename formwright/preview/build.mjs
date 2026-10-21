// Bundles the compiled web part code (lib/, from `gulp build`) with a fake
// SharePoint backend into preview/out/index.html. Run with `npm run preview`.
import * as esbuild from 'esbuild';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '..');
const out = path.join(here, 'out');

if (!fs.existsSync(path.join(root, 'lib'))) {
  console.error('Run "npx gulp build" first: the preview bundles the compiled code in lib/.');
  process.exit(1);
}

await esbuild.build({
  entryPoints: [path.join(here, 'entry.js')],
  bundle: true,
  outfile: path.join(out, 'bundle.js'),
  nodePaths: [path.join(root, 'node_modules')],
  alias: { '@microsoft/sp-http': path.join(here, 'sp-http-stub.js'), '@fs': path.join(root, 'lib') },
  // SPFx already scoped the class names; keep them as they are.
  loader: { '.css': 'global-css', '.module.css': 'global-css' },
  define: { 'process.env.NODE_ENV': '"development"' },
  logLevel: 'warning'
});

// SharePoint replaces theme tokens at runtime; use their defaults here.
const cssFile = path.join(out, 'bundle.css');
fs.writeFileSync(cssFile, fs.readFileSync(cssFile, 'utf8').replace(/["']\[theme:\w+, ?default: ?(#[0-9a-fA-F]+)\]["']/g, '$1'));

fs.mkdirSync(path.join(out, 'fonts'), { recursive: true });
const fonts = path.join(root, 'node_modules/@fluentui/font-icons-mdl2/fonts');
fs.readdirSync(fonts).forEach((f) => fs.copyFileSync(path.join(fonts, f), path.join(out, 'fonts', f)));

fs.writeFileSync(
  path.join(out, 'index.html'),
  `<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Formwright preview</title><link rel="stylesheet" href="bundle.css"><style>body{margin:0;padding:24px;font-family:'Segoe UI',system-ui,sans-serif;background:#f3f2f1}</style></head><body><div id="root"></div><script src="bundle.js"></script></body></html>`
);
console.log(`Preview written to ${path.relative(process.cwd(), path.join(out, 'index.html'))}`);
