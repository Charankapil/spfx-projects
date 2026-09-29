import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const here = path.dirname(fileURLToPath(import.meta.url));
export const PROJECT = path.resolve(here, '..');

/** esbuild plugin: stand-ins for SPFx runtime modules, and the AMD strings file loaded like SharePoint does. */
export function spfxStubs() {
  return {
    name: 'spfx-stubs',
    setup(build) {
      build.onResolve({ filter: /^@microsoft\/sp-http$/ }, () => ({ path: path.join(here, 'stubs/sp-http.js') }));
      build.onResolve({ filter: /^@microsoft\/sp-page-context$/ }, () => ({ path: path.join(here, 'stubs/sp-page-context.js') }));
      build.onResolve({ filter: /^@microsoft\/sp-webpart-base$/ }, () => ({ path: 'webpart-base', namespace: 'empty' }));
      build.onLoad({ filter: /.*/, namespace: 'empty' }, () => ({ contents: 'export {};', loader: 'js' }));
      build.onResolve({ filter: /^StoragePulseWebPartStrings$/ }, () => ({ path: 'strings', namespace: 'strings' }));
      build.onLoad({ filter: /.*/, namespace: 'strings' }, () => {
        const amd = fs.readFileSync(path.join(PROJECT, 'src/webparts/storagePulse/loc/en-us.js'), 'utf8');
        return {
          contents: `let s; (function (define) { ${amd} })(function (d, f) { s = f(); }); module.exports = s;`,
          loader: 'js'
        };
      });
    }
  };
}

/** Compiled CSS modules (lib/*.module.css) injected into the page with SharePoint's theme tokens resolved to defaults. */
export function cssModules() {
  return {
    name: 'css-modules',
    setup(build) {
      build.onLoad({ filter: /\.module\.css$/ }, (args) => {
        const css = fs.readFileSync(args.path, 'utf8').replace(/"\[theme:\w+, default:([^\]]+)\]"/g, '$1');
        return {
          contents: `const s = document.createElement('style'); s.textContent = ${JSON.stringify(css)}; document.head.appendChild(s);`,
          loader: 'js'
        };
      });
      build.onResolve({ filter: /\.png$/ }, (args) => ({ path: path.resolve(args.resolveDir, args.path), namespace: 'png' }));
      build.onLoad({ filter: /.*/, namespace: 'png' }, () => ({ contents: 'export default "";', loader: 'js' }));
    }
  };
}

/** Points imports of src/webparts/** at the compiled lib/webparts/**.js from the SPFx build. */
export function compiledLib() {
  const src = path.join(PROJECT, 'src', 'webparts') + path.sep;
  const lib = path.join(PROJECT, 'lib', 'webparts') + path.sep;
  return {
    name: 'compiled-lib',
    setup(build) {
      build.onResolve({ filter: /src\/webparts\// }, (args) => {
        const absolute = path.resolve(args.resolveDir, args.path).replace(/\.tsx?$/, '');
        if (absolute.indexOf(src) !== 0) {
          return undefined;
        }
        return { path: lib + absolute.substring(src.length) + '.js' };
      });
    }
  };
}
