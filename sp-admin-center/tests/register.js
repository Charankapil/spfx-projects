// Lets plain `node` load the TypeScript sources (transpile only; type-checking is done by the gulp build).
const ts = require('typescript');
const fs = require('fs');
const path = require('path');
const Module = require('module');

require.extensions['.ts'] = (module, filename) => {
  const out = ts.transpileModule(fs.readFileSync(filename, 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: process.env.TEST_TARGET === 'es5' ? ts.ScriptTarget.ES5 : ts.ScriptTarget.ES2019, downlevelIteration: true, esModuleInterop: true },
    fileName: filename
  });
  module._compile(out.outputText, filename);
};

const resolve = Module._resolveFilename;
Module._resolveFilename = function (request, ...rest) {
  if (request === '@microsoft/sp-http') {
    return path.join(__dirname, 'mocks', 'sp-http.js');
  }
  return resolve.call(this, request, ...rest);
};
