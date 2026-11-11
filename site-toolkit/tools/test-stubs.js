'use strict';

const Module = require('module');
const path = require('path');

const stub = path.join(__dirname, 'sp-http-stub.js');
const resolve = Module._resolveFilename;
Module._resolveFilename = function (request, ...rest) {
  return request === '@microsoft/sp-http' ? stub : resolve.call(this, request, ...rest);
};

// Browser storage, as far as the cache module needs it.
function memoryStorage() {
  const data = new Map();
  return {
    getItem: (key) => (data.has(key) ? data.get(key) : null),
    setItem: (key, value) => data.set(key, String(value)),
    removeItem: (key) => data.delete(key),
    clear: () => data.clear()
  };
}
global.window = global;
global.localStorage = memoryStorage();
global.sessionStorage = memoryStorage();
