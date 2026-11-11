'use strict';
const path = require('path');

const root = path.join(__dirname, '..', '..');

module.exports = {
  mode: 'development',
  devtool: false,
  entry: path.join(__dirname, 'harness.js'),
  output: { path: path.join(root, 'temp', 'harness'), filename: 'harness.bundle.js', hashFunction: 'sha256' },
  resolve: {
    alias: {
      '@microsoft/sp-http': path.join(__dirname, 'stubs', 'sp-http.js'),
      '@microsoft/sp-page-context': path.join(__dirname, 'stubs', 'sp-page-context.js'),
      '@microsoft/sp-application-base': path.join(__dirname, 'stubs', 'sp-application-base.js')
    }
  },
  module: { rules: [{ test: /\.css$/, use: [path.join(__dirname, 'css-loader.js')] }] },
  performance: { hints: false }
};
