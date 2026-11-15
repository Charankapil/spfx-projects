// Bundles the built web part (lib/) with the fake tenant so the real UI can be driven in a
// browser without SharePoint. Used by tests/ui.test.js; not part of the shipped package.
const path = require('path');

module.exports = {
  mode: 'development',
  devtool: false,
  entry: path.join(__dirname, 'main.js'),
  output: { path: path.join(__dirname, 'dist'), filename: 'bundle.js' },
  resolve: { alias: { '@microsoft/sp-http': path.join(__dirname, '..', 'mocks', 'sp-http.js') } },
  module: { rules: [{ test: /\.css$/, use: path.join(__dirname, 'themed-css-loader.js') }] },
  performance: { hints: false }
};
