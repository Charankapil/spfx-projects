'use strict';
// Bundles the test page with the webpack that SPFx already ships (no extra install).
const webpack = require('webpack');
const config = require('./webpack.config');

webpack(config, (err, stats) => {
  if (err) {
    console.error(err);
    process.exit(1);
  }
  const info = stats.toJson({ all: false, errors: true, warnings: true });
  info.errors.forEach((e) => console.error(String(e).substring(0, 1500)));
  if (stats.hasErrors()) {
    process.exit(1);
  }
  console.log(`Test page bundled (${info.warnings.length} warnings).`);
});
