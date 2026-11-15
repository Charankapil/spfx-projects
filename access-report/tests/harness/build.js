// Builds tests/harness/dist/bundle.js from lib/ (run `npm run build` first).
const webpack = require('webpack');
const config = require('./webpack.config');

webpack(config, (err, stats) => {
  if (err || stats.hasErrors()) {
    console.error(err || stats.toString({ colors: false, all: false, errors: true }));
    process.exit(1);
  }
  console.log(stats.toString({ colors: false, all: false, assets: true, warnings: true }));
});
