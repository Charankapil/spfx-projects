// Stands in for SPFx's themed-styles loader: registers a compiled .module.css at run time.
module.exports = function (source) {
  return `require('@microsoft/load-themed-styles').loadStyles(${JSON.stringify(source)});`;
};
