'use strict';
// Injects a compiled SPFx stylesheet into the page, with theme tokens replaced by their defaults.
module.exports = function (source) {
  const css = source.replace(/"?\[theme:\s*[^,\]]+,\s*default:\s*([^\]]+)\]"?/g, '$1');
  return `var s=document.createElement('style');s.textContent=${JSON.stringify(css)};document.head.appendChild(s);`;
};
