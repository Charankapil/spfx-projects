// Writes the importable Power Automate packages into releases/ (with a placeholder site URL;
// the web part's "Monthly schedule" tab produces the same packages with the real one filled in).
require('../tests/register');
const fs = require('fs');
const path = require('path');
const pkg = require('../src/webparts/accessReport/flow/package');
const { buildDefinition, PLACEHOLDER_URL } = require('../src/webparts/accessReport/flow/definition');

const out = path.join(__dirname, '..', 'releases');
fs.mkdirSync(out, { recursive: true });
fs.writeFileSync(path.join(out, pkg.LEGACY_FILE), Buffer.from(pkg.legacyPackage(PLACEHOLDER_URL)));
fs.writeFileSync(path.join(out, pkg.SOLUTION_FILE), Buffer.from(pkg.solutionPackage(PLACEHOLDER_URL)));
fs.writeFileSync(path.join(__dirname, 'definition.json'), JSON.stringify(buildDefinition({ homeUrl: PLACEHOLDER_URL }), null, 2) + '\n');
console.log(`Wrote releases/${pkg.LEGACY_FILE}, releases/${pkg.SOLUTION_FILE} and flow/definition.json`);
