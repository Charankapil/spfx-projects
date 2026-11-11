#!/usr/bin/env bash
# Rebuilds everything the browser checks depend on, then runs them.
# The browser page is built from lib/, so the SPFx build must run first.
set -euo pipefail
cd "$(dirname "$0")/../.."
npx gulp build
node tools/harness/build.js
node tools/harness/run.js "$@"
