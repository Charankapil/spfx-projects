# Changelog

## 2.0.0 — Storage Pulse

Renamed from Storage Activity Analyser. The solution ID and web part ID are
unchanged, so existing pages and saved scans keep working (see *Upgrading*
in the README).

### Security
- The saved results file (in Site Assets, often editable by members) is now
  treated as untrusted. Links from it are only rendered when they're `https`
  on the tenant's own host. A tampered `javascript:` URL could previously
  run script when clicked.
- Every value in the saved file is checked, and sizes are capped. A damaged
  or edited file shows a warning instead of crashing the web part.
- "Last scanned by" now comes from SharePoint's record of who last modified
  the file, not from the file's contents.

### Fixes
- Cancel now stops the scan immediately. Before, it could wait up to 5
  minutes during a throttling back-off.
- Throttling retries and network retries use separate counters. Before, one
  network blip after a few throttled attempts failed the request.
- SharePoint's error messages are read from `nometadata` responses. Before,
  they appeared as raw JSON.
- The progress bar compares items with items, so it now reaches 100%.
- Links to files and libraries with `#` or `%` in their names now work.

### Added
- New settings:
  - include hidden libraries, such as the Preservation Hold Library;
  - skip system libraries;
  - libraries to skip;
  - who can run a scan;
  - warn when results are older than N days.
- A stale-results banner.
- A dark variant for dark section backgrounds and dark themes. Microsoft
  Teams tab and personal app support.
- A new look:
  - animated logo, and aurora header;
  - ring gauge;
  - live scan panel with radar, steps, speed and time left;
  - glass cards, staggered chart animations and loading skeletons.
  All animation respects "reduce motion".
- All on-screen text moved to `loc/en-us.js` for translation.
- Screen-reader progress announcements, and focusable warning icons.
- An error boundary.
- A logo, App Catalog icon, toolbox icon, screenshot and store categories.
- A test suite (unit, scan engine, browser) and a GitHub Actions workflow.

### Changed
- By default only site owners and site collection admins can run a scan.
  v1 let anyone scan for themselves.
- Package file name: `storage-pulse.sppkg`.

## 1.1.0
- Storage by file type and by category, with a CSV export.

## 1.0.1
- Parallel reading of very large libraries, leaner responses, longer
  throttling tolerance, next-link fallback, reporting of items hidden by
  permissions.

## 1.0.0
- First release as Storage Activity Analyser.
