# Changelog

## 2.1.0

### Fixes
- **Cancel and error handling were broken in the deployed web part.** The
  SPFx build compiles to ES5, where subclasses of `Error` lose their type.
  So `instanceof` checks for a cancelled scan or an HTTP error were always
  false:
  - a cancel was recorded as library failures instead of stopping cleanly;
  - failed pages were not retried with smaller page sizes as designed;
  - throttling, timeouts and access denied all looked the same.

  Both error classes now restore their prototype. The unit tests run against
  the compiled `lib/` output, so this kind of problem is caught.
- **Failures no longer say "with your access".** Each failure now states its
  cause (throttling, server error or timeout, list view threshold, access
  denied, not found or network).
- **One unreadable item no longer stops the rest of its library.** Paging
  now walks through IDs itself. A batch that fails even at 500 items is
  skipped and recorded, and the rest of the library is still read.
- **"Time left" is based on the last minute of progress.** Early per-library
  overhead used to inflate the estimate.

### Faster
- Four libraries are read at once, with at most six requests in flight
  across the whole scan. Small libraries no longer wait for each other.
  Measured on a simulated site of 72 libraries at 300 ms per request:
  66 s → 17 s.
- New **Quick scan** (the default). A library nobody has changed for the
  chosen "Inactive after" period is measured as a whole from SharePoint's
  storage metrics: one request instead of reading every file. That drops
  the same simulated site to 12 s, with 30% fewer requests. The split
  stays exact up to that period. For longer periods, those libraries count
  by the date of their last change, and the dropdown marks those periods
  "(approx.)". Choose **Detailed** in the settings to read every file.
- The highest item ID is looked up only when it's needed (large libraries,
  or skipping a bad batch).

### Added
- **Scan issues** panel: every site and library that could not be read
  completely, with the reason.
- **Retry failed** button: re-reads only what failed and merges it into the
  saved result.
- **Where the site collection storage goes**: the site's storage figure
  split into inactive files, active files, version history (and how much
  of it is in dormant libraries), and everything else (recycle bins,
  unscanned hidden or skipped libraries, lists).
- A CSV column saying whether each library was measured as a whole, plus
  the library's last change.

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
