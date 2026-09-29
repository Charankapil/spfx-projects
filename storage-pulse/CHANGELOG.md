# Changelog

## 2.2.0

### Fixes
- **The scan could contribute to tenant-wide throttling.** It ran up to six
  requests at once with almost no gap, at the same pace all the time. When
  SharePoint answered `429`, only that request waited; the others kept
  sending and were refused too. Against a simulated tenant that throttles
  above 8 requests a second, v2.1.1 sent 26 requests during throttling
  holds; v2.2 sends none, with a third of the throttles.
- **Throttling made libraries "fail" and sent you to Retry failed.** After
  12 throttled retries a request gave up and the library was marked failed.
  Throttling is now a pause, never a failure.

### Added
- **Scan speed** setting: Gentle (default), Balanced or Fast. Gentle is
  about three to four times slower than v2.1.1; see the README.
- **One request governor for the whole scan.** It spaces requests, caps
  requests in flight, holds everything on `Retry-After` (with jitter),
  halves the allowed concurrency on each throttle and raises it again
  after a run of successes.
- **Pause and resume.** After an hour of accumulated throttling the scan
  pauses with its progress saved. Owners' progress is also saved about every
  minute as the scan runs. **Resume scan** reads only the libraries still
  unread, and an unfinished scan is offered on the next visit (with
  **Discard**). Ages stay measured from the original scan start.
- **Throttling shown as a calm status** in the scan panel: countdown, reduced
  speed and how many times the scan was slowed, instead of an error.
- **Automatic second pass** for libraries that failed with a server error,
  timeout or network drop, after a 15-second cool-down.
- A test hook and a simulated throttling tenant (a rate limit, and an
  outage) in the test suite, including a check that nothing is sent while
  a hold is active.

## 2.1.1

### Fixes
- **HTTP 406 ("Not Acceptable") on item queries abandoned whole
  libraries.** SharePoint rejects the format of a request without saying
  which part it dislikes, and it did so on the first item request of many
  libraries in one tenant (54 of 181). The scan treated any 4xx as final, so
  those libraries were skipped entirely and the totals were understated.
  Item queries now fall back in order:
  1. `$filter` on `Id` plus `$orderby=Id`, lean `Accept` header (fastest);
  2. the same without `$orderby`;
  3. the same with SharePoint's default `Accept` header.

  The working form is learned within a few libraries, so the wasted
  requests stay a small constant. The dashboard notes how many libraries
  needed a fallback. A 406 that survives all three forms is reported as a
  real failure with a clear message.
- The highest-ID lookup for large libraries also retries with the default
  header, and otherwise reads the library as one range instead of failing.

**After upgrading:** open the page and click **Retry failed**. It re-reads
only the libraries that failed in the last scan.

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
