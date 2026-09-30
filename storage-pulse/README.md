<p align="center">
  <img src="assets/logo/storage-pulse-lockup.png" alt="Storage Pulse: see what's alive in your SharePoint storage" width="560">
</p>

**How much of a SharePoint site's storage is in files people still use?**
Storage Pulse is one `.sppkg`. It needs no Azure app registration and no
API permissions, and runs as the person viewing the page.

![Storage Pulse dashboard](docs/images/dashboard.png)
<sub>Sample data from a simulated site collection.</sub>

SharePoint Online shows how much storage a site uses. It doesn't show how
much of that storage is in files nobody has touched for years. Storage
Pulse adds up the size of every file by when it was last modified and
splits the total into:

- **active** files, modified within the period you choose;
- **inactive** files, not modified for that long or longer.

For example: *626 GB inactive, not modified for 1 year or more; 204 GB
active.* Owners can then decide what to archive, move or delete.

## What it shows

- **Where the storage goes.** The site collection's storage figure split
  into inactive files, active files, version history, and the rest
  (recycle bins, unscanned or hidden libraries, lists). Version history is
  often most of a site's storage.
- **The split.** A ring gauge with the inactive share, plus size, share
  and file count for active files, inactive files and all files. The site
  collection storage figure SharePoint reports is shown alongside.
- **Insights in plain language.** How much space a clean-up would free,
  which libraries are dormant, how much version history adds, and anything
  that couldn't be read.
- **Inactive after: 3 or 6 months, or 1, 2, 3 or 5 years.** Changing it
  updates every figure instantly, with no new scan. The default comes from
  the web part settings.
- **Storage by time since last modified.** Bars from under 3 months to 10+
  years, split into active and inactive at a glowing threshold line.
- **Storage by file type** (`.pdf`, `.mp4`, `.zip` …) or **by category**
  (Video & audio, Images, Archives, PDF, Office documents, Design & CAD,
  Email, Databases & backups …). Each bar is split into inactive and active.
- **Where the inactive storage is.** A table by library or by site. It
  marks **Dormant** libraries (nothing changed within the period) and shows
  version history size where SharePoint reports it.
- **Largest inactive files**, with links.
- **Scan issues and Retry failed.** Every site or library that couldn't be
  read completely is listed with the reason: throttling, timeouts, a list
  view threshold, access denied and so on. **Retry failed** re-reads only
  those and merges them in.
- **CSV exports.** One export each for libraries, file types (per library)
  and the largest inactive files.
- **Saved for everyone.** A site owner's scan is saved in Site Assets.
  Anyone who opens the page sees it straight away, with a warning once it
  is older than a set number of days.

It follows the site's theme colour, works on dark section backgrounds and
in Microsoft Teams (light, dark and high contrast), and respects "reduce
motion".

## Permissions and security

Every call goes through SPFx's built-in `SPHttpClient` with the signed-in
user's existing SharePoint session.

- There's no Azure AD / Entra ID app registration, client ID or secret.
- The package requests **no API permissions**, so there's nothing for an
  admin to approve under *API access*.
- There's no elevated or app-only access, and no data leaves SharePoint.
- It only calls `_api/*` endpoints on the site collection it's added to.

| Who opens the page | What they can do |
| --- | --- |
| Site owner / site collection admin | Run a scan. The result is saved for everyone. |
| Member / visitor | See the last saved scan. They can run their own scan only if an owner allows it in the settings, and that scan is shown only to them. |

Results are security-trimmed: the scan only counts what the person running
it can open. Items hidden by item-level permissions are counted and
reported, but not included in the totals. Run the scan as a site
collection admin for the complete picture.

**Hardening in v2.**
- **The saved results file is treated as untrusted when read.** It sits in
  Site Assets, which members can usually edit. Every value is checked and
  sizes are capped.
- **Only safe links become links.** Only `https` URLs on the tenant's own
  SharePoint host are rendered as links. Anything else, such as
  `javascript:` or another domain, is shown as plain text.
- **"Last scanned by" comes from SharePoint.** It uses SharePoint's own
  record of who last modified the file, not the file's contents.
- **A broken or edited file can't take the page down.** It produces a
  warning, and a rendering error shows a message instead of breaking the
  page.

### What is stored, and where

After an owner's scan, one JSON file is written to the **Site Assets**
library of the site that hosts the page:
`storage-pulse-site-collection.json`, or `storage-pulse-site.json` for the
"this site and its subsites" scope. The file contains:

- site and library names and URLs;
- file counts and sizes per month of age and per file type;
- the names and paths of up to 200 of the largest files older than 3
  months;
- who ran the scan, when, and the scan settings.

It contains no file contents. Anyone who can read Site Assets can read it.
Nothing is stored anywhere else, and nothing is sent outside SharePoint.

## Settings (web part property pane)

| Group | Setting | Default |
| --- | --- | --- |
| General | Web part title | Storage Pulse |
| | Default inactivity period | 1 year |
| | Warn when results are older than (days, 0 = never) | 90 |
| What to scan | Scan: whole site collection, or this site and its subsites | Whole site collection |
| | How to scan: Quick, or Detailed (read every file) | Quick |
| | Scan speed: Gentle, Balanced or Fast | Gentle |
| | Include hidden libraries, such as the Preservation Hold Library | Off |
| | Skip Site Pages, Site Assets, Style Library and Form Templates | Off |
| | Libraries to skip (one title or URL name per line) | — |
| Who can scan | Site owners only, only the people you name (by email), or everyone (their scans are shown only to them) | Site owners |

`_catalogs` libraries (master pages, web parts, solutions) are never
scanned.

## How it works

1. **Find the sites and libraries.**
   `_api/web/getsubwebsfilteredforcurrentuser(...)` returns only the
   subsites the user can open. `_api/web/lists?$filter=BaseType eq 1`
   returns the document libraries, filtered by the settings above.
2. **Read each file's size and last-modified date.** It pages through
   `_api/web/lists(guid'…')/items?$select=Id,FSObjType,Modified,FileRef,File/Length&$expand=File&$top=5000`
   with `odata.metadata=nometadata` and follows SharePoint's next-page
   link. Paging by ID works at any library size, with no list view
   threshold error. Only numbers are kept:
   - a count and byte total per month of age;
   - the same per file type and age band;
   - the 200 largest files older than 3 months.

   Browser memory therefore stays flat however many files there are.
3. **Version history.** For each library, `RootFolder/StorageMetrics`
   gives SharePoint's own figure: `TotalSize − TotalFileStreamSize` is the
   version history. If the user can't read it, that column shows "–".
4. **Site storage.** `_api/site?$select=Usage` gives the site collection
   total for context.

### Quick and detailed scans

Reading files is the slow part: SharePoint returns at most 5,000 per
request.

- **Quick scan (the default)** skips that for libraries nobody has changed
  for the chosen "Inactive after" period. SharePoint records the last time
  a person changed anything in a list (`LastItemUserModifiedDate`), and
  every folder's storage metrics carry the total file count, current size,
  size with versions and last change. If both are older than the period,
  every file in the library must be at least that old. The library is then
  measured with one request.
  - The split is exact up to that period.
  - For longer periods, the dropdown marks the option "(approx.)" and those
    libraries count by the date of their last change.
  - They have no file-type breakdown and don't appear in the largest-files
    list.
- **Detailed scan** reads every file in every library.

How many libraries are read at once, and how many requests are in flight,
depends on the **scan speed** (see the next section).

### Throttling: being a good neighbour

SharePoint Online throttles by resource units, per user and per app and
also per tenant over a few minutes, with limits that grow with the number of
licences. One busy scan can therefore slow down other people and tools on
the same tenant, not only itself. Storage Pulse is built to stay well
clear of that, and to treat throttling as a pause and never as a failure.

**Scan speed** (a setting; the default is *Gentle*):

| Speed | Requests in flight | Gap between requests | Libraries at once |
| --- | --- | --- | --- |
| Gentle (default) | 2 | 400 ms | 2 |
| Balanced | 4 | 200 ms | 3 |
| Fast | 6 | 100 ms | 4 |

Measured on a simulated 72-library site at 300 ms per request:

| | Time | Requests per second (peak in flight) |
| --- | --- | --- |
| v2.1.1 | 17 s | 8.6 (4) |
| v2.2 Fast | 20 s | 7.4 (4) |
| v2.2 Balanced | 35 s | 4.1 (2) |
| v2.2 Gentle | 68 s | 2.1 (1) |
| v2.2 Gentle, quick scan | 50 s | 2.0 (1) |

Gentle is about three to four times slower than v2.1.1, and Fast is
what v2.1 did. Start with Gentle. If the tenant stays calm, switch to
Balanced, and use Fast only outside working hours. Real timings depend on
how long SharePoint takes per page.

![The scan panel while SharePoint has asked the scan to slow down](docs/images/scan-throttled.png)

**What the scan does when SharePoint throttles it:**

- **One shared brake.** Every request goes through a single governor. A
  `429` or `503` puts *all* requests on hold until `Retry-After` has passed
  (with a little random jitter), instead of only the request that was
  refused while the others carry on. Against a simulated tenant that
  throttles above 8 requests a second, v2.1.1 sent 26 requests while a hold
  was active, and v2.2 sent none.
- **Slows down, then recovers.** Each throttle halves the requests allowed
  at once and widens the gap; a run of successful requests raises them
  again, one step at a time.
- **Throttling never fails a library.** The old limit of 12 retries is
  gone. The scan panel says SharePoint asked it to slow down, shows the
  countdown, and carries on by itself.
- **Pause and resume.** If SharePoint has kept the scan waiting for more
  than an hour in total, it stops with a warning instead of an error. Its
  progress is saved and the page offers **Resume scan**, which reads only
  what is left.
- **Progress is saved as it goes** (site owners), roughly once a minute, in
  a separate file next to the results. Closing the tab, or a browser crash,
  costs at most the library being read at that moment. On the next visit an
  owner sees "An unfinished scan from … is saved" with **Resume scan** and
  **Discard**. Visitors never see it, and the last complete scan stays
  visible to everyone meanwhile.
- **Automatic second pass.** A library that failed for a reason that
  often passes on its own (server error, timeout, network drop) is tried
  once more, alone, after a 15-second cool-down, before it is reported.
  **Retry failed** is then only needed for real failures.

What it can't do: the `RateLimit-*` headers, which would let it slow down
*before* being throttled, are only returned to apps using application
permissions, and this web part runs as the signed-in user. A browser also
can't set the `User-Agent` header that Microsoft recommends for
identifying an app's traffic. So the scan reacts to a `429` instead of
predicting it, and keeps its normal pace low so that it rarely gets one.

If your tenant is already being throttled, stop running scans for now.
Throttling windows are short, usually minutes. If it lasts for hours,
something other than this web part is using the tenant's quota: look at
which apps are calling SharePoint, in the Microsoft 365 admin center.

### If SharePoint answers 406 "Not Acceptable"

SharePoint sometimes rejects a request's format without saying why (the
error reads "406: no error details returned"). Item queries therefore have
three forms, tried in order, and the working one is remembered:

1. `$filter=Id gt …` with `$orderby=Id` and a lean `Accept` header (fastest);
2. the same without `$orderby`;
3. the same with SharePoint's default `Accept` header.

Forms 2 and 3 rely on SharePoint returning items in ID order when no
`$orderby` is given, which it does. After each library the items read are
compared with the library's item count and any shortfall is reported. The
dashboard says how many libraries needed a fallback.

### Choosing what to scan (large sites)

On a site with many subsites and libraries, **Choose what to scan** shows a
map instead of starting one all-or-nothing scan.

1. **Map.** Every site, subsite and library is listed with SharePoint's own
   size (including version history) and file count, from one storage
   metrics request per library. No files are read.
2. **Select.** Tick libraries, or a whole site (it covers its subsites).
   Very large libraries (250,000+ files or 200 GB+) are flagged.
3. **Scan selected.** Only those libraries are read and merged into the
   saved result. Come back later for the rest; the header shows coverage.

Ages are always measured from the first scan's start, so libraries read on
different days add up consistently. Run **Run new scan** to start over from
scratch. The map is saved with the result, so a later visit shows what has
been read and what is left; **Refresh map** picks up new libraries.

Restricting who can scan: set *Who can run a scan* to *Only the people I
name* and list email addresses. Only they see the scan buttons. This limits
the page, not SharePoint permissions, and it cannot check tenant admin
roles (that would need a Microsoft Graph permission, which this web part
avoids), so list your admins by name.

### Very large libraries (a million files and more)

- **Paging by ID.** Each request asks for `Id gt <last ID read>`, ordered by
  ID, 5,000 at a time. If a batch still fails at 500 items (for example
  because of one damaged item), that batch is skipped and recorded, and
  the rest of the library is still read.
- **Parallel ID ranges.** A library with more than 20,000 items is split
  into 50,000-ID ranges (`$filter=Id gt … and Id le …`), and three readers
  work through them at once. The last range is open-ended, so files added
  during the scan are still read.
- **Patient with throttling.** See *Throttling: being a good neighbour*
  above. Network failures are retried separately, a few times.
- **Cancel is immediate**, even during a throttling wait.
- **Timeouts.** A page that times out is requested again with a smaller
  page size, down to 500.
- **Failures stay local.** If one range, library or subsite fails, it's
  marked and the rest of the scan carries on.
- **Nothing missed silently.**
  - After each library, the number of items read is compared with its
    `ItemCount`, and any shortfall is reported.

**How long it takes.** It's about one request per 5,000 files. A
million-file library is about 200 requests over three parallel readers.
Expect anywhere from several minutes to around half an hour, depending on
tenant load and throttling. The scan panel shows items read, data read,
speed and time left. Keep the page open until the scan finishes.

## What the numbers mean

- **Active or inactive is based on each file's *Modified* date,** counted
  in calendar months from the start of the scan. Editing a file's
  properties also updates *Modified*. SharePoint's REST API doesn't give
  ordinary users a per-file "last opened" date, so files that people only
  read count as inactive.
- **Sizes are the current version of each file.** Version history is
  reported separately.
- **Site storage used is larger than the file total.** It also includes
  version history, the recycle bins and list data.

## Get started

1. Download [`releases/storage-pulse.sppkg`](releases/storage-pulse.sppkg).
2. Upload it to the tenant App Catalog, or a site collection App Catalog,
   and click **Deploy**. The package has `skipFeatureDeployment: true`, so
   you can make it available to all sites or add it site by site.
3. Edit a page and add **Storage Pulse**, from the *Other* group.
4. Click **Start scan**.

For Microsoft Teams, select the app in the App Catalog and choose **Sync
to Teams**, then add it as a tab.

### Upgrading from Storage Activity Analyser (v1.x)

Storage Pulse is the same solution under a new name. The solution ID and
web part ID are unchanged, so pages that already host the web part keep
working, and their saved scans still load.

1. The package file name changed, and the App Catalog matches entries by
   file name. So **delete the old `storage-activity-analyser-webpart.sppkg`
   entry first**, then upload `storage-pulse.sppkg` and deploy it straight
   away.
2. New in v2: **only site owners can scan by default**. v1 let anyone scan
   for themselves. To keep that behaviour, set **Who can run a scan** to
   *Everyone*.

## Build from source

The SPFx build needs Node.js 18 (SPFx 1.18.2 supports 16.13–18.x).

```bash
npm install
npx gulp bundle --ship
npx gulp package-solution --ship    # -> sharepoint/solution/storage-pulse.sppkg
```

To try it on a real site, run `npx gulp serve` and open
`https://<tenant>.sharepoint.com/sites/<site>/_layouts/15/workbench.aspx`.

## Tests

The tests live in `tests/` and need Node.js 20+. They use a simulated
SharePoint REST API: webs, a 403 subsite, hidden and catalog libraries, a
large lazily generated library with deleted-ID gaps, item-level
permissions, missing next links, throttling and timeouts.

```bash
cd tests
npm ci
npm test          # unit + scan-engine tests (Node); run `gulp bundle` first, they test the compiled lib/
npm run e2e       # browser tests of the built web part; run `gulp bundle` first
```

- **Unit tests** cover:
  - age and threshold maths;
  - file-type splits matching the overall totals;
  - link and saved-file sanitising, including `javascript:` and
    other-domain URLs;
  - exact totals from a 130,000-item parallel scan;
  - hidden-item reporting, library exclusion and smaller-page retries;
  - Cancel during a 2-minute throttle wait.
- **Browser tests** cover:
  - the owner scan and save, and "saved by" taken from SharePoint;
  - a tampered results file (no script links rendered) and a corrupt one
    (a warning, no crash);
  - visitor permissions, the stale-results banner and Cancel;
  - dark mode, and phone width with no horizontal scrolling.

GitHub Actions (`.github/workflows/storage-pulse.yml`) builds, tests and
packages on every push, and publishes the `.sppkg` as an artifact.

## Project layout

```
src/webparts/storagePulse/
  StoragePulseWebPart.ts               Web part: settings, theme / section / Teams awareness
  StoragePulseWebPart.manifest.json
  assets/                              Toolbox icon
  components/
    StoragePulse.tsx                   Header, scan / cancel / export, load and save, banners
    ScanPanel.tsx                      Live scan view: radar, steps, speed, time left
    PulseMark.tsx                      The logo as animated inline SVG
    ErrorBoundary.tsx                  Keeps a rendering error from breaking the page
    text.ts                            Localized text helpers
    dashboard/                         Summary + gauge, age chart, file types, tables
  services/
    StorageScanService.ts              All SharePoint REST calls and the scan
    activity.ts                        Age buckets, thresholds, active / inactive maths
    safeData.ts                        Link safety and saved-file sanitising
    ResultsStore.ts                    Save / load the last scan in Site Assets
    ExportService.ts                   CSV exports
  models/IScanResult.ts
  loc/                                 All on-screen text (add a <locale>.js to translate)
sharepoint/assets/                     App Catalog icon and screenshot
assets/logo/                           Logo (SVG master, PNGs, lockup) and concepts
tests/                                 Simulated SharePoint, unit / scan / browser tests
```
