# File Type Analyser

An SPFx web part that scans a SharePoint site collection and renders a
tree-view overview of every document library and the file types stored in
it — with counts, a site-wide storage summary and CSV export. It is built
to remain usable on very large site collections (10+ TB, millions of
documents) and requires **no Azure AD app registration and no client
id/secret** — it only talks to SharePoint's own REST and Search APIs
using the current user's session.

## Why it scales

The obvious way to build this ("walk every folder, list every file") falls
over on a site collection with millions of documents — it means millions
of throttled REST calls. This solution avoids that entirely:

1. **Structure discovery is cheap.**
   `_api/web/getsubwebsfilteredforcurrentuser(...)` and `_api/web/lists`
   return metadata about subsites and document libraries only — never
   file listings. A site collection typically has dozens to low hundreds
   of libraries, not millions, so this step is fast regardless of how much
   content those libraries hold.
2. **File-type counts come from the search index, not file enumeration.**
   For each library, one call to `_api/search/query` with
   `querytext='IsDocument:1 Path:"<library URL>/*"'` and
   `refiners='FileType'` asks the already-built search index for an
   aggregated count per extension in that library. Whether the library
   holds 10 files or 10 million, this is a single request with a
   near-constant response size — the aggregation work happens server-side
   in the index, not in the browser.

   The trailing `/*` stops a library like `Documents` from also matching
   `Documents2`.

   **The search call must be made with OData 3.0.**
   `SPHttpClient.configurations.v1` sends an `OData-Version: 4.0` header,
   and the search REST endpoint only speaks OData 3.0 — it answers a v4
   request with HTTP 500 the moment it has real rows or refiners to
   serialize. An empty result set happens to serialize fine either way,
   which makes the failure look like a bad query (only libraries *with
   content* fail) when it is really a protocol-version mismatch. The
   service therefore calls search through
   `SPHttpClient.configurations.v1.overrideWith({ defaultODataVersion: ODataVersion.v3 })`;
   the `_api/web` and `_api/site` calls are fine on v4 and are left alone.
   Because OData 3 may wrap collections in a `{ results: [] }` envelope,
   the response parsing accepts both that and a bare array.

   Known limitation: `refiners='FileType'` without extra parameters caps
   the result at the 10 most common extensions per library. A library
   with more than 10 distinct file types will show its top 10 with the
   rest folded out of the per-library breakdown (the site-wide "Files
   indexed" total is unaffected — that comes from the query's total row
   count, not the refiner list).
3. **Storage numbers come from SharePoint's own tracked metrics.** The
   site-wide "storage used" figure comes from `_api/site?$select=Usage`,
   and each library's size from its root folder's storage metrics
   (`GetFolderByServerRelativeUrl('<library>')?$select=StorageMetrics&$expand=StorageMetrics`,
   `TotalSize` in bytes, versions included — the same number the Storage
   Metrics page shows). One request per library; no scanning of files.
   These exact figures appear in the site tree and the CSV.
4. **Storage per file type is estimated from search size bands.** Search
   cannot add up file sizes, but it can count files per size band. For each
   file type (up to 100, most common first) one site-wide query —
   `IsDocument:1 FileType:"<ext>" Path:"<site>/*"` with
   `refiners='Size(discretize=manual/10KB/50KB/…/10GB)'` — returns how many
   files fall in each of 19 roughly log-spaced bands, and the estimate is
   count × the band's geometric middle. Tested against synthetic lognormal
   file sizes the estimate lands within about 3% of the true total; real
   data can be further off if many files sit at one end of a band. This
   is one request per file type, independent of how many files exist.

   Limits, which the dashboard states: it covers **current file versions
   only** (version history is not in the search index), so the total is
   lower than site storage, which also counts versions, the recycle bin
   and metadata. If the tenant returns no size bands, the treemap falls
   back to file counts and says so.

The exact alternative — reading every file's size — would be a deep scan of
every item, which this tool is explicitly designed to avoid.

## No Azure, no client id

Every call goes through SPFx's built-in `context.spHttpClient`, which
reuses the current signed-in user's SharePoint session (the SPFx host page
already establishes this — form digest + cookie auth). There is:

- no Azure AD app registration,
- no client id / client secret / certificate,
- no Microsoft Graph, and
- no `AadHttpClient`.

The web part only ever calls `_api/*` endpoints on the SharePoint site it
is added to, using whatever permissions the current user already has
there.

## Features

- **Saved results** — the last scan is saved to the site and shown to
  everyone who opens the page, so nobody lands on an empty web part (see
  *Saved results* below).
- **Dashboard**
  - Headline figures — site storage, files, libraries, sites, file types —
    each with the change since the previous scan.
  - **Storage by file type** (WinDirStat's extension view) — one tile per
    extension across the whole site collection, sized by estimated storage
    and coloured by category (Word & text, PowerPoint, Excel & data, PDF,
    Images, Video & audio, Archives, Web & code, Other). Hover for the
    estimate, share, file count and average file size.
  - **File types list** — the top 12 extensions by number of files, with
    exact counts and shares; side by side with the treemap it shows which
    types are numerous versus which take the space.
- **Scan controls** — only site owners / site collection admins see
  **Start scan** / **Run new scan**; **Cancel** stops a scan and puts the
  previous results back.
- **Site tree** — site collection → subsites → libraries → file types, with
  per-library file counts and sizes, rendered live during a scan.
- **Export CSV** — one row per (web, library, file type) with counts and
  library size; works from saved results too.
- **Resilient** — an inaccessible subsite or failing library is marked in
  place and the scan continues; SharePoint `429`/`503` throttling is retried
  honouring `Retry-After`.

## Saved results

When a scan finishes, the web part writes it to
`SiteAssets/file-type-analyser-scan.json` in the **root site** of the site
collection (creating the Site Assets library if it does not exist), and
every visitor loads that file when the page opens. The file also keeps the
previous scan's headline figures, which is where the "since …" changes on
the dashboard come from.

- **Who can run a scan:** people with *Manage Web* permission on the site
  or site collection admins. Saving needs write access to Site Assets on
  the root site; if the save fails, the results are still shown to the
  person who ran the scan and a warning explains why they weren't saved.
- **Who can see the results:** anyone who can read the root site's Site
  Assets. The saved scan reflects what *the person who ran it* could see —
  run by an admin, it can list names and counts of libraries in subsites
  that other readers cannot open. It never contains file names or
  contents, only library names, counts and sizes.
- The saved file itself is counted as one `.json` file in Site Assets.

## Project layout

```
src/webparts/fileTypeAnalyser/
  FileTypeAnalyserWebPart.ts        Web part entry point, property pane
  FileTypeAnalyserWebPart.manifest.json
  components/
    FileTypeAnalyser.tsx            Main component: load saved scan, scan controls, save
    WebNodeTree.tsx                 Recursive tree renderer (webs, libraries, file types)
    IFileTypeAnalyserProps.ts
    FileTypeAnalyser.module.scss
    dashboard/
      Dashboard.tsx                 KPI row + treemap + file types panels
      KpiRow.tsx                    Headline figures with change since previous scan
      Treemap.tsx                   Storage-by-file-type treemap with hover tooltip
      squarify.ts                   Squarified treemap layout (no chart library)
      TopFileTypes.tsx              Labelled per-extension bar list
      fileTypeCategories.ts         Extension -> category and validated colour palette
      format.ts                     Number / date formatting
      Dashboard.module.scss
  services/
    SharePointService.ts            All SharePoint REST/Search calls + scan orchestration
    ResultsStore.ts                 Save / load the last scan in Site Assets
    ExportService.ts                CSV export
    httpErrors.ts                   Reads SharePoint's error message from a failed response
    sizeEstimate.ts                 Size-band parsing and per-type storage estimate
    formatBytes.ts                  Byte formatting helper
  models/                           Shared TypeScript interfaces
  loc/                              Localized strings
```

## Prerequisites

- Node.js `18.17.1` – `18.x` (SPFx 1.18 does not support Node 19+; use
  `nvm use 18` if your default Node is newer).
- SharePoint Online tenant (SPFx 1.18 targets SPO).

## Setup

```bash
npm install
```

## Run locally (workbench)

```bash
gulp serve
```

Open `https://<your-tenant>.sharepoint.com/_layouts/15/workbench.aspx` and
add the **File Type Analyser** web part to test against a real site
collection (the local workbench at `https://localhost:5432/workbench` has
no SharePoint context to scan against).

## Build and package for deployment

```bash
gulp bundle --ship
gulp package-solution --ship
```

This produces `sharepoint/solution/filetype-analyser.sppkg`.

## Deploy

1. Upload `filetype-analyser.sppkg` to your tenant or site collection
   **App Catalog**. When upgrading, upload a file with the **same file
   name** and choose **Replace** — the catalog replaces entries by file
   name, so a renamed file becomes a second entry for the same solution ID
   and can leave the web part undeployed and missing from the toolbox.
2. Deploy it when prompted ("Make this solution available to all sites in
   the organization" is optional — the solution has
   `skipFeatureDeployment: true`, so it can also be added site-by-site
   from the site contents / app store without a farm-wide feature
   activation).
3. Add the **File Type Analyser** web part to a page on the site
   collection you want to analyse.
4. Click **Start scan**.

No app registration, consent screen, or tenant admin step is required —
deployment is entirely through the App Catalog, and the web part runs
with the permissions of whichever user opens the page.

## Notes on scope

- The scan always starts from the root of the site collection the page
  belongs to and walks every subsite beneath it.
- Everything runs as the signed-in user, so results are security-trimmed:
  subsites the user cannot open are skipped (they are discovered with
  `getsubwebsfilteredforcurrentuser`, not `/webs`), and libraries or files
  they cannot see are not counted. Run it as a site collection
  administrator for the complete picture. If a subsite or library still
  fails (for example a 403 from unique permissions), the error is shown on
  that node and in the CSV, and the rest of the scan carries on.
- Document libraries only are scanned (SharePoint lists, e.g. calendars
  or custom lists, are intentionally excluded — this tool is about
  *files*).
- Results reflect what the search index currently has crawled. On a
  tenant with a large crawl backlog, very recently added files may take a
  short time to appear in the counts, exactly as they would in the normal
  SharePoint search experience.
