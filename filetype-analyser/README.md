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

1. **Structure discovery is cheap.** `_api/web/webs` and `_api/web/lists`
   return metadata about subsites and document libraries only — never
   file listings. A site collection typically has dozens to low hundreds
   of libraries, not millions, so this step is fast regardless of how much
   content those libraries hold.
2. **File-type counts come from the search index, not file enumeration.**
   For each library, one call to `_api/search/query` with
   `refiners='FileType'` asks the already-built search index for an
   aggregated count per extension under that library's path. Whether the
   library holds 10 files or 10 million, this is a single request with a
   near-constant response size — the aggregation work happens server-side
   in the index, not in the browser.
3. **Storage numbers come from SharePoint's own tracked usage.** The
   overall "storage used" figure shown in the summary comes from
   `_api/site?$select=Usage`, which SharePoint already maintains — no
   scanning required for it.

This is a deliberate trade-off: you get an accurate **file-count**
breakdown per type, per library, per site, essentially in real time, but
not an exact **byte size** per file type (the search index does not expose
a summable size property). If you need that, you would need a true deep
scan of every file's `Length` property, which is the expensive operation
this tool is explicitly designed to avoid. The site-wide storage figure in
the summary gives you the overall size context instead.

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

- **Start scan** button — kicks off an asynchronous scan; **Cancel**
  stops it early (already-collected results are kept).
- **Tree viewer** — site collection → subsites → document libraries →
  file types, expandable/collapsible, rendered incrementally as each
  library finishes so you don't stare at a blank screen on a big site
  collection.
- **Summary tiles** — site storage used, webs scanned, libraries scanned,
  total files indexed, distinct file types.
- **Site-wide file type chips** — aggregated totals across the whole
  site collection.
- **Export CSV** — one row per (web, library, file type, count),
  generated entirely client-side from the last completed scan.
- **Throttling-aware** — requests are gently paced and automatically
  retry with backoff on SharePoint `429`/`503` responses.

## Project layout

```
src/webparts/fileTypeAnalyser/
  FileTypeAnalyserWebPart.ts        Web part entry point, property pane
  FileTypeAnalyserWebPart.manifest.json
  components/
    FileTypeAnalyser.tsx            Main React component (scan controls, summary, tree)
    WebNodeTree.tsx                 Recursive tree renderer (webs, libraries, file types)
    IFileTypeAnalyserProps.ts
    FileTypeAnalyser.module.scss
  services/
    SharePointService.ts            All SharePoint REST/Search calls + scan orchestration
    ExportService.ts                CSV export
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
   **App Catalog**.
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

- The scan starts from the current site (the site the web part is added
  to) and walks its subsites. To get a true site-collection-wide view,
  add the web part to the root site of the site collection.
- Document libraries only are scanned (SharePoint lists, e.g. calendars
  or custom lists, are intentionally excluded — this tool is about
  *files*).
- Results reflect what the search index currently has crawled. On a
  tenant with a large crawl backlog, very recently added files may take a
  short time to appear in the counts, exactly as they would in the normal
  SharePoint search experience.
