# SPFx Projects

A collection of SharePoint Framework (SPFx) web part solutions.

## Projects

- [`filetype-analyser`](./filetype-analyser) — Scans a SharePoint site
  collection and shows what it is made of by file type: a dashboard,
  treemap, site tree and CSV export, fast even on multi-terabyte sites
  with millions of files. Uses only
  SharePoint REST/Search APIs — no Azure AD app registration or client id
  required. See its own [README](./filetype-analyser/README.md) for setup,
  architecture and deployment details.
- [`storage-pulse`](./storage-pulse) — **Storage Pulse**: shows how much
  of a site or site collection's file storage is in files people still use,
  and how much is in files nobody has modified for a chosen period (3
  months to 5 years). The breakdown is by age, file type, library and site,
  with the largest inactive files and CSV exports. It runs with the
  current user's permissions only, with no app registration and no API
  permissions. See its own [README](./storage-pulse/README.md).
