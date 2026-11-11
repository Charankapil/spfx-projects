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
- [`site-toolkit`](./site-toolkit) — Fourteen everyday web parts and a
  site-wide alert banner: storage, stale content, duplicates, recycle bin,
  bulk metadata editing, policy acknowledgements, polls, templates, org
  chart, FAQ, access review and broken links. No API permissions, and every
  request goes through one throttle-safe queue. See its own
  [README](./site-toolkit/README.md).
