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
- [`sp-admin-center`](./sp-admin-center) — One dashboard for everyday
  SharePoint administration: site inventory, people and permissions (with bulk
  add), lists and libraries, storage insights, recycle bin, activity feed and a
  health check. Runs as the signed-in user with no app registration, and is
  built to stay clear of throttling. Prebuilt package in
  [`releases/`](./sp-admin-center/releases). See its own
  [README](./sp-admin-center/README.md).
