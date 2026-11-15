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
- [`access-report`](./access-report) — Every site and library that chosen
  people and groups can reach, saved as a CSV in a SharePoint folder: on
  demand from the page, and on the first of every month from the included
  Power Automate flow. SharePoint REST only — no API permissions, no app
  registration, no premium connector. See its
  [README](./access-report/README.md).
