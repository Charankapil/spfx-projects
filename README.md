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
- [`ecm-foundation`](./ecm-foundation) — Module 1 of the ECM solution: the
  ECM tenant theme, the **ECM Hub Navigator** (a security-trimmed hub and
  department mega menu on every page, driven by one SharePoint list) and the
  **Metadata Explorer** web part (documents grouped by their metadata
  categories, with links to each file and a metadata coverage view). Includes
  an **ECM Setup** web part that provisions the theme and the navigation list
  from the browser, with no app registration or client id needed, and the
  **Governance Scorecard** (module 2), which measures metadata, retention
  labelling, stale content, storage and ownership per hub with saved trends. See its own [README](./ecm-foundation/README.md).
