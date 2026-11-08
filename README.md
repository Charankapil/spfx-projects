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
- [`experience-suite`](./experience-suite) — SharePoint Experience Suite: one
  employee web part with seven modules (MySpace, RequestHub, DocuSmart,
  SiteNavigator, ApprovalFlow, TeamPulse, PeopleConnect) backed by
  SharePoint lists and Search. Uses the signed-in user's SharePoint session —
  no app registration for the web part. See its [README](./experience-suite/README.md)
  and the [code review notes](./experience-suite/REVIEW.md).
