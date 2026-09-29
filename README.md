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
- [`reinherit`](./reinherit) — **ReInherit**: finds subsites, lists,
  libraries, folders and files with unique permissions and restores
  permission inheritance in bulk, on the selected objects only or
  recursively down to a chosen folder depth or the last file. Built for
  libraries with millions of items (parallel ID-range reads, `$batch`
  resets), with a saved, exportable report of everything restored and
  who had access before. SharePoint REST only — no app registration. See
  its [README](./reinherit/README.md).
