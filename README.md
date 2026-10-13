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
- [`esignature`](./esignature) — Electronic signatures inside SharePoint:
  send PDF or Word documents to any number of signers in parallel,
  sequential or mixed order, drag-and-drop signature/initials/date/text/
  checkbox fields, draw/type/upload signatures, decline and delegate,
  reminders, expiry, templates, a certificate of completion with SHA-256
  tamper evidence, and a "Request signatures" button in every document
  library. No Azure app registration and no third-party signing service.
  See its own [README](./esignature/README.md).
