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
- [`platform-lens`](./platform-lens) — A Power Platform admin
  center inside SharePoint: environments, apps, cloud flows, Copilot Studio agents,
  makers, connectors and DLP, licensing and Dataverse capacity, and a governance
  worklist, Power BI administration (workspaces, content, refresh health,
  activity and usage), plus Microsoft 365 dashboards (SharePoint sites, Teams and groups,
  licensing and adoption, identity). Every page exports to Excel, PDF and CSV.
  Reads live data with the admin's delegated Graph, Power Platform and Dataverse
  permissions (approved in SharePoint API access); also has a sample-data mode.
  See its [README](./platform-lens/README.md).
