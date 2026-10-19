# ECM Foundation (module 1)

The first module of the custom SharePoint layer for the ECM strategy. It puts
the hub-and-spoke information architecture in front of users and makes
document metadata visible:

| Component | Type | What it does |
| --- | --- | --- |
| **ECM theme** | Tenant theme + CSS variables | One palette for SharePoint's own chrome (tenant theme) and for the ECM components (`--ecm-*` variables). |
| **ECM Hub Navigator** | Application customizer | Brand bar, search, hub tabs with a mega menu, breadcrumb and site classification on every page. Security-trimmed. Driven by one SharePoint list, so governance changes the menu without a redeploy. |
| **Metadata Explorer** | Web part | Every document in the site's libraries (or a hub / the tenant through search), grouped by any metadata category with a second-level "then by", facet filters, links to each file, a metadata coverage view and CSV export. |

Everything runs on the current user's SharePoint session through
SharePoint REST and Search: **no Azure AD app registration or API permission
is needed** for the package itself.

![ECM Hub Navigator with the Real Estate & Infrastructure hub open (sample data)](docs/images/hub-navigator.png)

![Metadata Explorer grouping a Leasing site's documents by Document Type, then Geography (sample data)](docs/images/metadata-explorer.png)

## Modules in the ECM solution

| # | Module | Status |
| --- | --- | --- |
| 1 | **ECM Foundation**: theme, Hub Navigator, Metadata Explorer | This project |
| 2 | Governance Scorecard: metadata/ownership coverage, retention labelling and storage per hub and department (pilot value targets) | Planned |
| 3 | Site Request and provisioning: standard department template, `ATC-SP-{Dept}-*` Entra groups, sensitive-site flow | Planned |
| 4 | Document Compliance Panel: label, owner, review date and lifecycle state on each document | Planned |
| 5 | Content Health: duplicate, stale and ownerless content before migration (extends [`filetype-analyser`](../filetype-analyser)) | Planned |

## Deploy

1. **Build** (Node 18, as SPFx 1.18.2 requires):

   ```bash
   npm install
   npm run build
   npm run package-solution   # -> sharepoint/solution/ecm-foundation.sppkg
   ```

   A pre-built package is in [`releases/ecm-foundation.sppkg`](releases/ecm-foundation.sppkg).

2. **Provision** with [PnP.PowerShell](https://pnp.github.io/powershell/) 2.x
   (needs SharePoint admin, and your own PnP app registration's client id):

   ```powershell
   ./provisioning/Deploy-EcmFoundation.ps1 `
     -TenantAdminUrl https://contoso-admin.sharepoint.com `
     -EcmHubUrl      https://contoso.sharepoint.com/sites/ecm-hub `
     -ClientId       <your PnP app client id> `
     -ApplyThemeTo   https://contoso.sharepoint.com/sites/real-estate, https://contoso.sharepoint.com/sites/finance-hub `
     -PackagePath    ./releases/ecm-foundation.sppkg `
     -GrantEveryoneRead
   ```

   The script is safe to re-run. It:
   - registers **ECM Theme** ([`provisioning/theme/ecm-theme.json`](provisioning/theme/ecm-theme.json)) as a tenant theme and applies it to the sites you list;
   - creates the **ECM Navigation** list on the Corporate ECM Hub and seeds it from
     [`provisioning/navigation-seed.json`](provisioning/navigation-seed.json):
     the 6 functional hubs, 21 department sites, hub resources and 6 shared enterprise
     services from the strategy. The seed URLs are placeholders (`/sites/leasing`, …):
     **edit them to your real site URLs** before running, or in the list afterwards;
   - uploads and deploys the package tenant-wide, and points the navigator's
     **Tenant Wide Extensions** entry at your hub.

3. **Everyone must be able to read the navigation list.** If they can't, the navigator shows
   "Navigation could not be loaded". Either use `-GrantEveryoneRead` or make sure the
   hub's Visitors group covers all employees.

To limit the navigator to some sites, use the Tenant Wide Extensions list's
template filter. You can also delete that entry and add the customizer to specific sites
with `Add-PnPApplicationCustomizer -ClientSideComponentId 293dd12c-1019-42fd-b262-79b6f949e9b7`.

## ECM Hub Navigator

### The navigation list

| Column | Internal name | Use |
| --- | --- | --- |
| Title | `Title` | Text shown in the menu |
| Type | `EcmNavType` | `Hub`, `Department site`, `Shared service`, `Hub resource` or `Header link` |
| Parent hub | `EcmParentHub` | Hub title, for department sites and hub resources |
| URL | `EcmUrl` | Where the entry links to |
| Description | `EcmDescription` | Line under a site name; on a hub, the governance note in the menu |
| Order | `EcmOrder` | Sort order within its group |
| Restricted | `EcmRestricted` | Show only to people who can open the site |
| Active | `EcmActive` | Untick to hide without deleting |
| Hub colour | `EcmColor` | Optional hex accent for a hub (defaults follow the strategy deck) |

The first **Header link** is shown as the primary button (the seed makes this
"Request a site").

### How it behaves

- **Fast.** The menu is cached in `localStorage` for `cacheMinutes` (30 by default). Pages paint from
  the cache and refresh in the background. The bundle is about 44 KB because it uses no Fluent UI.
- **Security-trimmed.** Rows marked *Restricted* are kept only if a search query (which SharePoint
  security-trims) returns that site for the current user. One small query covers every restricted
  site. If it fails, restricted rows stay hidden (fail closed). The result is cached per session.
  Unrestricted rows show to everyone, so leave a site unrestricted only if everyone can open it.
- **Knows where you are.** It matches the current site against the list. A site that isn't listed
  but is associated with a hub still gets that hub's breadcrumb, through SharePoint's hub association.
- **Works on phones.** Below 768 px the tabs become a drawer with an accordion per hub.
  Escape and clicking outside close the menu.

### Properties

Set these in the Tenant Wide Extensions entry, or with `Add-PnPApplicationCustomizer -ClientSideComponentProperties`:

```json
{
  "configSiteUrl": "/sites/ecm-hub",
  "listTitle": "ECM Navigation",
  "cacheMinutes": 30,
  "homeTitle": "Corporate ECM Hub",
  "brandBadge": "AT",
  "searchUrl": "https://contoso.sharepoint.com/_layouts/15/search.aspx/?q={q}"
}
```

Only `configSiteUrl` is required.

## Metadata Explorer

Add **Metadata Explorer** (group "ECM") to any page. It has two sources:

### Document libraries in this site (default, no setup)

The web part reads every document library on the site. You can limit it to the libraries
you list in the settings. Every custom column becomes a category:

- choice, managed metadata, person, lookup, text, yes/no, number, date (grouped by year);
- plus **Content type**, **Retention label**, **Library**, **File type** and **Modified (year)**.

Values are the labels people see, so taxonomy terms appear as their names.
Libraries over the 5,000-item view threshold are paged safely, and folders are skipped.

### Search index (site, hub or whole organisation)

Use this for a hub-wide or tenant-wide view. Each category is a managed property,
listed one per line as `Title|ManagedProperty`. Map the ECM core columns in the search
schema first (SharePoint admin center → Search → Manage Search Schema), for example:

| Category | Crawled property | Managed property |
| --- | --- | --- |
| Document Type | `ows_taxId_DocumentType` (or `ows_DocumentType`) | `RefinableString00` |
| Department | `ows_taxId_Department` | `RefinableString01` |
| Record Category | `ows_RecordCategory` | `RefinableString02` |
| Confidentiality | `ows_Confidentiality` | `RefinableString03` |
| Status | `ows_Status` | `RefinableString04` |

Both plain and taxonomy (`L0|#0guid|Label`) values are read, and multi-value properties are
split on `;`. Hub scope uses the page's hub (`DepartmentId`). On a site that has no hub, it
falls back to the current site.

### Using it

- **Group by / Then by**: any category, two levels deep. A document with several values
  (multi-value columns) appears under each of them. "(No value)" groups show the metadata
  gaps that need fixing.
- **Facets**: tick values to filter. Filters combine with AND across categories and OR within one.
  Each facet's counts reflect the other filters.
- **Metadata coverage**: the share of documents with each category filled in. This is the
  pilot's "metadata and ownership coverage" measure, and it follows the current filters.
- **Export CSV**: the filtered documents with every category, for offline review.
- File names open in the browser for Office files and directly for other types.

**Maximum documents** (5,000 by default) caps how much the web part loads in one go.
A warning appears when there are more.

## Project layout

```
src/
  common/                 shared by every ECM component
    theme/ecmTheme.ts     palette, hub colours, --ecm-* CSS variables
    services/SpRestClient.ts  REST + search with throttling retry
  extensions/ecmHubNavigator/
    EcmHubNavigatorApplicationCustomizer.ts
    services/NavigationService.ts   list load, cache, security trimming, location
    components/HubNavigator.tsx     bar, mega menu, drawer, breadcrumb
  webparts/metadataExplorer/
    services/LibrariesSource.ts     RenderListDataAsStream per library
    services/SearchSource.ts        search query, managed property mapping
    services/grouping.ts            filters, facets, groups, coverage, CSV
    components/                     explorer, facets, groups, tables
provisioning/             theme, navigation seed, Deploy-EcmFoundation.ps1
sharepoint/assets/        ClientSideInstance.xml (tenant-wide registration)
```
