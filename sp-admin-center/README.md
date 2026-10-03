# SharePoint Admin Center

**One dashboard for everyday SharePoint administration, running as you.**
No Azure app registration, no client ID or secret, no Graph permissions, no
tenant-admin consent. One `.sppkg`, and whatever permissions you already have.

## What it does

| Section | What an admin can do there |
| --- | --- |
| **Dashboard** | Storage vs quota, libraries / items / people / guests / subsites, health score, what needs attention, quick actions |
| **Sites** | Inventory of every site collection you can see (type, created, last activity, inactive 90/180/365+ days, M365-group or not), CSV export, **Manage** any site from the same screen |
| **People & permissions** | SharePoint groups and their members (add / remove), all users and guests, make or remove site collection admin, remove a user from the site, who has direct access, "what can this person do here?" checker, **bulk add** from pasted e-mail addresses |
| **Lists & libraries** | Item counts, view-threshold warnings, version history and search visibility toggles (single or bulk), unique-permission flags, CSV export |
| **Storage insights** | Files by type, files untouched for 1 / 2 / 3+ years, largest files (top 50 above 10 MB to 1 GB), all from the search index |
| **Recycle bin** | Newest 300 items of the site collection, search, bulk restore, bulk delete permanently |
| **Activity** | Permission and structure changes (members added or removed, role changes, lists / webs / groups created or deleted) for 24 h to 60 days |
| **Health check** | Rule-based findings: storage nearly full, single or too many admins, guests, organisation-wide groups, lists near the 5,000-item threshold, libraries without version history, hidden-from-search lists, inactive site |
| **Settings & log** | Edit site title and description, session audit log of every change made with the tool (CSV export) |

**Switch site** (or **Manage** in Sites) points every section at another site or
subsite on the same tenant, so one page administers many sites.

## Throttling: designed in, not bolted on

Everything goes through one client (`core/SPClient.ts`):

- At most **2 requests in flight**, request starts spaced **150 ms** apart.
- A `429` / `503` pauses the **whole queue**, honours `Retry-After` (with capped
  exponential back-off and jitter if absent), and retries at most 4 times.
  `RateLimit-*` headers, when SharePoint sends them, slow the queue down early.
  The header shows a live "SharePoint asked us to slow down" notice.
- **Nothing runs in the background.** No polling, no timers that call SharePoint;
  every request starts from a click or from opening a view.
- Counts and top-N lists (file types, stale files, largest files, site inventory)
  are **single search-index queries**. The tool never enumerates items or files.
- Paged reads have hard caps (5,000 users, 5,000 sites, 300 recycle-bin items,
  500 changes). Bulk actions are capped (200 people, 100 recycle-bin items) and
  run one item at a time.
- GET responses are cached for 5 minutes and identical in-flight requests are
  shared, so switching views never repeats a call. A write clears the cache.
- Requests to any origin other than the tenant the page is on are refused.

## How it runs without app registration

All calls use SPFx's `SPHttpClient` against `_api/*` of the SharePoint tenant
the page is on, with the signed-in user's own session. Search uses OData 3.0
(`SPHttpClient` defaults to 4.0, which the search endpoint rejects with HTTP 500
once it has rows). The package declares no `webApiPermissionRequests`.

Consequences, by design:

- You can do what your permissions allow. A refused change shows a clear
  "Access denied" message; nothing is bypassed.
- The Sites inventory and Storage insights are security-trimmed by search, so
  they cover what *you* can see. Being a SharePoint admin does not by itself
  make every site visible to search; pick a site by URL with **Switch site**.
- Tenant-level settings (sharing policy, storage quotas, site creation) live on
  the SharePoint admin center API and need a separate sign-in; they are out of
  scope on purpose.

## Deploy

1. Download [`releases/sharepoint-admin-center.sppkg`](releases/sharepoint-admin-center.sppkg).
2. Upload it to the tenant (or site collection) **App Catalog** and click **Deploy**.
   No API permission approval is needed.
3. Add **SharePoint Admin Center** to a page (a full-width / single-part app page
   works best) on the site you want as the home base.
4. Sign-in is automatic. Use **Switch site** to administer other sites.

To upgrade, upload the new file with the **same name** and choose Replace.

## Build

SPFx 1.18 needs Node 16.13 to 18.x.

```bash
npm install
npx gulp bundle --ship
npx gulp package-solution --ship   # sharepoint/solution/sharepoint-admin-center.sppkg
```

## Tests

```bash
npm run test:unit
```

Runs the real service code (throttling queue, retry / pause behaviour,
concurrency and spacing, caching, origin guard, search parsing for OData v3 and
v4, REST call shapes, health rules, CSV-injection escaping) under Node, once at
a modern target and once compiled to ES5 (the shipped target, where subclassing
`Error` behaves differently).

## Layout

```
src/webparts/adminCenter/
  core/SPClient.ts            Throttle-safe request queue, cache, retry, origin guard
  services/                   AdminApi (REST), SearchApi (search index), HealthEngine,
                              exportCsv, addresses, format
  components/AdminCenter.tsx  Shell: nav, site switcher, request meter, confirm dialog
  components/views/           One file per section
  components/shared/          Context, loader hook, charts (plain SVG) and cards
tests/                        Unit tests and a stand-in for @microsoft/sp-http
```
