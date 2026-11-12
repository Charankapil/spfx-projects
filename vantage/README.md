<img src="assets/vantage-icon-300.png" alt="Vantage icon" width="96" align="right" />

# Vantage: work management for SharePoint Online

![SPFx 1.24.1](https://img.shields.io/badge/SPFx-1.24.1-green.svg)
![Node.js 22 | 24 | 26](https://img.shields.io/badge/Node.js-22%20%7C%2024%20%7C%2026-green.svg)
![Version 1.0.0](https://img.shields.io/badge/version-1.0.0-blue.svg)

Vantage is a single SPFx web part. It brings boards, tasks, comments, checklists, dependencies and notifications to SharePoint Online, grouped into secure workspaces. When work is finished, Vantage moves it to an archive list and can restore it later.

Everything is stored in SharePoint lists in your site. Vantage calls only the SharePoint REST API, using the signed-in user's permissions. It needs no Entra ID app registration, no API permission requests and no backend.

## Release package

The pre-built package is **[`releases/vantage.sppkg`](releases/vantage.sppkg)**. It appears in the App Catalog as **Vantage**, with its own icon and screenshots.

## Features

| Area | What you get |
|---|---|
| Overview | Open, overdue, assigned-to-me and completed counts, tasks by status, and open work by assignee. |
| Boards & tasks | Multiple boards per workspace. Filter tasks by text, status or "assigned to me". Overdue dates are highlighted. On narrow screens the table becomes a compact card list. |
| Task panel | Edit tasks with a people picker and a date picker. Includes comments, a checklist (with progress), dependencies and an activity history. If you try to close with unsaved changes, Vantage asks you to confirm. |
| Notifications | A header bell shows assignments, comments and status changes for the current workspace. |
| Archive & retention | Archive by date, by list size, or both. The page shows a plain-language summary of the policy and requires you to save before running. Progress, stop, job history, audit log and an archive explorer with **Restore** are included. |
| Workspaces | Create workspaces and see your role and the access status of each one. |
| Access | Add people as Administrators, Contributors or Readers with a people picker, and remove them. One click secures workspaces created before permissions were set up. |
| Setup & upgrade | Setup runs once and is safe to repeat. It also repairs a partial setup. Sites set up with 0.2.0 get an **Upgrade now** button that keeps their data. |
| Design | Follows SharePoint themes, section background colours and dark mode. Responsive using container queries. Keyboard and screen-reader friendly (Fluent UI 8). |

## Build from source

You need Node.js 22.14+, 24 or 26 LTS, and npm.

```bash
npm install
npm run build        # heft test --clean --production && heft package-solution --production
```

The build writes the package to **`sharepoint/solution/vantage.sppkg`**. For local development, run `npm start` and open `https://{tenant}.sharepoint.com/_layouts/workbench.aspx`.

## Deploy

1. Upload `vantage.sppkg` to the tenant or site-collection App Catalog and choose **Enable this app and add it to all sites**.
2. On a modern page, add the **Vantage** web part. It is in the *Advanced* group and works in full-width sections and on single-part app pages.
3. As a **site owner**, select **Set up Vantage**, then create a workspace. Vantage secures the new workspace automatically.
4. On the **Access** tab, add people. Leave "Also give them read access to this site" ticked for anyone who isn't already a site member or visitor.
5. Follow [docs/TEST_PLAN.md](docs/TEST_PLAN.md) on a test site before you use Vantage for business data.

**Upgrading from 0.2.0:** deploy the new package over the old one; the solution ID and web part ID are unchanged. A site owner then opens the page and selects **Upgrade now**.

## Documentation

- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md): data model, archive and restore protocol, scale limits
- [docs/SECURITY.md](docs/SECURITY.md): the permission model and what it does and doesn't protect
- [docs/TEST_PLAN.md](docs/TEST_PLAN.md): tenant acceptance tests
- [docs/REVIEW.md](docs/REVIEW.md): review findings for the uploaded sources and the fixes made
- [docs/RELEASE_NOTES.md](docs/RELEASE_NOTES.md)

## Source layout

| Path | Purpose |
|---|---|
| `src/webparts/vantage/VantageWebPart.ts` | Entry point: React 18 root, theme, user and permission context |
| `components/VantageApp.tsx` | Shell: header, workspace switcher, notifications, setup and upgrade, role-based tabs |
| `components/OverviewTab.tsx`, `BoardsTab.tsx`, `TaskPanel.tsx`, `ArchiveTab.tsx`, `WorkspacesTab.tsx`, `AccessTab.tsx` | Feature screens |
| `components/common.tsx` | App context, people picker, confirm dialog, badges, formatting |
| `services/SharePointService.ts` | REST wrapper: nometadata format, throttling retry, `CreateFieldAsXml`, attachments, groups and permissions |
| `services/ProvisioningService.ts` | Declarative schema, setup and upgrade, workspaces, boards |
| `services/WorkService.ts` | Tasks, related records, notifications, dashboard |
| `services/ArchiveService.ts` | Archive, restore, lock, audit |
| `services/PermissionService.ts` | Workspace groups, roles, unique permissions |
| `assets/` | Icon sources (SVG) |
