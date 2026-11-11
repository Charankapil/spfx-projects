# VANTAGE — SharePoint Online Work Management Foundation

![SPFx 1.24.1](https://img.shields.io/badge/SPFx-1.24.1-green.svg)
![Node.js 22 | 24 | 26](https://img.shields.io/badge/Node.js-22%20%7C%2024%20%7C%2026-green.svg)

**Version 0.2.0**: a reviewed and rebuilt development foundation, not a finished product.

VANTAGE is an SPFx web part that provisions *workspaces* in the current SharePoint site. Each workspace gets an active work-item list and an archive list. The web part then moves completed or cancelled task records to the archive list with a verified copy. It uses only SharePoint REST and the signed-in user's permissions. It needs no Entra ID app registration and no API permission requests.

> **Scope warning:** this release has no board or task UI, no attachment transfer, no restore, no unattended scheduling and no workspace-specific permissions. See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).

## Release package

The pre-built package is checked in at [`releases/vantage.sppkg`](releases/vantage.sppkg).

## Features

- One-time, idempotent site setup with a **Repair setup** option if a previous setup did not finish.
- Workspace creation, with duplicate-name protection and links to the provisioned lists.
- An archive policy editor with date, size-limit or combined methods, inline validation, a plain-language summary of what the policy does, and save-before-run.
- An archive run that copies each item, reads the copy back to check it, writes an audit entry, and only then deletes the source. Items with attachments are skipped and kept.
- Progress is checkpointed to the job record after every page. You can stop a run gracefully, and the browser warns you before you leave the page during a run. A run that stops reporting progress (for example, because the browser was closed) is released automatically after 30 minutes.
- Job history for each workspace, plus a per-item audit log in the `VNT_ArchiveLog` list.
- Theme-aware UI that follows the site theme, section background variants and dark themes. It is responsive (container queries) and keyboard and screen-reader friendly through Fluent UI.
- Permission-aware UI: setup and workspace creation need *Manage Lists*; archive runs need add, edit and delete item permissions.

## Build from source

Prerequisites: Node.js 22.14+ / 24 / 26 LTS and npm.

```bash
npm install
npm run build        # heft test --clean --production && heft package-solution --production
```

The build writes the package to **`sharepoint/solution/vantage.sppkg`**.

For local development, run `npm start` (`heft start`) and open the hosted workbench at `https://{tenant}.sharepoint.com/_layouts/workbench.aspx`.

## Deploy to a test site

1. Upload `vantage.sppkg` to the tenant or site-collection App Catalog. Choose **Enable this app and add it to all sites**. The package uses `skipFeatureDeployment: true`.
2. Add the **VANTAGE** web part (in the *Advanced* group) to a modern page. It supports full-width sections and single-part app pages.
3. As a site owner, select **Set up VANTAGE**, then create a test workspace.
4. Add synthetic task items to the workspace's `_Items` list. Set `Record Type = Task`, `Status = Completed` or `Cancelled`, and fill in `Completed Date`.
5. Configure, save and run an archive policy. Check **Job history** and `VNT_ArchiveLog`.
6. Don't enable archiving on business data until the checks in [docs/REVIEW.md](docs/REVIEW.md) pass in your tenant.

## Upgrading from the 0.1 / original 0.2 source

The original source could not create its own schema correctly (see [docs/REVIEW.md](docs/REVIEW.md), finding R1). Use a **clean test site**. If you already ran setup with the old package, delete the `VNT_*` lists it created first. The solution ID and web part ID are unchanged, so pages that already host the web part keep working after the upgrade.

## Source layout

| Path | Purpose |
|---|---|
| `src/webparts/vantage/VantageWebPart.ts` | SPFx entry point: React 18 root, theme and permission context |
| `src/webparts/vantage/components/VantageApp.tsx` | UI |
| `src/webparts/vantage/components/VantageApp.module.scss` | Scoped, theme-token styles |
| `src/webparts/vantage/services/SharePointService.ts` | REST wrapper (nometadata, throttling retry, `CreateFieldAsXml`) |
| `src/webparts/vantage/services/ProvisioningService.ts` | Declarative list and field schema, setup and workspace provisioning |
| `src/webparts/vantage/services/ArchiveService.ts` | Archive engine |
| `docs/ARCHITECTURE.md` | Architecture, schema and operational limits |
| `docs/REVIEW.md` | Architecture and UI review findings, and the fixes made |
