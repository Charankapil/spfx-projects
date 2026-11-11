# VANTAGE 0.2.0: architecture and UI review

This review covers the uploaded `VANTAGE-0.2.0-Audited-Source.zip`. Each finding is marked as *confirmed* (reproduced by building the source or by running the code) or *by inspection*.

## How it was verified

- **Original source:** `npm install`, then `gulp bundle --ship`. **The build failed** (B1).
- **Rebuilt source:** `npm run build` (Heft: TypeScript 5.8, ESLint 9, Sass, webpack, package-solution) passes with no errors and no lint warnings. The output is `sharepoint/solution/vantage.sppkg`. The package was unzipped and its manifests checked: solution ID, web part ID, version 0.2.0.0, no API permission requests, and `skipFeatureDeployment`.
- **Behaviour:** the compiled UI and services were run in Chromium against a strict in-memory fake of the SharePoint REST API. The fake rejects unknown field names, unknown `$select` fields, missing headers and fields created without an internal-name hint. Scenario: setup, then create a workspace, try a duplicate name, seed 260 items (mixed statuses, attachments, one untitled item), then run a "both" policy (cutoff 2024-06-01, maximum 120 items, batch 50). Result: 140 archived, active list reduced to exactly 120, all open items and attachment items kept, the untitled item archived, the job record and audit log correct, and no REST contract errors.
- **Not verified:** this was not tested against a live tenant. Run the deployment checklist below before production use.

## Build and toolchain

| # | Severity | Finding | Fix |
|---|---|---|---|
| B1 | **Blocker** (confirmed) | The build failed. `vantage.scss` imported `~@microsoft/sp-office-ui-fabric-core/.../SPFabricCore.scss`, which no longer resolves with SPFx 1.24's Sass (`Can't find stylesheet to import`). No `.sppkg` could be produced. | Removed the import. Styles moved to a scoped CSS module. |
| B2 | High | The toolchain mixed gulp and Heft: gulp `sp-build-web` plus the Heft rig `spfx-web-build-rig`. It was missing `tslib`, `@types/webpack-env`, `sp-component-base`, ESLint config, `src/index.ts`, `config/sass.json` and `write-manifests.json`. | Rebuilt on the official SPFx 1.24.1 **Heft** scaffold, which is Microsoft's default toolchain since 1.22. The gulp toolchain is legacy. |
| B3 | Medium | `engines.node` was `>=24 <27`, which wrongly excluded Node 22 LTS, a supported version. | Uses the generator's range: 22.14+ / 24 / 26. |
| B4 | Medium | The README said 0.2.0, but `package.json` and `package-solution.json` said 0.1.0. The README's package path was also wrong for Heft. | Version 0.2.0 everywhere (manifest `"*"` takes it from `package.json`). Path documented as `sharepoint/solution/vantage.sppkg`. |
| B5 | Low | `package-solution.json` had no `developer`, `metadata` or `features` sections. `skipFeatureDeployment: false` blocked tenant-wide deployment. | Added full metadata and a feature. Set `skipFeatureDeployment: true`. |
| B6 | Low | The manifest group said "Other" but used the *Advanced* group ID. Strings were empty, with no `mystrings.d.ts`. | Fixed the group. Added typed localized strings. |
| B7 | Low | `scripts/source-audit.js` "verified" the code by grepping for string literals. It passed while the build was broken. | Removed. The real build, lint and package step is the gate now. |

## SharePoint architecture and runtime (data-safety critical)

| # | Severity | Finding | Fix |
|---|---|---|---|
| R1 | **Blocker** (by inspection; documented SharePoint behaviour) | Fields were created by POSTing `SP.Field` with `Title = displayName` and `StaticName = internalName`. SharePoint derives the **internal name from the title**, so you get `Workspace_x0020_ID`, not `WorkspaceId`. Every later item write (`WorkspaceId: …`) fails with "property does not exist". Re-running setup also adds duplicate columns, because the existence check looks up `WorkspaceId`. Setup could never produce a working schema. | Fields are created with `fields/createfieldasxml` and `AddFieldInternalNameHint`, from a declarative schema. One field read per list instead of one per field. |
| R2 | **Blocker** (by inspection) | Requests sent `Accept: application/json;odata=verbose` through `SPHttpClient.configurations.v1`, which adds `OData-Version: 4.0`. SharePoint commonly rejects that combination ("The HTTP header ACCEPT is missing or its value is invalid") unless `odata-version` is cleared. | Moved to OData v4 `nometadata`, with `odata-version: ''`. No `__metadata` is needed, which removes an extra entity-type GET on **every** item create (it previously doubled the request count during archiving). |
| R3 | High | When an archive run failed, `LastError` (a 255-character text field) received the full SharePoint error body. That MERGE failed, so the job stayed `Running` forever. | `LastError` is now a note field, and values are truncated. |
| R4 | High | Combined with R3, a closed browser left a `Running` job that blocked the workspace **permanently**, with no recovery in the UI. | Progress is checkpointed after every page. A `Running` job that hasn't checkpointed for 30 minutes is marked `Abandoned` automatically. Graceful **Stop** and a `beforeunload` warning were added. |
| R5 | High | Verification always failed for items without a title: the copy was written as `(untitled)`, but the check compared it to an empty string. Those items could never be archived. | The comparison now normalises the title. Confirmed in the test run. |
| R6 | Medium | `isInitialized` checked only 2 of the 5 lists, so a partial setup looked complete. The Initialize button was then disabled with no way to repair. | All 5 lists are checked. Missing lists are reported, and **Repair setup** re-runs the idempotent setup. |
| R7 | Medium | Generic internal names (`Status`, `Description`, `Mode`) risk clashing with built-in and site columns. | Renamed to `WorkStatus`, `WorkDescription`, `ArchiveMode`, `JobStatus` and `WorkspaceStatus`. Display names are unchanged. |
| R8 | Medium | The workspace registry row was written *after* the lists, but there was no duplicate-name check, so users could create the same workspace twice. | Case-insensitive duplicate check. The registry row is still written last. |
| R9 | Medium | `class extends Error` with an ES5 target breaks `instanceof`, so 404 detection silently failed. This was caught during the rebuild. | `Object.setPrototypeOf` in the `SharePointError` constructor. |
| R10 | Medium | An item skipped in the date phase was counted and logged again in the count phase. | Each item is counted and logged once per run. |
| R11 | Low | A retried DELETE that had actually succeeded returned 404 and was reported as a failure. | A 404 after a retry counts as success. |
| R12 | Low | The `savePolicy` OData filter didn't escape the workspace ID, unlike the other queries. | All filters escape values through `SharePointService.odataString`. |
| R13 | Low | Dates entered as `YYYY-MM-DD` relied on browser parsing. Future cutoffs were accepted. | A Fluent DatePicker stores UTC midnight of the chosen calendar date. Future dates are rejected. |
| R14 | Low | `nextLink` continuation URLs were followed without checking the host. | Continuation links must stay on the same host. |
| R15 | Info | React 18 was used with the legacy `ReactDom.render`, which runs in React 17 mode with a deprecation warning. | `createRoot`, with `unmount` on dispose. |

## UI and UX

| # | Finding | Fix |
|---|---|---|
| U1 | The sidebar showed *Workspaces, Boards, My Tasks, Administration* as plain `div`s. They did nothing and couldn't be focused: false affordances that aren't accessible. | Removed. The real sections are Fluent `Pivot` tabs (Overview, Workspaces, Archive & retention, Job history), with keyboard support and an overflow menu on narrow widths. |
| U2 | Colours were hard-coded (a navy header bar and greys) despite `supportsThemeVariants: true`. The web part ignored site themes, section backgrounds and dark mode, and the page-level `.vnt-*` styles leaked. | Scoped CSS module. All colours come from theme semantic colours passed as CSS variables, and the component is wrapped in `ThemeProvider`. |
| U3 | The "Connected to SharePoint" chip and the "Good day" greeting were static, even after errors. | Removed. Status is shown only where it is real (KPI tiles, badges). |
| U4 | A fixed 680px minimum height and a 205px sidebar didn't fit one-third or one-half columns. | Fluid layout with container queries, so it responds to the web part's own width rather than the viewport. Supports full-bleed sections. |
| U5 | A browser `window.confirm` guarded a destructive action. | A Fluent blocking dialog repeats in plain language what the policy will do, and names the source and target lists. |
| U6 | Run used unsaved form values, so a run could differ from the stored policy. | Save-before-run (the button is disabled while there are unsaved changes, with a tooltip), plus **Discard changes**. |
| U7 | No validation feedback: invalid numbers were sent to SharePoint. | Inline validation message. Save and Run are disabled while the policy is invalid. |
| U8 | No permission awareness: readers could press Initialize and Create and got raw 403 errors. | `SPPermission` checks drive the disabled states, with guidance ("ask a site owner"). |
| U9 | No job history or archive visibility. Progress was only a text line. | Job history table with status badges, KPI tiles (workspaces, active and archived counts, last run), progress counters and links to the lists and the audit log. |
| U10 | Workspace selection was buried in the "Create workspace" panel. | Workspace switcher in the header, plus a workspace table with Select actions. |

## Remaining limitations before production

These are unchanged from the original scope, and are deliberately out of scope for this fix:

- Attachment transfer, related-record archival and restore.
- Unattended or scheduled runs: the page must stay open.
- An atomic distributed lock. The stale-job check is best effort; don't start runs for the same workspace from two browsers.
- Per-workspace permissions. Lists inherit the site's permissions.
- Board and task UI.

## Deployment checklist (tenant)

1. On a clean test site, deploy `releases/vantage.sppkg`, add the web part and run setup. Check that the `VNT_*` lists have the exact internal names listed in ARCHITECTURE.md.
2. Create a workspace, add about 50 synthetic tasks (some with attachments, some open), and run each policy method.
3. Check the job record, `VNT_ArchiveLog`, the archive items' `OriginalItemId` and `ArchivedAt`, and that items with attachments and open items are still in the active list.
4. Close the browser during a run. Reopen after 30 minutes and confirm the old job is marked `Abandoned` and a new run starts.
5. Test as a Member and as a Visitor: the actions they can't perform should be disabled.
6. Repeat step 2 with more than 5,000 items to check list view threshold behaviour.
