# Vantage: architecture and UI reviews

This file has two sections: the review of the **1.0.0 source** (`VANTAGE-1.0.0-Source.zip`, the current release) and the earlier review of the **0.2.0 source**.

## Vantage 1.0.0 source review

The 1.0.0 upload added boards, tasks, collaboration, permissions, attachments and restore. However, it was built on the *original* 0.2 source, not on the corrected 0.2.0 rebuild, so most blocking defects from the 0.2 review came back. It also introduced new ones.

Vantage 1.0 keeps the uploaded feature set and design intent. It is a port onto the corrected foundation, with the fixes below.

### How it was verified

- `npm run build` (Heft: TypeScript 5.8, ESLint 9, Sass, webpack, package-solution) passes with **no errors and no warnings**.
- I unpacked the `.sppkg` and checked it: app name `Vantage`, version 1.0.0.0, the app icon (`AppIconPath`), two screenshots, the toolbox icon, unchanged solution and web part IDs, and no API permission requests.
- I ran the compiled UI in Chromium against a strict in-memory SharePoint REST fake. The fake rejects unknown fields, `__metadata` in nometadata payloads, bad headers, unique-constraint violations and role assignments made before inheritance is broken. The end-to-end scenario passed:
  1. Setup.
  2. Create a workspace, which is secured automatically: 3 groups and unique list permissions.
  3. Create a board and two tasks with the people picker and the date picker.
  4. Add a comment, a checklist item and a dependency, check the activity history, then change the status.
  5. Check the dashboard and the notifications.
  6. Seed 60 tasks plus 120 related records, including 4 tasks with attachments and one 40 MB attachment.
  7. Run a date-based archive. Result: **29 tasks / 87 records archived**, the oversized task **skipped with no partial copy**, both attachments copied and checksum-verified for every task that had them, no orphaned records left behind, and the lock released.
  8. **Restore** a task: it came back with its 2 related records and 2 attachments, and no residue was left in the archive.
  9. Add a Reader through the Access tab. Remount the page as that reader: they see only Overview, Boards & tasks and Workspaces, and "New task" is disabled.
  10. **Upgrade** from schema 2 to schema 3.
  11. Phone-width layout.
- **Not verified:** running against a live tenant. Use [TEST_PLAN.md](TEST_PLAN.md).

### Build and packaging

| # | Severity | Finding in the 1.0.0 source | Fix |
|---|---|---|---|
| V1 | **Blocker** (confirmed) | The build failed: the same `~@microsoft/sp-office-ui-fabric-core/.../SPFabricCore.scss` import error, and the same mix of gulp and Heft. `BUILD_STATUS.md` confirmed no `.sppkg` had ever been produced. | Heft 1.24.1 project (from the 0.2.0 rebuild). A real build, lint and package step replaces the grep-based `source-audit`, `validate-syntax` and `run-tests` scripts. |
| V2 | Medium | Version metadata was inconsistent (the manifest said `1.0.0`, with no `"*"`), the solution name was `vantage-enterprise-work-management-client-side-solution`, there was no app icon, and the web part was in the "Other" group with the Advanced group ID. | Solution name **Vantage**, version 1.0.0.0 everywhere, a new icon set (app catalog, toolbox, Teams, header), catalog screenshots and full metadata. |

### SharePoint runtime (each of these stops the app working)

| # | Severity | Finding | Fix |
|---|---|---|---|
| V3 | **Blocker** | Fields were still created with `SP.Field` + `StaticName`, so internal names became `Workspace_x0020_ID` and so on (0.2 finding R1). | `CreateFieldAsXml` with `AddFieldInternalNameHint`, from a declarative schema, including unique and indexed columns. |
| V4 | **Blocker** | `Content-Type: odata=nometadata`, but every create sent `__metadata` (list items, lists, groups, group users), which SharePoint rejects in that format. List and group creation and **every item write** would fail with a 400 error. | No `__metadata` anywhere. The fake rejects it, so this regression is now caught by the test. |
| V5 | **High** | Notification queries put the claims login (`i:0#.f\|membership\|…`) into the URL without encoding it. `#` starts a URL fragment, so the filter was cut off and notifications never loaded. The dashboard and board filters were also unencoded. | All filters go through `SharePointService.filter()`, which encodes them. |
| V6 | **High** | Concurrency checks read `@odata.etag`, which `nometadata` never returns, so they were silently skipped while appearing to work. | Removed the dead checks. Each source record's `Modified` value and child count are re-checked before deletion. |
| V7 | **High** (scale) | **Every** audit-log, job and lock row got unique permissions: break inheritance plus 2 role assignments, about 5 extra calls per row and 2–3 log rows per archived task. A few thousand archived tasks would pass SharePoint's recommended limit of 5,000 unique scopes per list and slow every query. | Control lists are secured at list level. Only the single policy row per workspace gets item-level permissions. Log and job rows hold IDs and outcomes only. Explained in SECURITY.md. |
| V8 | **High** | The archive lock check-then-create wasn't atomic. The unique field also didn't take effect because of V3. | A unique `WorkspaceId` on `VNT_ArchiveLocks`, created correctly, makes the insert atomic. Heartbeat per page; reclaimed after 30 minutes and the old job marked `Abandoned`. |
| V9 | High | An attachment over 25 MB was detected *after* the archive copy had been created, which left an orphan copy. Restoring that copy would then duplicate a still-active task. | Attachment sizes are checked before anything is written. Restore is refused while the original task still exists in the active list. |
| V10 | Medium | `isWorkspaceAdmin` loaded every user of the admin group for each workspace. Site owners who weren't in that group couldn't administer anything. | One `currentuser/groups` call resolves roles for all workspaces. Site owners (Manage Permissions) are always administrators. |
| V11 | Medium | Workspace-only users couldn't open the page at all: the site page and the workspace and board registries need site read access, which group membership doesn't give. | The Access tab can add people to the site Visitors group at the same time (on by default), and this is documented. |
| V12 | Medium | Deleting a task left orphaned dependency rows on other tasks that pointed at it. | `ParentTaskId eq X or RelatedTaskId eq X` rows are deleted with the task. |
| V13 | Medium | Field verification compared Booleans as strings, so `null` and `false` (default 0) never matched and archiving failed for older rows. | Comparison is type-aware for Boolean fields. |
| V14 | Low | `isInitialized` required `SchemaVersion === '1.0.0'`, and 0.2 sites had no upgrade path. The generic internal names `Status`, `Description`, `Type`, `Read` and `Message` risked clashing with built-in and site columns. | Numeric schema version with **Upgrade now** (idempotent; it also upgrades existing workspace lists). Prefixed names such as `WorkStatus` and `NotificationRead` are kept compatible with 0.2.0. |
| V15 | Low | Legacy `ReactDom.render`, `window.confirm`, `<input type="date">` and the 0.2 issues R3–R15 had all returned. | React 18 `createRoot`, Fluent dialogs and DatePicker, and all 0.2 fixes. |

### UI and UX

| # | Finding | Fix |
|---|---|---|
| X1 | One 70-line component with 40+ state hooks, written as minified one-liners. Hard to maintain and review. | Split into Shell, Overview, Boards, TaskPanel, Archive, Workspaces, Access and common components, with a shared app context. |
| X2 | Hard-coded navy and grey styles leaked globally (`.vnt-root .ms-Button{margin-top:8px}` changed every Fluent button). Themes, section backgrounds and dark mode were ignored. | Scoped CSS module driven by theme tokens. Panels and dialogs get the same variables. |
| X3 | The task editor sat inline above the list, and three buttons per row (Open / Edit / Delete) crowded the table. | Click a task to open a side **panel** with Details, Comments, Checklist, Dependencies and Activity tabs. Delete is in a danger zone at the bottom of the panel, with a confirmation dialog. |
| X4 | Assignee was a free-text "login/email" box, so typos silently broke notifications. | SharePoint **people picker**. The display name is stored for presentation. |
| X5 | No filtering or sorting; "latest 100" only. | Text, status and "assigned to me" filters. Sorted by due date then priority. Up to 500 tasks per board. |
| X6 | Escape, or closing the panel, silently discarded typed changes. Found by the test run. | Confirm-discard guard. Escape inside a picker no longer closes the panel. |
| X7 | No feedback for checklist ticks until the server round-trip finished. | Optimistic update. |
| X8 | KPI cards said "Tasks in current sample" with no explanation. The "SharePoint connected" chip and the footer were static. | Clear KPIs and a status breakdown. The sampling note appears only when the 2,000-task cap is reached. Static chrome removed. |
| X9 | Every module was visible to every user. Non-admins saw disabled archive and permission forms. | Tabs depend on role (Archive and Access for admins only). Readers get read-only panels. |
| X10 | The notifications tab and "Mark read" button cluttered the main navigation. | A header **bell** with an unread count opens a side panel with "Mark all as read". |
| X11 | Phone width: the table columns were truncated. | Below 600 px the table becomes a compact list of task cards. |
| X12 | Workspaces created without permissions were easy to miss. | A persistent warning and a **Secure workspace** action. Each workspace's access status and your role are shown in the list. |

---

## Vantage 0.2.0 source review

This review covers the uploaded `VANTAGE-0.2.0-Audited-Source.zip`. Each finding is marked as *confirmed* (reproduced by building the source or by running the code) or *by inspection*.

### How it was verified

- **Original source:** `npm install`, then `gulp bundle --ship`. **The build failed** (B1).
- **Rebuilt source:** `npm run build` (Heft: TypeScript 5.8, ESLint 9, Sass, webpack, package-solution) passes with no errors and no lint warnings. The output is `sharepoint/solution/vantage.sppkg`. The package was unzipped and its manifests checked: solution ID, web part ID, version 0.2.0.0, no API permission requests, and `skipFeatureDeployment`.
- **Behaviour:** the compiled UI and services were run in Chromium against a strict in-memory fake of the SharePoint REST API. The fake rejects unknown field names, unknown `$select` fields, missing headers and fields created without an internal-name hint. Scenario: setup, then create a workspace, try a duplicate name, seed 260 items (mixed statuses, attachments, one untitled item), then run a "both" policy (cutoff 2024-06-01, maximum 120 items, batch 50). Result: 140 archived, active list reduced to exactly 120, all open items and attachment items kept, the untitled item archived, the job record and audit log correct, and no REST contract errors.
- **Not verified:** this was not tested against a live tenant. Run the deployment checklist below before production use.

### Build and toolchain

| # | Severity | Finding | Fix |
|---|---|---|---|
| B1 | **Blocker** (confirmed) | The build failed. `vantage.scss` imported `~@microsoft/sp-office-ui-fabric-core/.../SPFabricCore.scss`, which no longer resolves with SPFx 1.24's Sass (`Can't find stylesheet to import`). No `.sppkg` could be produced. | Removed the import. Styles moved to a scoped CSS module. |
| B2 | High | The toolchain mixed gulp and Heft: gulp `sp-build-web` plus the Heft rig `spfx-web-build-rig`. It was missing `tslib`, `@types/webpack-env`, `sp-component-base`, ESLint config, `src/index.ts`, `config/sass.json` and `write-manifests.json`. | Rebuilt on the official SPFx 1.24.1 **Heft** scaffold, which is Microsoft's default toolchain since 1.22. The gulp toolchain is legacy. |
| B3 | Medium | `engines.node` was `>=24 <27`, which wrongly excluded Node 22 LTS, a supported version. | Uses the generator's range: 22.14+ / 24 / 26. |
| B4 | Medium | The README said 0.2.0, but `package.json` and `package-solution.json` said 0.1.0. The README's package path was also wrong for Heft. | Version 0.2.0 everywhere (manifest `"*"` takes it from `package.json`). Path documented as `sharepoint/solution/vantage.sppkg`. |
| B5 | Low | `package-solution.json` had no `developer`, `metadata` or `features` sections. `skipFeatureDeployment: false` blocked tenant-wide deployment. | Added full metadata and a feature. Set `skipFeatureDeployment: true`. |
| B6 | Low | The manifest group said "Other" but used the *Advanced* group ID. Strings were empty, with no `mystrings.d.ts`. | Fixed the group. Added typed localized strings. |
| B7 | Low | `scripts/source-audit.js` "verified" the code by grepping for string literals. It passed while the build was broken. | Removed. The real build, lint and package step is the gate now. |

### SharePoint architecture and runtime (data-safety critical)

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

### UI and UX

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

### Remaining limitations (as of 0.2.0; most are addressed in 1.0.0)

These are unchanged from the original scope, and are deliberately out of scope for this fix:

- Attachment transfer, related-record archival and restore.
- Unattended or scheduled runs: the page must stay open.
- An atomic distributed lock. The stale-job check is best effort; don't start runs for the same workspace from two browsers.
- Per-workspace permissions. Lists inherit the site's permissions.
- Board and task UI.

### Deployment checklist (0.2.0; superseded by TEST_PLAN.md)

1. On a clean test site, deploy `releases/vantage.sppkg`, add the web part and run setup. Check that the `VNT_*` lists have the exact internal names listed in ARCHITECTURE.md.
2. Create a workspace, add about 50 synthetic tasks (some with attachments, some open), and run each policy method.
3. Check the job record, `VNT_ArchiveLog`, the archive items' `OriginalItemId` and `ArchivedAt`, and that items with attachments and open items are still in the active list.
4. Close the browser during a run. Reopen after 30 minutes and confirm the old job is marked `Abandoned` and a new run starts.
5. Test as a Member and as a Visitor: the actions they can't perform should be disabled.
6. Repeat step 2 with more than 5,000 items to check list view threshold behaviour.
