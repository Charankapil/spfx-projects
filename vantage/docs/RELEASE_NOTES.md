# Vantage release notes

## 1.0.0

### New
- **Boards and tasks:** multiple boards per workspace; tasks with status, priority, assignee (people picker), due date and description; filtering, overdue highlighting, and a compact layout for mobile widths.
- **Task panel:** comments, a checklist with progress, dependencies and an activity history. Vantage asks before discarding unsaved changes.
- **Notifications:** a header bell lists assignments, comments and status changes, with "mark all as read".
- **Overview dashboard:** open, overdue and assigned-to-me counts, tasks by status, and workload by assignee.
- **Workspace security:** Administrators, Contributors and Readers groups, unique permissions on workspace lists, role-aware tabs and actions, and an Access tab for managing members.
- **Archive:** each task moves with its related records and attachments (up to 25 MB each, verified by checksum). An atomic per-workspace lock prevents concurrent runs. The run summary includes a remaining-over-limit count.
- **Restore** from the archive explorer, idempotent through `RestoredFromArchiveId`.
- **Upgrade path:** sites set up with 0.2.0 show **Upgrade now**, which keeps their data.
- **Branding:** the solution is named **Vantage** in the App Catalog, with a new icon used for the app, the web part toolbox, Teams and the header, plus catalog screenshots.

### Fixed (compared with the uploaded 1.0.0 source)
See [REVIEW.md](REVIEW.md#vantage-100-source-review) for the full list. In short: the source didn't build, its schema and REST payloads couldn't work against SharePoint, notification queries were truncated, and per-row permissions would not scale.

### Known limitations
- Archiving is browser-driven: the page must stay open while a run is in progress.
- A restored task gets a new ID; dependencies on other tasks that point at the old ID are not rewritten.
- No Kanban drag-and-drop, Gantt or calendar views, email or Teams notifications, or custom-field designer.

## 0.2.0
A reviewed rebuild of the original source: SPFx 1.24.1 Heft toolchain, correct field provisioning, a verified archive engine, and a theme-aware UI. See [REVIEW.md](REVIEW.md#vantage-020-source-review).
