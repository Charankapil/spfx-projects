# VANTAGE 0.2.0 architecture

## Application boundary

VANTAGE is a single SPFx 1.24.1 web part (Heft toolchain, React 18, Fluent UI 8). It is hosted on a modern SharePoint page or in a Teams tab. All calls go to the *current web's* REST API through `SPHttpClient`, using the signed-in user's session and permissions. There is no backend, no app registration and no `webApiPermissionRequests`.

## Provisioned objects

Setup creates these lists (generic list template 100). It is idempotent, so you can re-run it to repair a partial setup. Every field is created with `CreateFieldAsXml` and `AddFieldInternalNameHint`, so the internal names below are exact.

| List | Fields (internal name) |
|---|---|
| `VNT_Settings` | `Title` (key), `SettingValue`. Holds `SchemaVersion = 2`. |
| `VNT_Workspaces` | `WorkspaceKey`\*, `WorkspaceStatus`\*, `WorkItemsList`, `ArchiveList` |
| `VNT_ArchivePolicies` | `WorkspaceId`\*, `Enabled`, `ArchiveMode`, `MaxActiveItems`, `ArchiveBefore`, `EligibleStatuses`, `BatchSize` |
| `VNT_ArchiveJobs` | `WorkspaceId`\*, `JobStatus`\*, `ArchiveMode`, `Started`, `Finished`, `Processed`, `Archived`, `Skipped`, `FailedCount`, `LastError` (note) |
| `VNT_ArchiveLog` | `WorkspaceId`\*, `OriginalItemId`\*, `ArchiveItemId`, `ArchiveJobId`\*, `Outcome`, `Detail` (note) |

Each workspace also gets two lists:

- `VNT_<key>_Items` (active): `WorkspaceId`\*, `BoardId`\*, `RecordType`\*, `WorkStatus`\*, `Priority`, `AssignedToKey`\*, `DueDate`\*, `CompletedDate`\*, `Progress`, `WorkDescription` (plain-text note), plus custom slots `CustomText1-8`, `CustomNumber1-4`, `CustomDate1-4` and `CustomBoolean1-2`.
- `VNT_<key>_Archive`: the same fields, plus `OriginalItemId`\*, `ArchivedAt`\*, `ArchiveJobId`, `OriginalCreated` and `OriginalModified`.

\* = indexed. The active list has 7 indexes, well inside the limit of 20.

Fields use `WorkStatus`, `WorkDescription`, `WorkspaceStatus`, `JobStatus` and `ArchiveMode` instead of generic names like `Status`, `Description` or `Mode`. This avoids clashes with built-in and site columns. Display names are still "Status", "Description" and so on.

Each workspace adds two lists to the site. Monitor the site's list count; this version doesn't create extra site collections.

## Archive policy semantics

- **Date:** archives terminal (Completed or Cancelled) `Task` records whose `CompletedDate` is before the cutoff. The cutoff is stored as UTC midnight of the chosen calendar date and can't be in the future.
- **Count:** re-reads the active list's `ItemCount`, then archives the oldest eligible tasks until the list is at or below the maximum. `ItemCount` includes non-task records, so the target may not be reachable when the list holds many of them.
- **Both:** runs the date phase first, then the count phase on what is left.

## Archive sequence

1. Validate the policy. The UI also requires the policy to be saved, so a run always matches the stored policy.
2. Check for a `Running` job in this workspace. If that job checkpointed less than 30 minutes ago, refuse to start. Otherwise mark it `Abandoned`.
3. Create the job record.
4. Page through eligible items with server-side filters (`RecordType`, `WorkStatus`, optional `CompletedDate`). Pages hold at most 200 items and use key-based `$skiptoken` continuation, so deleting items doesn't skip any.
5. For each item:
   1. Skip items with attachments. They are logged once and kept.
   2. Reuse an existing archive copy (looked up by `OriginalItemId`) or create one.
   3. Read the copy back and compare `Title`, `WorkStatus`, `RecordType`, `DueDate`, `CompletedDate`, `WorkDescription` and `AssignedToKey`.
   4. Write an `Archived` audit entry. This write must succeed before the next step.
   5. Delete the source item.
6. Checkpoint the counters to the job record after every page.
7. Set the final status: `Completed`, `CompletedWithErrors`, `Failed` (the error is recorded in full), or `Abandoned` (stopped by the user).

This is a compensating workflow, not a transaction across the two lists. A copy can stay in the archive if a later step fails. The next run reuses that copy instead of creating a duplicate.

## Query and throttling

- GET, MERGE and DELETE requests are retried on 429 and 503 responses, honouring `Retry-After`, up to 6 attempts. A 404 on a retried DELETE counts as success.
- Create POSTs are never retried automatically, to avoid duplicate items.
- Indexed filters reduce the risk of hitting the list view threshold but don't remove it. The first filter clause (`RecordType eq 'Task'`) must stay selective. Test with realistic volumes.

## Security

Lists inherit the site's permissions; this version has no per-workspace security boundary. The UI hides or disables actions the user can't perform, using the page context's `SPPermission` checks. That is a usability measure only: SharePoint enforces the real permissions on every REST call.

## Explicit non-goals of this release

Board and task UI, workspace-specific permissions, attachment transfer, related-record archival, restore, server-side or unattended scheduling, atomic distributed locking, and cross-site-collection provisioning.
