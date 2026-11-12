# Vantage 1.0 architecture

## Principles

1. SharePoint Online is the system of record. The web part uses one SPFx package and one SharePoint session, with no backend or app registration.
2. The **workspace** is the security boundary. Boards inherit the workspace's permissions.
3. All list reads are filtered on the server, indexed and bounded. Continuation uses key-based `$skiptoken` paging.
4. Destructive operations follow **copy → verify → audit → delete**, and are safe to retry.

## Runtime

`VantageWebPart` creates a `SharePointService` over `SPHttpClient` for the current web. It passes the user and the page-context permissions (`managePermissions`, `manageLists`, add and edit items) to `VantageApp`. `VantageApp` builds the services and shares them with the tabs through a React context.

REST calls use the OData `nometadata` format. Request bodies never contain `__metadata`, which SharePoint rejects in this format, and filter expressions are always URL-encoded. Claims logins contain `#` and `|`, which would otherwise truncate the query.

## Data model

Every field is created with `CreateFieldAsXml` and `AddFieldInternalNameHint`, so internal names are exact. `*` marks an indexed field, `U` a unique one.

### Control lists

| List | Fields |
|---|---|
| `VNT_Settings` | `Title` (key), `SettingValue`. Holds `SchemaVersion` (3 = 1.0.0). |
| `VNT_Workspaces` | `WorkspaceKey`\*, `WorkspaceStatus`\*, `WorkItemsList`, `ArchiveList`, `OwnerKey`, `PermissionsConfigured` |
| `VNT_Boards` | `BoardKey` U, `WorkspaceId`\*, `BoardStatus`\*, `BoardDescription` |
| `VNT_ArchivePolicies` | `WorkspaceId`\*, `Enabled`, `ArchiveMode`, `MaxActiveItems`, `ArchiveBefore`, `EligibleStatuses`, `BatchSize` |
| `VNT_ArchiveJobs` | `WorkspaceId`\*, `JobStatus`\*, `ArchiveMode`, `Started`, `Finished`, `Processed`, `Archived`, `Records`, `Skipped`, `FailedCount`, `LastError` |
| `VNT_ArchiveLog` | `WorkspaceId`\*, `OriginalItemId`\*, `ArchiveItemId`, `ArchiveJobId`\*, `Outcome`, `Detail` |
| `VNT_ArchiveLocks` | `WorkspaceId` U, `JobId`, `HeartbeatEpoch`, `LockOwner` |

### Per-workspace lists

Each workspace has `VNT_<key>_Items` (active) and `VNT_<key>_Archive`. Every record type lives in the same list, distinguished by `RecordType`:

| RecordType | Purpose | Key fields |
|---|---|---|
| `Task` | The unit of work | `WorkStatus`\*, `Priority`, `AssignedToKey`\* / `AssignedToName`, `DueDate`\*, `CompletedDate`\*, `Progress`, `WorkDescription`, `BoardId`\* |
| `Comment`, `Checklist`, `Dependency`, `Activity` | Child records of a task | `ParentTaskId`\*, `ChecklistDone`, `RelatedTaskId`\*, `ActorKey` / `ActorName` |
| `Notification` | In-app notification | `NotificationTo`\* (lower-case claims login), `NotificationRead`, `ParentTaskId` |

Both lists also carry custom slots: `CustomText1-8`, `CustomNumber1-4`, `CustomDate1-4` and `CustomBoolean1-2`.

- **Active list only:** `RestoredFromArchiveId` (U). It makes restore idempotent.
- **Archive list only:** `OriginalItemId` (U), `ArchivedAt`\*, `ArchiveJobId`, `OriginalCreated`, `OriginalModified`. `OriginalItemId` makes archiving idempotent.

The active list has 11 indexed columns and the archive list 12, well within SharePoint's limit of 20.

## Archive protocol

1. Validate the policy. Only Completed and Cancelled tasks are eligible.
2. **Lock:** insert a row into `VNT_ArchiveLocks`. Because `WorkspaceId` is unique, the insert is atomic: a second concurrent run is rejected by SharePoint. A lock whose heartbeat is older than 30 minutes is reclaimed, and its job is marked `Abandoned`.
3. Create the job row.
4. Run the date phase (`CompletedDate lt cutoff`) and then the count phase (oldest first, until `ItemCount` is at or below the maximum). Each page of tasks is processed as follows:
   1. Read the task's child records (`ParentTaskId`). Check attachment sizes first; any file over 25 MB means the task is **skipped** with nothing written.
   2. Copy the task and every child to the archive. An existing copy, matched by `OriginalItemId`, is reused. Then copy attachments and verify **every copied field** and the **SHA-256 checksum of every attachment**.
   3. Re-read the source. If `Modified` changed or new children appeared, keep everything for the next run.
   4. Write an `Archived` audit entry. This write must succeed before the next step.
   5. Delete the children, then the task.
5. After each page, checkpoint the counters and refresh the lock heartbeat. The user can stop gracefully, and the browser warns before leaving the page during a run.
6. Set the final status and release the lock.

## Restore protocol

Restore is refused while an archive run holds the lock. It is also refused if the original task is still active, which happens when an earlier run stopped before deleting it; in that case, run the archive again instead.

1. Recreate the task with `RestoredFromArchiveId`, or reuse it if it already exists.
2. Recreate its children with `ParentTaskId` pointing at the new task.
3. Copy attachments and verify fields and checksums.
4. Write a `Restored` audit entry.
5. Delete the archive copies.

The restored task gets a **new ID**. Dependencies on *other* tasks that pointed at the old ID are not rewritten.

## Scale and limits

- Each workspace uses 2 lists. Workspace creation is refused near the limit of 2,000 lists per site collection.
- Task lists show the 500 most recent tasks per board. Dashboard figures use the 2,000 most recent tasks; the UI says so when that cap is reached.
- Archiving runs in the browser and is sequential. Expect a few seconds per task, more with attachments. For very large backlogs, run in batches outside business hours.
- The archive log and job rows inherit list-level permissions. There are no per-item scopes, so the logs scale with volume (see SECURITY.md).

## Non-goals in 1.0

- Unattended or scheduled archiving (a Power Automate trigger is a possible extension).
- Kanban drag-and-drop, Gantt and calendar views.
- Custom-field designer.
- Cross-site-collection roll-up.
- Email or Teams notifications.
