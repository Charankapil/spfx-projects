# Vantage security and permissions

## Identity

Vantage runs as the signed-in user, through SPFx `SPHttpClient`. It requests no Microsoft Graph or other API permissions and needs no app registration. SharePoint enforces every permission on every request. The UI only hides or disables actions the user can't perform; it is not a security boundary.

## Workspace security boundary

When a workspace is created (or later, on **Access → Secure workspace**), Vantage does the following:

1. Creates three SharePoint groups: `VNT_<key>_Admins`, `VNT_<key>_Members` and `VNT_<key>_Readers`. Their descriptions read "Vantage · <name> · <role>".
2. Adds the creator to Admins.
3. Breaks permission inheritance on the workspace's active and archive lists (`copyRoleAssignments=false`) and grants:

   | Principal | Role on both lists |
   |---|---|
   | Site Owners group | Full Control |
   | Workspace Admins | Full Control |
   | Workspace Members | Contribute |
   | Workspace Readers | Read |

4. Marks the workspace as `PermissionsConfigured` in the registry.

All three steps need **Manage Permissions** on the site, so creating a workspace needs a site owner. If securing fails, the workspace is still created. The UI then shows a persistent warning that the workspace uses the site's permissions until a site owner secures it.

People who are only in workspace groups also need **read access to the site** to open the page and the workspace and board registries. The Access tab can add them to the site's Visitors group in the same step, and does so by default.

## Control lists

| List | Permissions |
|---|---|
| `VNT_Settings`, `VNT_Workspaces`, `VNT_Boards` | Inherit the site's permissions. They contain names and keys only, no task content. |
| `VNT_ArchivePolicies`, `VNT_ArchiveJobs`, `VNT_ArchiveLog`, `VNT_ArchiveLocks` | Unique permissions: site Owners have Full Control; each workspace's Admins group has Contribute. |
| Each workspace's policy row | Item-level unique permissions: site Owners and that workspace's Admins only. |

**Design decision.** The 1.0.0 source gave *every* job, log and lock row its own unique permissions. That costs about 5 extra REST calls per row, and the number of uniquely-permissioned items would pass SharePoint's recommended maximum of 5,000 unique scopes per list after a few thousand archived tasks. That degrades every query on those lists. Vantage 1.0 secures these lists at list level instead, and writes only IDs, counts and outcomes into job and log rows, never task content.

Consequently, workspace administrators can read the job and log rows of other workspaces, but not their tasks or policies. Treat workspace administrators as trusted operators. If that is not acceptable, host sensitive workspaces on a separate site.

## Data handling

- Assignees are stored as claims login (`AssignedToKey`) plus display name (`AssignedToName`). These fields grant no access.
- Notifications live in the workspace's own list, so only people with access to the workspace can read them. All workspace members can technically read each other's notifications; the UI shows each user only their own.
- Archiving copies attachments of up to 25 MB each. Every copy is verified by SHA-256 checksum before the source is deleted.
- Restore deletes archive copies only after the restored records are verified.

## Operational guidance

- Test on a non-production site first ([TEST_PLAN.md](TEST_PLAN.md)).
- Keep Microsoft 365 retention or backup for business-critical records. Vantage's archive isn't a backup.
- Review **Job history** and the **Audit log** after archive runs.
