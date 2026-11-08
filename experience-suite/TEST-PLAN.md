# SharePoint Experience Suite – Validation Plan

## Automated checks

Run `npm run build` from the solution root (it runs lint, Jest and packaging). The Jest suites cover helpers, Search mapping, the REST service (against a fake client) and a full UI render; see `BUILD-STATUS.md`.

## Tenant smoke tests

Use at least two standard employee accounts, one approver account and one site owner account.

1. **Shell/MySpace:** confirm the signed-in name/email loads, navigation works, refresh reloads all module data and responsive layout works on mobile widths.
2. **RequestHub:** submit a request with title/category/description; verify the SharePoint item fields and status; confirm the requester sees their own item.
3. **DocuSmart:** search for a known document; open a result; confirm inaccessible content is not returned.
4. **SiteNavigator:** search for a known site; confirm only security-trimmed search results appear.
5. **ApprovalFlow:** create a Pending item assigned to the approver email; approve and reject separate test items; verify status/decision fields; verify a different user does not see the assigned item.
6. **TeamPulse:** verify announcements; change a task status; confirm list item update and permissions enforcement.
7. **PeopleConnect:** search a known employee by name, department and a configured skill; verify profile fields; test disabled setting and tenants with empty Skills.
8. **Permissions:** test with users who have read-only access, no access to one list, and no permission to edit. Confirm writes fail safely and no privilege escalation occurs.
9. **Concurrency:** open the same approval in two sessions and attempt decisions; confirm stale ETag updates are rejected where SharePoint returns an ETag.
10. **Throttling:** verify list reads are capped at 100 and search results at 50; avoid rapid repeated broad wildcard queries.

## Release acceptance

- Production build exits successfully.
- A real `.sppkg` exists in `sharepoint/solution/` and uploads to the App Catalog.
- No unexpected API permission request appears during deployment.
- Each module has been tested against the target tenant's actual list schema, search configuration and permissions.
