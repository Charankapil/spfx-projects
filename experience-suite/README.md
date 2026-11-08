# SharePoint Experience Suite

An integrated employee-facing SharePoint Framework (SPFx) web part containing seven modules:

1. **MySpace** – personal landing page, signed-in profile, quick navigation and workspace metrics.
2. **RequestHub** – submit service requests and view requests submitted by the signed-in employee.
3. **DocuSmart** – discover documents through SharePoint Search, respecting the current user's access.
4. **SiteNavigator** – discover sites and webs returned by SharePoint Search.
5. **ApprovalFlow** – view approvals assigned to the signed-in user and approve/reject pending items.
6. **TeamPulse** – read team announcements and update task status.
7. **PeopleConnect** – search people by name, department, job title or searchable profile expertise.

The product is a single SPFx web part with shared navigation and a responsive light-theme UI. It uses the current SharePoint session via `WebPartContext` and `SPHttpClient`; there is no custom Entra ID app registration, client ID, Graph permission or app-only access in this project.

## Package status

Version 1.0.1 builds cleanly with the official SPFx 1.24.1 Heft toolchain: ESLint (SPFx React profile) reports no errors or warnings and all 38 Jest tests pass. A pre-built package is in [`releases/sharepoint-experience-suite.sppkg`](./releases/sharepoint-experience-suite.sppkg). It still needs validation on a real tenant (see `TEST-PLAN.md`) before production use. See [`REVIEW.md`](./REVIEW.md) for the issues found in the 1.0.0 source and how each was fixed.

Build output path:

```text
sharepoint/solution/sharepoint-experience-suite.sppkg
```

## Requirements

- SharePoint Online modern site.
- Node.js 22 (`>=22.14.0`) or Node.js 24, as supported by SPFx 1.24.1.
- npm and network access to the npm registry for first-time dependency installation.
- Permission to add an SPFx app to the target App Catalog and add the web part to a page.
- A site owner to provision the four backing lists and review list permissions.

The project currently pins SPFx `1.24.1`. Verify the supported Node/SPFx combination against Microsoft's SPFx compatibility documentation for your tenant's approved toolchain before upgrading or changing the pinned versions.

## Build and test

From the project root:

```bash
npm install
npm test
npm run build
```

`npm run build` runs lint, TypeScript, Jest and the production bundle (`heft test --production`) and then packages the solution. Confirm that the following file exists after a successful build:

```text
sharepoint/solution/sharepoint-experience-suite.sppkg
```

For local workbench/debugging, set `initialPage` in `config/serve.json` (or the `SPFX_SERVE_TENANT_DOMAIN` environment variable), then run:

```bash
npm start
```

A local workbench does not replace validation on a real SharePoint page. Search, profile properties, list permissions and REST writes must be tested in the target tenant.

## Provision the backing lists

Install PnP.PowerShell in an administrator/owner workstation if it is not already available:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

Run the included idempotent provisioning script as a site owner:

```powershell
./scripts/Provision-ExperienceSuite.ps1 -SiteUrl https://contoso.sharepoint.com/sites/EmployeeHub -ClientId <your-pnp-app-client-id>
```

Since September 2024 PnP.PowerShell no longer ships a shared sign-in app, so interactive sign-in needs an Entra ID app of your own. Create it once with `Register-PnPEntraIDAppForInteractiveLogin` and pass its client ID, or set `ENTRAID_APP_ID`. This app is only used by the provisioning script; the web part itself does not need one.

The script adds new columns to each list's default view and gives Status columns a default value (`Submitted`, `Pending` or `Not Started`).

It creates these lists and columns if missing:

### ES Requests

- Title (built-in)
- Description (multiple lines)
- Category (text)
- Status (Submitted, Assigned, In Progress, Completed, Rejected)
- RequesterName (text)
- RequesterEmail (text)
- Owner (text)
- DueDate (date/time)

### ES Approvals

- Title (built-in)
- Description (multiple lines)
- Category (text)
- Status (Pending, Approved, Rejected)
- Owner (text; may contain approver email)
- ApproverEmail (text)
- DueDate (date/time)
- DecisionBy (text)
- DecisionDate (date/time)

### ES Announcements

- Title (built-in)
- Description (multiple lines)
- Category (text)
- Owner (text)

### ES Tasks

- Title (built-in)
- Description (multiple lines)
- Category (text)
- Status (Not Started, In Progress, Completed)
- Owner (text; use the assigned employee's email for task assignment)
- DueDate (date/time)

The application treats a missing list (HTTP 404) as an empty module. Other errors on a list (for example, access denied) appear as a warning naming that list, and the other modules keep working.

## Configure data and permissions

1. Add the web part to a modern page on the site where the lists were provisioned.
2. Add request items through RequestHub. The app writes `RequesterName`, `RequesterEmail`, `Category`, `Description` and initial `Status`.
3. Populate ES Approvals with `ApproverEmail` set to the intended approver's email. The approval UI only shows approvals assigned to the signed-in user's email. Decisions update `Status`, `DecisionBy` and `DecisionDate`.
4. Populate ES Tasks with `Owner` set to the assignee's email. TeamPulse shows the tasks whose `Owner` matches the signed-in user's email and lets them update status where SharePoint permissions allow.
5. Add announcements to ES Announcements using SharePoint list forms or a governed publishing process.
6. Review list permissions and item-level permissions. The app does not elevate access; SharePoint enforces authorization for every read and write.

**Important governance note:** The current UI filters requests and approvals for convenience; client-side filtering is not a security boundary. Configure SharePoint list/item permissions so employees cannot read other employees' confidential requests or update approvals they do not own. For sensitive workflows, enforce assignment and transition rules in a trusted server-side workflow (for example, Power Automate with appropriately scoped permissions), not only in browser code.

## Search and PeopleConnect

- DocuSmart uses SharePoint Search with `IsDocument:1` and a 50-result cap.
- SiteNavigator searches for `STS_Site` and `STS_Web` results with a 50-result cap.
- PeopleConnect queries the **Local People Results** result source (`b09a7990-05ea-4af9-81ef-edfab16c4e31`) and requests `Title`, `Path`, `WorkEmail`, `Department`, `JobTitle`, `BaseOfficeLocation`, `OfficeNumber`, `PictureURL` and `Skills`.
- Search results respect the current user's SharePoint search security trimming. Search schema and user profile configuration determine which fields are populated.
- The `Skills` managed property may be empty or not queryable in a tenant until profile data is populated and the search schema is configured. No user directory scraping or tenant-wide account enumeration is performed.
- PeopleConnect can be disabled through the web part property pane.

## Security and operational boundaries

- The web part uses no app-only token, custom app registration, client ID, Microsoft Graph permission or tenant administrator API. (Only the PnP provisioning script needs a sign-in app; see above.)
- SharePoint REST calls use the signed-in user's existing SharePoint context.
- Search result count is capped at 50; list reads are capped at 100 and sorted by Modified descending.
- Approval and task writes use the current user's SharePoint permissions and the item's ETag, so a conflicting change returns a "changed by someone else" message rather than overwriting. The UI checks the assigned email and pending status, but list permissions/workflow policy must remain the actual security control.
- No automatic tenant-wide site enumeration is attempted. SiteNavigator searches only results returned by SharePoint Search.
- Avoid storing sensitive employee data in searchable profile properties. Follow your organization's HR/privacy policy.

## Deployment

1. Run `npm install` and `npm run build`, or use the pre-built package in `releases/`.
2. Verify `sharepoint/solution/sharepoint-experience-suite.sppkg` exists and the build completed without errors.
3. Upload the `.sppkg` to the tenant App Catalog (or an approved site collection App Catalog).
4. Deploy the package. This solution does not request tenant-wide API permissions.
5. Provision the lists with the script above and review permissions.
6. Add **SharePoint Experience Suite** to a modern SharePoint page.
7. Validate all modules with a standard employee account and a site owner account.

## Project layout

```text
config/                         SPFx Heft, serve, TypeScript and package config
scripts/Provision-ExperienceSuite.ps1  Creates backing lists and fields
src/webparts/experienceSuite/   SPFx web part, React UI, models and REST service
src/webparts/experienceSuite/test/  Jest tests (helpers, search mapping, service, UI render)
config/jest.config.json         Maps SPFx runtime packages and CSS to test stubs
releases/                       Pre-built .sppkg
sharepoint/solution/            Generated .sppkg output after a successful build
```

## Known tenant-dependent behavior

- Search may return no people if the tenant has disabled people search or does not expose the relevant profile properties.
- Skills search depends on the tenant's profile and search schema; the app cannot populate authoritative employee skills automatically.
- Approvals require an upstream process to create approval records and assign `ApproverEmail`.
- Team tasks and announcements require data owners or an upstream process to maintain the lists.
- Notifications, SLA escalation, attachment upload and approval workflow orchestration are not part of this SPFx client release; integrate them through an approved Power Automate flow if required.
