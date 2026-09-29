# Work Boards: plan and architecture

Work Boards replaces monday.com work management with an SPFx solution that
runs inside SharePoint Online. It keeps all data in SharePoint lists on the
site where it is installed.

> **Status:** this is a plan for review. No solution code has been written yet.
> The wireframes are in [`wireframes.html`](./wireframes.html).

---

## 1. Goals and constraints

| # | Requirement | How the design meets it |
|---|-------------|-------------------------|
| G1 | Cover at least 80% of the monday.com features the organisation uses | Feature map in section 3. Phase 1 and 2 cover the core; phase 3 adds the extras |
| G2 | Deploy to any site collection | One `.sppkg`, deployed from the tenant or site collection App Catalog. Each site is its own workspace |
| G3 | Create its own lists and libraries | A setup wizard in the web part creates and upgrades every list over SharePoint REST, with versioned migrations |
| G4 | **No app registration, no client id, no Graph permissions** | All calls go through SPFx's built-in `SPHttpClient` (through PnPjs `@pnp/sp` with the SPFx behaviour) against `/_api` on the current site. No `AadHttpClient`, no `MSGraphClient`, no `webApiPermissionRequests` in `package-solution.json` |
| G5 | Lower cost | Uses licences the organisation already pays for (Microsoft 365 and SharePoint Online). No per-seat fee and no Azure resources |

Other rules we hold to:

- Everything runs **as the signed-in user**, so SharePoint permissions decide
  what each person can see and do. There is no service account and no
  elevated backend.
- **No PnP PowerShell in the install path.** Since 2024 PnP PowerShell needs
  its own Entra ID app registration, which would break G4. Setup runs from
  the browser instead.
- **No `SP.Utilities.Utility.SendEmail`.** Microsoft has announced that this
  REST API is being retired, so the solution does not depend on it (see 7.4).

---

## 2. monday.com concepts mapped to SharePoint

| monday.com | Work Boards | SharePoint storage |
|------------|-------------|--------------------|
| Account | Tenant | Tenant App Catalog |
| Workspace | Workspace = one site collection (or subsite) | Site `/sites/marketing` |
| Folder | Board folder | `Folder` field on the board registry |
| Board | Board | One **hidden list** per board (`WB_Board_<key>`), plus one registry row in `WB_Boards` |
| Group | Group | Group definitions in the board config JSON. Each item has an indexed `WB_GroupId` field |
| Item | Item | List item |
| Subitem | Subitem | List item in the same list, with `WB_ParentId` set |
| Column | Column | A real SharePoint field on the board list, plus display metadata (colours, labels, width) in the board config |
| Update (comment) | Update | Item in the board's companion list `WB_Board_<key>_Updates` |
| Files column | Files | List item attachments (inherit board permissions) |
| Activity log | Activity | **SharePoint version history** of the item, which records field-level changes for free |
| Views | Views | `WB_Views` list (shared) or `WB_UserPrefs` (personal) |
| Automations | Automations | `WB_Automations` list (rules as JSON), run by the client-side rule engine |
| Dashboards | Dashboards | `WB_Dashboards` list (widget layout as JSON) |
| Forms | Forms | `WB_Forms` definitions + `WB_FormInbox` intake list |
| Templates | Templates | Built-in templates in the bundle + `WB_Templates` library (JSON files) |
| Inbox / Notifications | Inbox | Computed from @mentions and assignments (pull model, see 7.3) |
| My Work | My Work | Cross-board query of people columns, using `$batch` |
| Board owner / member / viewer | Board owner / member / viewer | SharePoint permission levels on the board list |
| Guest | Guest | SharePoint external sharing (if the tenant allows it) |

---

## 3. Feature coverage

Priority reflects typical monday.com usage. We should confirm it against the
organisation's own usage in phase 0.

| Area | monday.com feature | Phase | Notes |
|------|--------------------|:-----:|-------|
| **Structure** | Workspaces, folders, boards, groups, items, subitems | 1 | |
| | Board templates (Project, Tasks, CRM pipeline, Content calendar, Bug tracker, Onboarding, Blank) | 1 | Also "Save board as template" |
| | Duplicate, archive and restore a board or group | 1 | Archive = flag. Delete goes to the SharePoint recycle bin |
| **Columns** | Status, Priority, Text, Long text, Numbers, Date, People, Dropdown, Checkbox, Link, Email, Phone, Tags, Files | 1 | |
| | Timeline, Dependency, Rating, Progress (from status), Last updated, Creation log, Item ID | 2 | |
| | Formula, Time tracking, Connect boards, Mirror, Vote, Country, Week | 3 | Formula runs client-side from a safe expression parser, no `eval` |
| **Views** | Main table (inline edit, group collapse, drag to reorder, column resize and reorder, group summaries) | 1 | Virtualised rows |
| | Kanban (by any Status or Dropdown column) | 1 | |
| | Filter, sort, search, person filter, hide columns, group by | 1 | |
| | Saved views (shared and personal) | 2 | |
| | Calendar | 2 | |
| | Timeline / Gantt (with dependencies and drag to reschedule) | 2 | |
| | Chart view | 2 | |
| | Files gallery | 2 | |
| | Workload view | 3 | |
| **Item card** | Updates with rich text, @mentions, replies, likes | 1 | |
| | Files (upload, preview in SharePoint viewer) | 1 | |
| | Activity log (who changed what, from and to) | 1 | Read from version history |
| **Collaboration** | Inbox, bell with unread count | 2 | |
| | My Work (everything assigned to me, across boards, by date) | 1 | |
| | Near-live refresh of other people's edits | 1 | Change-token polling (7.5) |
| | Favourites and recent boards | 1 | |
| **Automations** | "When status changes to X, then …" / "When item created …" / "When column changes …" / "When person assigned …" | 2 | Actions: set column, assign person, move to group, notify, create item on another board, create subitem, archive |
| | "When date arrives" / "Every week, create item" (recurring) | 3 | Needs a scheduler (7.2) |
| | Email notifications | 3 | Optional Power Automate template (7.4) |
| **Dashboards** | Widgets: numbers, chart (bar, pie, line), battery (status mix), workload, table, timeline, text. Across several boards | 2 | Also as a separate web part for any page |
| **Forms** | Form view for a board, with a shareable page and intake review | 2 | Submitters do not need access to the board |
| **Import / export** | Import from Excel or CSV, including monday.com's own Excel export | 1 | Column mapping wizard, creates status labels |
| | Export board to Excel (CSV) | 1 | |
| | Migrate updates and files from monday.com | 3 | Optional one-off Node script with the user's own monday.com API token |
| **Admin** | Board permissions (main or private), board owners, read-only viewers | 1 | |
| | Workspace settings, schema upgrade, usage stats | 1 | |
| **Anywhere** | Teams tab | 3 | SPFx web parts can run as Teams tabs without any app registration |
| | Mobile | 1 | Responsive layout. Works in the SharePoint mobile app and mobile browser |

**Not covered, on purpose:** monday.com's AI features, its integrations
marketplace (Power Automate covers most of these), WorkDocs (use Word or
Loop), native mobile apps, the CRM and Dev product add-ons, live cursors, and
guaranteed to-the-minute time-based automations without Power Automate.

By count of commonly used features, phases 1 and 2 cover roughly 80%. Phase
3 takes it to about 90%.

---

## 4. Solution architecture

```
┌──────────────────────── SharePoint Online site (one workspace) ────────────────────────┐
│                                                                                        │
│  Modern page "Work Boards" (single-part app page, full width)                          │
│  ┌─────────────────────────────── SPFx bundle (React 17, Fluent UI 8) ───────────────┐ │
│  │  Web parts: WorkBoards (app) · Board (embed) · Dashboard · Form                   │ │
│  │                                                                                   │ │
│  │  UI layer ── hash router (#/board/12/item/345) · views · item panel · wizards     │ │
│  │  State ───── board store (items, columns, groups) · optimistic updates · undo     │ │
│  │  Services ── BoardService · ItemService · UpdateService · ViewService             │ │
│  │              AutomationEngine · InboxService · ImportExport · PermissionService   │ │
│  │  Platform ── Provisioner (migrations) · ChangePoller · ETag concurrency            │ │
│  │              PnPjs @pnp/sp  →  SPFx SPHttpClient  (signed-in user, no tokens)     │ │
│  └───────────────────────────────────────┬───────────────────────────────────────────┘ │
│                                          │ /_api (REST, $batch, RenderListDataAsStream)│
│  Lists (hidden) ─────────────────────────▼─────────────────────────────────────────────│
│  WB_Meta            schema version, settings        WB_Views         saved views       │
│  WB_Boards          board registry + config JSON     WB_Automations   rules            │
│  WB_Board_<key>     items of one board (×N)          WB_Dashboards    widget layouts   │
│  WB_Board_<key>_Updates  comments of that board (×N) WB_Forms / WB_FormInbox           │
│  WB_UserPrefs       favourites, last seen (own rows only)  WB_Templates (library)      │
└────────────────────────────────────────────────────────────────────────────────────────┘
        Optional, outside the bundle: Power Automate flow templates for email digests
        and date-based automations (standard connectors, run under the maker's account).
```

### 4.1 Why one list per board

We looked at three storage models:

| Option | Pros | Cons |
|--------|------|------|
| A. One central Items list, column values in JSON | Easy cross-board queries, simple setup | Cannot filter or sort on JSON server-side. Private boards would need item-level permissions (slow, limited). The 5,000-item view threshold hits every board together |
| B. **One list per board, real fields per column** (chosen) | Real SharePoint fields: server-side filter and sort, indexes, version history per field, Excel and Power BI connect directly. Board permissions = list permissions. Each board gets its own 5,000 threshold | Cross-board queries need a fan-out (solved with `$batch`). Creating a board needs the Manage Lists permission, which site Members have by default |
| C. Per-board list plus a central index list | Fast cross-board queries | Index can drift out of sync, and it leaks titles of private boards |

Option B uses what SharePoint is good at. Version history gives us the
activity log. List permissions give us private boards. Excel, Power BI and
Power Automate can all read the data without anything extra.

### 4.2 Fixed fields on every board list

Every board list gets these fields from a site content type, `WB Work Item`.
Cross-board features (My Work, dashboards, automations) can rely on their
internal names.

| Field | Type | Indexed | Purpose |
|-------|------|:------:|---------|
| `Title` | Text | | Item name |
| `WB_GroupId` | Text(32) | ✓ | Group the item belongs to |
| `WB_SortOrder` | Text(64) | ✓ | Fractional order key, so moving one item writes one row |
| `WB_ParentId` | Number | ✓ | Subitem parent (empty for top-level items) |
| `WB_Owner` | Person (multi) | ✓ | The board's main "People" column, used by My Work |
| `WB_Status` | Choice | ✓ | The board's main "Status" column |
| `WB_DueDate` | DateTime | ✓ | The board's main date column, used by My Work and calendar |
| `WB_Archived` | Yes/No | ✓ | Archived items |
| `WB_ItemKey` | Text | | Readable item ID, for example `MKT-142` |

`Created`, `Author`, `Modified` and `Editor` provide the Creation log and
Last updated columns.

### 4.3 User-added columns

Each column a user adds becomes a real field named `WB_c_<6-char id>`. Its
display name, colour labels and width live in the board config.

| Column type | SharePoint field | Stored in board config |
|-------------|------------------|------------------------|
| Status / Priority | Choice | Label colours, "done" label, order |
| Text / Long text | Text / Note (plain or rich) | |
| Numbers | Number | Unit, decimals, summary (sum or avg) |
| Date | DateTime | Show time? |
| Timeline | Two DateTime fields (`_s`, `_e`) | |
| People | User (multi) | |
| Dropdown / Tags | MultiChoice | Colours |
| Checkbox | Boolean | |
| Link | URL | |
| Email / Phone | Text | Format validation |
| Rating | Number (0–5) | Max |
| Files | Item attachments | |
| Dependency | Lookup (multi) to the same list | Type (finish-to-start) |
| Connect boards | Lookup (multi) to another board list on the same site | Target board |
| Formula | Not stored. Computed in the browser | Expression |
| Progress | Not stored. Computed from status columns | Source columns |
| Time tracking | Number (seconds) + Text (running since) | |

Renaming a column only changes the config. Deleting a column removes the
field after the user confirms.

### 4.4 Board registry (`WB_Boards`)

| Field | Type | Notes |
|-------|------|-------|
| `Title` | Text | Board name |
| `BoardKey` | Text, indexed, unique | Short key used in list names, for example `MKT` |
| `ItemsListId`, `UpdatesListId` | Guid (Text) | |
| `Folder` | Text | Folder in the sidebar |
| `Privacy` | Choice: Main, Private | |
| `BoardOwners` | Person (multi) | Who can change columns, automations and permissions |
| `Config` | Note (JSON) | Columns, groups, labels, default view, board settings |
| `SchemaVersion` | Number | Config format version |
| `Icon`, `Color` | Text | |
| `IsArchived`, `IsTemplate` | Yes/No | |

A private board's registry row has its permissions broken the same way as
its lists, so people outside the board cannot see its name either.

### 4.5 The other site lists

| List | Purpose | Notable settings |
|------|---------|------------------|
| `WB_Meta` | Installed schema version, workspace settings, setup log | One row |
| `WB_Views` | Shared saved views: type, filters, sort, hidden columns, grouping, view-specific settings | |
| `WB_UserPrefs` | Favourites, recent boards, Inbox "last seen", personal views | **Item-level security: people read and edit only their own rows** |
| `WB_Automations` | Board, trigger, conditions, actions (JSON), enabled, run count, last error | Only board owners can edit |
| `WB_Dashboards` | Name, boards used, widgets (JSON) | |
| `WB_Forms` | Board, fields, questions, visibility, thank-you text | |
| `WB_FormInbox` | Form submissions (JSON) waiting to be copied onto the board | Submitters can add items and read only their own |
| `WB_Templates` | Document library of board templates (JSON) | |

---

## 5. Provisioning and deployment

### 5.1 Package

- One solution: `work-boards.sppkg`, SPFx 1.23 (latest GA at time of
  writing), Node 22, React 17, Fluent UI 8, PnPjs 4.
- `skipFeatureDeployment: true`, so an admin can make it available to every
  site at once. It also works from a site collection App Catalog if the
  organisation prefers to limit which sites get it.
- `webApiPermissionRequests` is empty, so there is nothing to approve in the
  API access page.
- Web part manifests set `supportsFullBleed` and `SharePointFullPage` (so the
  app can be a full-page app page) and `TeamsTab` (phase 3).

### 5.2 First run on a site (setup wizard)

1. A site owner adds the **Work Boards** app to the site (or it is already
   there from tenant-wide deployment) and creates a page with the Work Boards
   web part. The wizard can also create the page itself.
2. The web part reads `WB_Meta`. If it is missing:
   - A user with **Manage Web** permission sees the setup wizard.
   - Anyone else sees "Work Boards isn't set up on this site yet. Ask a site
     owner." with the owners listed.
3. The wizard shows what it will create, then runs the migrations in order,
   with a progress list. Each step is idempotent (it checks before it
   creates), so a failed or interrupted setup can simply run again.
4. Optionally the wizard creates starter boards from templates and imports
   Excel files exported from monday.com.

### 5.3 Upgrades

Each release ships a list of numbered migrations (`001_core_lists`,
`002_add_views_list`, `003_board_field_WB_ItemKey`, …). When a site owner
opens the app and `WB_Meta.SchemaVersion` is behind, a banner offers
"Update Work Boards on this site". Board-level migrations (a new fixed field
on every board list) run the same way, board by board. Other users keep
working on the older schema until an owner upgrades; the code reads both.

### 5.4 Creating a board

1. The user chooses a template or blank board and a privacy setting.
2. The client creates the lists `WB_Board_<KEY>` and `WB_Board_<KEY>_Updates`
   (hidden, versioning on, attachments on), adds the `WB Work Item` content
   type and the template's columns, and creates indexes.
3. For a private board it breaks permission inheritance on both lists and the
   registry row and grants the selected people Edit or Read.
4. It writes the registry row last, so an interrupted create leaves no
   half-built board visible. A cleanup check lists orphaned lists for owners.

Permissions needed: Members can create main boards (Manage Lists is part of
the default Edit level). Only owners can create private boards, because
breaking inheritance needs Manage Permissions. A member can ask an owner to
make a board private.

---

## 6. Permissions

| Role | SharePoint mapping | Can |
|------|-------------------|-----|
| Workspace admin | Site Owners | Everything, including setup and upgrades |
| Board owner | Listed in `BoardOwners` and has Edit on the board list | Change columns, groups, automations, board permissions |
| Board member | Edit (or Contribute) on the board list | Add and edit items, post updates |
| Viewer | Read on the board list | See the board. The UI hides all edit controls |
| Form submitter | Add on `WB_FormInbox` only | Submit a form |

SharePoint enforces every rule on the server. The UI only hides what the
user can't do anyway, based on `EffectiveBasePermissions`.

---

## 7. Key technical designs

### 7.1 Fast board loading

- One `RenderListDataAsStream` call per board view (up to 5,000 rows per
  page, with paging tokens for more). All filters use indexed fields so large
  boards stay under the list view threshold.
- The table and Kanban views are virtualised, so boards with 10,000+ items
  still scroll smoothly.
- The last loaded board is cached in IndexedDB, so it appears straight away
  and then refreshes.

### 7.2 Automations engine

Rules are stored as JSON:

```json
{
  "trigger":   { "type": "columnChanged", "column": "WB_Status", "to": "Done" },
  "conditions":[ { "column": "WB_c_x7k2p1", "op": "eq", "value": "High" } ],
  "actions":   [ { "type": "moveToGroup", "groupId": "g_done" },
                 { "type": "notify", "people": "column:WB_Owner",
                   "text": "{item} is done" } ]
}
```

- **Event triggers** (item created, column changed, status changed, person
  assigned, update posted) run in the browser straight after the user's own
  change succeeds. Every change goes through the Work Boards UI, so every
  event is seen. The actions run as that user, so they can never do more
  than the user is allowed to. A guard stops rules from triggering each
  other in a loop (maximum depth 3).
- **Changes made outside the app** (for example directly in the SharePoint
  list or through Excel) are picked up by the change poller the next time a
  board member has the board open. They are marked as "applied late" in the
  automation log.
- **Time triggers** (date arrives, recurring items) run in a "lazy
  scheduler": when anyone opens the workspace, the client checks for due
  rules and claims each run with an ETag-guarded write on the rule, so two
  browsers never run it twice. For exact timing, an optional Power Automate
  scheduled flow can call the same rules (phase 3).

### 7.3 Inbox and notifications (pull model)

Notifications are not copied into a central list, because that list would
be readable by everyone on the site and would leak private board content.
Instead the Inbox is **computed** for the signed-in user:

- Updates where they are @mentioned (`Mentions` field on each board's
  Updates list).
- Items where they are in a People column and that someone else changed
  since they last looked (`Modified` after "last seen", `Editor` not them).
- Replies to their updates.

Queries go to every board the user can access, in one `$batch` request.
SharePoint returns only what they are allowed to see. The automation
"notify" action posts an update that mentions the person, so it lands in the
same Inbox. The last-seen time is kept in `WB_UserPrefs`.

### 7.4 Email

Microsoft has announced that `SP.Utilities.Utility.SendEmail` is being
retired, and the Graph `sendMail` API would need permissions we have ruled
out. So email is optional and uses Power Automate:

- A **flow template** (standard SharePoint and Outlook connectors, no premium
  licence) sends a daily or hourly digest of new mentions and assignments.
  It runs under the account of whoever imports it. That is a normal Microsoft
  365 sign-in, not an app registration.
- People can also use SharePoint's own "Alert me" on any board list.

### 7.5 Near-live updates and conflicts

- While a board is open and the tab is visible, the client asks SharePoint
  for changes since its last change token every 15 seconds. This is one
  cheap call (`GetChanges`), and it backs off when the tab is hidden or the
  user is idle.
- Every edit sends the item's ETag. If someone else changed that item in the
  meantime, the client merges field by field when the changes don't overlap,
  and otherwise shows "Anna changed Status to Done a moment ago" with a
  choice.
- Edits show immediately (optimistic update) and roll back with a message if
  the save fails.

### 7.6 Throttling and scale

- All bulk operations go through PnPjs batching (up to 100 operations per
  `$batch`) with a small concurrency limit.
- 429 and 503 responses are retried with backoff, honouring `Retry-After`,
  the same way the File Type Analyser in this repository does.
- Guidance: hundreds of boards per site and up to about 20,000 items per
  board. Larger programmes should be split across sites (workspaces).

### 7.7 Security

- Rich text in updates is sanitised with DOMPurify before it is saved and
  again before it is shown.
- Formulas use a small expression parser with a fixed set of functions. There
  is no `eval` or `new Function`.
- No secrets, tokens or external calls. The only network traffic is to the
  site's own `/_api`.

---

## 8. Code layout (planned)

```
work-boards/
  config/                      SPFx config (package-solution.json: no API permissions)
  src/
    webparts/
      workBoards/              Full app (router, shell, sidebar)
      boardEmbed/              Single board or view on any page
      dashboard/               Dashboard on any page
      boardForm/               Intake form on any page
    app/
      shell/                   Sidebar, top bar, search, workspace home
      board/                   Board header, toolbar, filters, groups
      views/                   table/, kanban/, calendar/, timeline/, chart/, files/, workload/
      columns/                 One folder per column type: cell renderer, editor, filter, SP field mapping
      item/                    Item panel: updates, files, activity
      automations/             Rule builder UI
      dashboards/              Widgets
      forms/                   Form builder and form renderer
      import/                  Excel/CSV import wizard, export
      setup/                   Setup wizard, upgrade banner
    services/
      sp/                      PnPjs setup (SPFx behaviour), batching, retry
      provisioning/            Migration runner, migrations/NNN_*.ts
      boards/ items/ updates/ views/ automations/ inbox/ permissions/
      realtime/                Change-token poller
    engine/
      automation/              Triggers, conditions, actions, loop guard, scheduler
      formula/                 Expression parser and functions
      ordering/                Fractional sort keys
    templates/                 Built-in board templates (JSON)
  tests/                       Jest unit tests for engine/, services/ (with a fake SP)
  docs/                        This plan, wireframes, admin guide, user guide
  tools/monday-export/         (Phase 3) Optional Node script for full monday.com migration
  releases/work-boards.sppkg
```

Main libraries (all MIT or similar): `@pnp/sp`, `@fluentui/react` 8,
`@dnd-kit` (drag and drop), `react-window` (virtualisation), `chart.js`
(charts, loaded only when needed), `dompurify`, `date-fns`, `xlsx` from
SheetJS's own CDN build (import and export, loaded only when needed).
Timeline and calendar are custom-built to keep the bundle small.

---

## 9. Delivery plan

| Phase | Scope | Result |
|-------|-------|--------|
| **0. Discovery** | Export a list of the organisation's monday.com boards, column types, automations, integrations and user counts. Confirm the priorities in section 3 | Agreed scope |
| **1. Core (MVP)** | Setup wizard and migrations, workspace home, boards, groups, items, subitems, core columns, table and Kanban views, filter/sort/search, item panel (updates, files, activity), My Work, templates, Excel import/export, main and private boards, change polling | Teams can move day-to-day boards off monday.com. Pilot with one or two teams |
| **2. Parity** | Calendar, timeline/Gantt, dependencies, timeline and extra columns, saved views, chart and files views, Inbox, event automations, dashboards and dashboard web part, forms and form web part, board embed web part | ~80% coverage. Wider rollout |
| **3. Extras** | Time-based automations, recurring items, formula, time tracking, connect and mirror boards, workload view, Teams tab, email digest flow template, full monday.com migration tool | ~90% coverage. Cancel monday.com |

Each phase ends with a tested `.sppkg` in `releases/`, an updated admin
guide, and a pilot on a test site before production.

**Testing:** Jest unit tests for the automation engine, formula parser,
ordering, column mapping and migrations (against an in-memory fake of the
SharePoint API). Playwright smoke tests against a developer tenant for the
main flows.

---

## 10. Cost comparison (illustrative)

monday.com is priced per seat. At list prices of roughly $12 to $19 per seat
per month on annual billing (Standard and Pro), 200 users cost about
**$29,000 to $46,000 a year**, before Enterprise add-ons. Please check these
figures against the actual contract.

Work Boards uses licences the organisation already has. The running cost is
maintenance time, plus Power Automate only if the optional email digests are
used, and those need only standard connectors.

---

## 11. Risks and mitigations

| Risk | Mitigation |
|------|-----------|
| Users expect monday.com's exact look and speed | Familiar layout (see wireframes), optimistic updates, cached board loads, pilot feedback before wide rollout |
| No server means time-based automations are not exact | Lazy scheduler for most cases, optional Power Automate flow where timing matters |
| Members can't create private boards | Owners create them, or a member requests one in the app |
| Very large boards | Indexed fields, paging, virtualisation, and guidance to split boards over ~20,000 items |
| Microsoft changes to SharePoint APIs | Only documented REST APIs are used, all behind one service layer |
| Losing monday.com history | Excel import in phase 1, full migration of updates and files in phase 3, and the monday.com export kept as an archive |

---

## 12. Questions for you

1. Roughly how many people and boards use monday.com today, and which
   features do they rely on most (automations, dashboards, forms,
   integrations, time tracking)?
2. Should Work Boards be available to every site (tenant-wide deployment)
   or only to chosen sites (site collection App Catalog)?
3. Is it acceptable for email notifications to be an optional Power Automate
   flow, with the in-app Inbox as the main channel?
4. How much history do you need to migrate: current items only, or updates
   and files too?
5. Do you use external guests on monday.com boards?
6. Is "Work Boards" a good name, or do you have one in mind?
