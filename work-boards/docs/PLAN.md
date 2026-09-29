# Work Boards: plan and architecture

Work Boards replaces monday.com work management with an SPFx solution that
runs inside SharePoint Online. It keeps all data in SharePoint lists on the
site where it is installed.

> **Status:** Phase 1 (core) is built. See the [README](../README.md) for what it
> does and how to deploy it, and [`wireframes.html`](./wireframes.html) for the
> design. Section 12 records the decisions that shaped this version of the plan.

---

## 1. Goals and constraints

| # | Requirement | How the design meets it |
|---|-------------|-------------------------|
| G1 | Cover at least 80% of the monday.com features the organisation uses | Feature map in section 3. Phase 1 and 2 cover the core; phase 3 adds the extras |
| G2 | Deploy to chosen site collections | One `.sppkg` in the App Catalog with `skipFeatureDeployment: false`, so a site owner adds the app only to the sites that should use it. Each site is its own workspace |
| G3 | Create its own lists and libraries | A setup wizard in the web part creates and upgrades every list over SharePoint REST, with versioned migrations |
| G4 | **No app registration, no client id, no Graph permissions** | All calls go through SPFx's built-in `SPHttpClient` (a small wrapper, `SpClient`, adds retries) against `/_api` on the current site. No `AadHttpClient`, no `MSGraphClient`, no `webApiPermissionRequests` in `package-solution.json` |
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
| Account | Tenant | App Catalog |
| Workspace | Workspace = one site collection (or subsite) | Site `/sites/marketing` |
| Folder | Board folder | `Folder` field on the board registry |
| Board | Board | One **hidden list** per board (`WB_Board_<key>`), plus one registry row in `WB_Boards` |
| Group | Group | Group definitions in the board config JSON. Each item has an indexed `WB_GroupId` field |
| Item | Item | List item |
| Subitem | Subitem | List item in the same list, with `WB_ParentId` set |
| Column | Column | A real SharePoint field on the board list, plus display metadata (colours, labels) in the board config |
| Update (comment) | Update | Item in the board's companion list `WB_Board_<key>_Updates` |
| Files column | Files | List item attachments (inherit board permissions) |
| Activity log | Activity | **SharePoint version history** of the item, which records field-level changes for free |
| Views | Views | Phase 1: search, filters, sort and hidden columns remembered per person and board in the browser. Phase 2: shared saved views in `WB_Views` |
| Automations | Automations | `WB_Automations` list (rules as JSON), run by the client-side rule engine |
| Dashboards | Dashboards | `WB_Dashboards` list (widget layout as JSON) |
| Forms | Forms | `WB_Forms` definitions + `WB_FormInbox` intake list |
| Templates | Templates | Built-in templates in the bundle + `WB_Templates` library (JSON files) |
| Inbox / Notifications | Inbox | Computed from @mentions and assignments (pull model, see 7.3) |
| My Work | My Work | Cross-board query of each board's Owner column (`WB_Owner`), a few boards at a time |
| Board owner / member / viewer | Board owner / member / viewer | SharePoint permission levels on the board list |

---

## 3. Feature coverage

The department uses monday.com for work and project management, so the timeline
and Gantt view moved into phase 1. Nothing is migrated from monday.com (see 12).

| Area | monday.com feature | Phase | Notes |
|------|--------------------|:-----:|-------|
| **Structure** | Workspaces, folders, boards, groups, items, subitems | 1 ✓ | |
| | Board templates: Project plan, Team tasks, Team requests, Blank | 1 ✓ | "Save board as template" and more templates in phase 2 |
| | Archive and restore a board; delete board, group or item | 1 ✓ | Delete goes to the SharePoint recycle bin. Duplicate board: phase 2 |
| **Columns** | Status, Priority, Timeline, Date, People, Text, Long text, Numbers, Dropdown (tags), Checkbox, Link, Files | 1 ✓ | Priority is a Status column; files are item attachments |
| | Dependency, Rating, Progress (from status), Last updated, Creation log, Email, Phone | 2 | Item IDs (`KEY-12`) are shown in phase 1 |
| | Formula, Time tracking, Connect boards, Mirror, Vote, Country, Week | 3 | Formula runs client-side from a safe expression parser, no `eval` |
| **Views** | Main table (inline edit, group collapse, drag to reorder, column reorder, group summaries, bulk move and delete) | 1 ✓ | Column resize and row virtualisation: phase 2 |
| | Kanban (by any Status column) | 1 ✓ | |
| | Filter, sort, search, person filter, hide columns | 1 ✓ | Group by another column: phase 2 |
| | Saved views (shared and personal) | 2 | |
| | Calendar | 2 | |
| | Timeline / Gantt (days, weeks, months; drag to move and resize) | 1 ✓ | Dependency arrows: phase 2 |
| | Chart view | 2 | |
| | Files gallery | 2 | |
| | Workload view | 3 | |
| **Item card** | Updates with @mentions, replies, likes | 1 ✓ | Plain text in phase 1; formatting in phase 2 |
| | Files (upload, open, delete) | 1 ✓ | |
| | Activity log (who changed what, from and to) | 1 ✓ | Read from version history |
| **Collaboration** | Inbox, bell with unread count | 2 | |
| | My Work (everything assigned to me, across boards, by date) | 1 ✓ | |
| | Near-live refresh of other people's edits | 1 ✓ | Polling (7.5) |
| | Favourites and recent boards | 1 ✓ | |
| **Automations** | "When status changes to X, then …" / "When item created …" / "When column changes …" / "When person assigned …" | 2 | Actions: set column, assign person, move to group, notify, create item on another board, create subitem, archive |
| | "When date arrives" / "Every week, create item" (recurring) | 3 | Needs a scheduler (7.2) |
| | Email notifications | Later | Deferred at your request (7.4) |
| **Dashboards** | Widgets: numbers, chart (bar, pie, line), battery (status mix), workload, table, timeline, text. Across several boards | 2 | Also as a separate web part for any page |
| **Forms** | Form view for a board, with a shareable page and intake review | 2 | Submitters do not need access to the board |
| **Import / export** | Export board to Excel (CSV) | 1 ✓ | |
| | Import from Excel or CSV | 3 | Not needed to start: Work Boards starts fresh with no migration |
| **Admin** | Board permissions (main or private), board owners, read-only viewers | 1 ✓ | |
| | Setup and schema upgrade | 1 ✓ | Usage stats: phase 3 |
| **Anywhere** | Teams tab | 3 | SPFx web parts can run as Teams tabs without any app registration |
| | Mobile | 1 ✓ | Responsive layout. Works in the SharePoint mobile app and mobile browser |

**Not covered, on purpose:** monday.com's AI features, its integrations
marketplace (Power Automate covers most of these), WorkDocs (use Word or
Loop), native mobile apps, the CRM and Dev product add-ons, live cursors, and
guaranteed to-the-minute time-based automations without Power Automate.

✓ = built in phase 1. By count of commonly used features, phases 1 and 2
cover roughly 80%. Phase 3 takes it to about 90%.

---

## 4. Solution architecture

```
┌──────────────────────── SharePoint Online site (one workspace) ────────────────────────┐
│                                                                                        │
│  Modern page "Work Boards" (single-part app page, full width)                          │
│  ┌─────────────────────────────── SPFx bundle (React 17, Fluent UI 8) ───────────────┐ │
│  │  Web parts: WorkBoards (app) · later: Board (embed) · Dashboard · Form            │ │
│  │                                                                                   │ │
│  │  UI layer ── hash router (#/board/12/item/345) · views · item panel · wizards     │ │
│  │  State ───── useBoard (items, columns, groups) · optimistic updates · polling     │ │
│  │  Services ── BoardService · ItemService · UpdateService · PeopleService · Prefs  │ │
│  │              later: AutomationEngine · InboxService · ViewService                 │ │
│  │  Platform ── Provisioner (migrations) · ETag concurrency · retry on throttling     │ │
│  │              SpClient  →  SPFx SPHttpClient  (signed-in user, no tokens)          │ │
│  └───────────────────────────────────────┬───────────────────────────────────────────┘ │
│                                          │ /_api (REST, as the signed-in user)         │
│  Lists (hidden) ─────────────────────────▼─────────────────────────────────────────────│
│  Phase 1:                                             Later phases:                    │
│  WB_Meta            schema version                    WB_Views         saved views     │
│  WB_Boards          board registry + config JSON      WB_Automations   rules           │
│  WB_Board_<key>     items of one board (×N)           WB_Dashboards    widget layouts  │
│  WB_Board_<key>_Updates  comments of that board (×N)  WB_Forms / WB_FormInbox          │
│  WB_UserPrefs       favourites, recent (own rows)     WB_Templates     library         │
└────────────────────────────────────────────────────────────────────────────────────────┘
        Later, optional, outside the bundle: Power Automate flow templates for email
        digests and date-based automations (standard connectors, no app registration).
```

### 4.1 Why one list per board

We looked at three storage models:

| Option | Pros | Cons |
|--------|------|------|
| A. One central Items list, column values in JSON | Easy cross-board queries, simple setup | Cannot filter or sort on JSON server-side. Private boards would need item-level permissions (slow, limited). The 5,000-item view threshold hits every board together |
| B. **One list per board, real fields per column** (chosen) | Real SharePoint fields: server-side filter and sort, indexes, version history per field, Excel and Power BI connect directly. Board permissions = list permissions. Each board gets its own 5,000 threshold | Cross-board queries need a fan-out (a few boards at a time). Creating a board needs the Manage Lists permission, which site Members have by default |
| C. Per-board list plus a central index list | Fast cross-board queries | Index can drift out of sync, and it leaks titles of private boards |

Option B uses what SharePoint is good at. Version history gives us the
activity log. List permissions give us private boards. Excel, Power BI and
Power Automate can all read the data without anything extra.

### 4.2 Fixed fields on every board list

Every board list gets these fields, added directly to the list when the board
is created. (A shared site content type was considered; per-list fields are
simpler to create over REST and keep each board self-contained.) Cross-board features (My Work, dashboards, automations)
can rely on their internal names.

| Field | Type | Indexed | Purpose |
|-------|------|:------:|---------|
| `Title` | Text | | Item name |
| `WB_GroupId` | Text(32) | ✓ | Group the item belongs to |
| `WB_SortOrder` | Number | ✓ | Order key. Moving an item writes only that item: its key becomes the midpoint of its new neighbours, and a group is renumbered in the rare case the gap gets too small |
| `WB_ParentId` | Number | ✓ | Subitem parent (empty for top-level items) |
| `WB_Owner` | Person (multi) | | The board's Owner column, used by My Work. SharePoint cannot index multi-person fields, so My Work falls back to a paged scan on boards over 5,000 items |
| `WB_Status` | Choice (fill-in allowed) | ✓ | The board's main Status column |
| `WB_StartDate`, `WB_DueDate` | Date | ✓ | The board's main Timeline (start and end) or Date column, used by My Work and the Timeline view |

The readable item ID (for example `MKT-142`) is the board key plus the
SharePoint item ID, so it needs no field. `Created`, `Author`, `Modified` and
`Editor` provide the Creation log and Last updated columns. Dates are written at
12:00 UTC so the calendar day reads the same in every time zone.

Fixed fields are reused: a board's first People, Status and Timeline (or Date)
columns use them, so every template's Owner, Status and Timeline columns work
in My Work.

### 4.3 User-added columns

Each other column becomes a real field named `WB_c_<6-char id>`. Its display
name and colour labels live in the board config.

| Column type | SharePoint field | Stored in board config |
|-------------|------------------|------------------------|
| Status / Priority | Choice with fill-in values allowed, so labels can be added or renamed without changing the field | Label colours, "done" label, order |
| Text / Long text | Text / Note (plain) | |
| Numbers | Number | Unit, decimals |
| Date | Date | |
| Timeline | Two Date fields (`WB_c_x` and `WB_c_xe`) | |
| People | User (multi) | |
| Dropdown / Tags | MultiChoice (fill-in allowed) | Options and colours |
| Checkbox | Boolean | |
| Link | Text (the address) | |
| Email / Phone | Text | Format validation |
| Rating | Number (0–5) | Max |
| Files | Item attachments | |
| Dependency | Lookup (multi) to the same list | Type (finish-to-start) |
| Connect boards | Lookup (multi) to another board list on the same site | Target board |
| Formula | Not stored. Computed in the browser | Expression |
| Progress | Not stored. Computed from status columns | Source columns |
| Time tracking | Number (seconds) + Text (running since) | |

Renaming a column changes the config and the field's display name. Deleting a
`WB_c_` column removes the field after the user confirms. Deleting a column that
uses a fixed field only removes it from the board, and adding that kind of column
again brings its values back.

### 4.4 Board registry (`WB_Boards`)

| Field | Type | Notes |
|-------|------|-------|
| `Title` | Text | Board name |
| `WB_Key` | Text, indexed, unique | Short key used in list names and item IDs, for example `MKT` |
| `WB_ItemsListId`, `WB_UpdatesListId` | Text (list id) | |
| `WB_Description` | Note | |
| `WB_Folder` | Text | Folder in the sidebar |
| `WB_Color` | Text | |
| `WB_Privacy` | Choice: Main, Private | |
| `WB_BoardOwners` | Person (multi) | Who can change columns and settings |
| `WB_Config` | Note (JSON) | Columns, groups, labels, Kanban and Timeline columns, default view. Has a format version (`v`) |
| `WB_Archived` | Yes/No | |

A private board's registry row has its permissions broken the same way as
its lists, so people outside the board cannot see its name either.

### 4.5 The other site lists

Phase 1 creates `WB_Meta` and `WB_UserPrefs`. The others arrive with the
features that need them, through new migrations.

| List | Purpose | Notable settings |
|------|---------|------------------|
| `WB_Meta` | Installed schema version | One row |
| `WB_Views` | Shared saved views: type, filters, sort, hidden columns, grouping, view-specific settings | |
| `WB_UserPrefs` | Favourites and recent boards (later: Inbox "last seen", personal views) | **Item-level security: people read and edit only their own rows** |
| `WB_Automations` | Board, trigger, conditions, actions (JSON), enabled, run count, last error | Only board owners can edit |
| `WB_Dashboards` | Name, boards used, widgets (JSON) | |
| `WB_Forms` | Board, fields, questions, visibility, thank-you text | |
| `WB_FormInbox` | Form submissions (JSON) waiting to be copied onto the board | Submitters can add items and read only their own |
| `WB_Templates` | Document library of board templates (JSON) | |

---

## 5. Provisioning and deployment

### 5.1 Package

- One solution: `work-boards-webpart.sppkg` (shown as **Work Boards** in the App Catalog, with its own icon), SPFx 1.23 (latest GA at time of
  writing), Node 22, React 17, Fluent UI 8. No other runtime libraries.
- `skipFeatureDeployment: false`, so the app is available only on the sites
  where a site owner adds it (**Settings > Add an app**). This matches the
  decision to deploy to chosen sites. It works from the tenant App Catalog or a
  site collection App Catalog.
- `webApiPermissionRequests` is empty, so there is nothing to approve in the
  API access page.
- The web part manifest sets `supportsFullBleed` and `SharePointFullPage` (so
  the app can fill a full-width section or be a full-page app page) and
  `TeamsTab` (tested in phase 3).

### 5.2 First run on a site (setup wizard)

1. A site owner adds the **Work Boards** app to the site and creates a page
   with the Work Boards web part.
2. The web part reads `WB_Meta`. If it is missing:
   - A user with **Manage Web** permission sees the setup wizard.
   - Anyone else sees "Work Boards isn't set up on this site yet. Ask a site
     owner to open this page and set it up."
3. The wizard shows what it will create, then runs the migrations in order,
   with a progress list. Each step is idempotent (it checks before it
   creates), so a failed or interrupted setup can simply run again.
4. The first board is then created from **New board**, from a template, with
   optional example items.

### 5.3 Upgrades

Each release ships a list of numbered migrations (`MIGRATIONS` in
`src/services/Provisioner.ts`; phase 1 has three). When a site owner opens the
app and the version in `WB_Meta` is behind, they see **Update now**. Anyone else
sees a message asking a site owner to open the page. Future board-level
migrations (a new fixed field on every board list) will run the same way, board
by board.

### 5.4 Creating a board

1. The user chooses a template or blank board and a privacy setting.
2. The client creates the lists `WB_Board_<KEY>` and `WB_Board_<KEY>_Updates`
   (hidden, versioning on, attachments on), adds the fixed fields and the
   template's columns (with indexes), and renames the list's Title field to
   "Item" so the list also reads well in SharePoint.
3. It writes the registry row last, so an interrupted create leaves no
   half-built board visible. Creating it again reuses the lists that exist.
4. For a private board it breaks permission inheritance on both lists and the
   registry row, keeps the site's Owners group with Full Control, and grants
   the creator Owner access and the chosen people Member access.
5. Optional example items from the template are added.

Permissions needed: Members can create main boards (Manage Lists is part of
the default Edit level). Only owners can create private boards, because
breaking inheritance needs Manage Permissions. A member can ask an owner to
make a board private.

---

## 6. Permissions

| Role | SharePoint mapping | Can |
|------|-------------------|-----|
| Workspace admin | Site Owners | Everything, including setup and upgrades |
| Board owner | Listed in `WB_BoardOwners`; on a private board, Edit on the board lists | Change columns, groups and board settings |
| Board member | On a main board, site Members (Edit). On a private board, Contribute on the board lists | Add and edit items, post updates. On main boards, members can also change columns because Edit includes Manage Lists |
| Viewer | Read on the board list | See the board. The UI hides all edit controls |
| Form submitter (phase 2) | Add on `WB_FormInbox` only | Submit a form |

SharePoint enforces every rule on the server. The UI only hides what the
user can't do anyway, based on `EffectiveBasePermissions`.

---

## 7. Key technical designs

### 7.1 Fast board loading

- A board loads with one paged `items` query (5,000 rows per page, following
  `odata.nextLink`), selecting only the board's fields and expanding People
  columns. Filtering and sorting then run in the browser, which keeps views
  instant and avoids threshold errors from unindexed filters.
- Phase 2: virtualised rows for boards over a few thousand items, and a cached
  copy of the last board so it appears straight away.

### 7.2 Automations engine (phase 2)

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

### 7.3 Inbox and notifications (phase 2, pull model)

Notifications are not copied into a central list, because that list would
be readable by everyone on the site and would leak private board content.
Instead the Inbox is **computed** for the signed-in user:

- Updates where they are @mentioned (`Mentions` field on each board's
  Updates list).
- Items where they are in a People column and that someone else changed
  since they last looked (`Modified` after "last seen", `Editor` not them).
- Replies to their updates.

Queries go to every board the user can access, a few boards at a time.
SharePoint returns only what they are allowed to see. The automation
"notify" action posts an update that mentions the person, so it lands in the
same Inbox. The last-seen time is kept in `WB_UserPrefs`.

### 7.4 Email (later)

**Deferred** (decision 3 in section 12). When it is added: Microsoft has
announced that `SP.Utilities.Utility.SendEmail` is being retired, and the Graph
`sendMail` API would need permissions we have ruled out. So email will be
optional and use Power Automate:

- A **flow template** (standard SharePoint and Outlook connectors, no premium
  licence) sends a daily or hourly digest of new mentions and assignments.
  It runs under the account of whoever imports it. That is a normal Microsoft
  365 sign-in, not an app registration.
- People can also use SharePoint's own "Alert me" on any board list.

### 7.5 Near-live updates and conflicts

- While a board is open and the tab is visible, the client reads the list's
  `LastItemUserModifiedDate` and `LastItemDeletedDate` every 20 seconds. This is
  one small request. When either changes (and no save of ours is in flight), the
  board reloads quietly.
- Every edit sends the item's ETag and only the fields that changed. If someone
  else saved the item in the meantime (HTTP 412), the client re-sends just its
  own fields, reloads the item, and shows "Someone else changed this item at the
  same time. Both changes were kept."
- Edits show immediately (optimistic update) and roll back with a message if
  the save fails.
- Board settings (columns, groups, labels) are saved the same way: the change
  is re-applied to the latest saved settings on a conflict.

### 7.6 Throttling and scale

- Bulk operations (renaming a label on many items, deleting several items,
  My Work across boards) run with a concurrency limit of four requests.
- 429 and 503 responses are retried with backoff, honouring `Retry-After`,
  the same way the File Type Analyser in this repository does.
- Guidance: hundreds of boards per site and up to about 20,000 items per
  board. Larger programmes should be split across sites (workspaces).

### 7.7 Security

- Updates are stored and shown as plain text (React escapes everything), with
  mentions as `@[Name](userId)` tokens. When formatting arrives in phase 2,
  HTML will be sanitised with DOMPurify before it is saved and shown.
- Formulas use a small expression parser with a fixed set of functions. There
  is no `eval` or `new Function`.
- No secrets, tokens or external calls. The app's own traffic goes only to
  the site's `/_api` (the Fluent UI icon font comes from Microsoft's CDN, as on
  any SharePoint page).

---

## 8. Code layout

```
work-boards/
  config/                      SPFx config (package-solution.json: no API permissions)
  src/
    webparts/workBoards/       The web part: creates SpClient, passes theme colours
    app/
      App.tsx, router.ts       Bootstrap (setup check), hash routes, shared context
      shell/                   Sidebar, Home, My Work, New board dialog, Setup
      board/                   Board page, useBoard state hook, label editor, settings
      views/                   TableView, KanbanView, TimelineView
      cells/                   Cell renderers and editors for every column type
      item/                    Item panel: fields, subitems, updates, files, activity
      common/                  Avatars, people picker, confirm dialog, colour swatches
    services/
      SpClient.ts              SPHttpClient wrapper: retry on 429/503, errors, paging
      Provisioner.ts           Numbered migrations and setup status
      lists.ts                 Idempotent list and field helpers
      BoardService.ts          Registry, board creation, columns, groups, privacy
      ItemService.ts           Items, attachments, version history, My Work
      UpdateService.ts         Updates (comments), likes
      PeopleService.ts         People search (SharePoint people picker API)
      PrefsService.ts          Favourites and recent boards
      permissions.ts           Permission bits and board rights
      test/                    Provisioner tests against a fake SharePoint
    engine/                    Pure logic: fieldMap, viewQuery, ordering, dates,
                               activity, mentions, csv, ids (+ test/)
    models/                    Types, templates, colours
  docs/                        This plan, wireframes, screenshots
  releases/work-boards-webpart.sppkg
```

Runtime libraries: only React 17 and Fluent UI 8, both provided by SPFx. Drag
and drop uses the browser's own drag events (table and Kanban) and pointer
events (Timeline). Later phases may add `chart.js` for dashboards and
`dompurify` for formatted updates, loaded only when needed.

---

## 9. Delivery plan

| Phase | Scope | Status |
|-------|-------|--------|
| **0. Discovery** | Confirm how monday.com is used and the deployment model | Done (section 12) |
| **1. Core** | Setup and migrations; Home; My Work; boards, groups, items, subitems; Status, People, Timeline, Date, Text, Long text, Numbers, Dropdown, Checkbox and Link columns; table, Kanban and Timeline (Gantt) views; search, filters, sort and hide; item panel with updates, @mentions, files and activity; templates; private boards and board settings; CSV export; near-live refresh | **Built.** Next: pilot with one or two teams on a test site |
| **2. Parity** | Automations (event triggers), Inbox with unread count, dashboards and a Dashboard web part, forms and a Form web part, calendar view, shared saved views, dependencies, more column types, duplicate board and save as template, formatted updates, board embed web part, row virtualisation | ~80% coverage. Wider rollout |
| **3. Extras** | Date-based and recurring automations, email digests (Power Automate), formula, time tracking, connect and mirror boards, workload view, Teams tab, Excel import, usage stats | ~90% coverage. Cancel monday.com |

Each phase ends with a tested `.sppkg` in `releases/` and a pilot on a test
site before production.

**Testing:** Jest unit tests for the engine (field mapping, filters, ordering,
dates, activity diff, mentions, CSV), routing, permissions and the migrations
against an in-memory fake of the SharePoint API. The UI is also driven end to end
in a browser (Playwright) against that fake. Before each release, a smoke test on
a real test site covers setup, board creation, editing, private boards and My
Work.

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
| No server means time-based automations are not exact (phase 3) | Lazy scheduler for most cases, optional Power Automate flow where timing matters |
| Members can't create private boards | Owners create them, or a member requests one in the app |
| Very large boards | Indexed fields, paging, virtualisation, and guidance to split boards over ~20,000 items |
| Microsoft changes to SharePoint APIs | Only documented REST APIs are used, all behind one service layer |
| Starting fresh leaves monday.com history behind | Export monday.com boards to Excel before cancelling and keep them as an archive |

---

## 12. Decisions

Answers to the phase 0 questions, and what they changed:

1. **Use:** work management and project management within the department.
   The Timeline column and Gantt view moved from phase 2 into phase 1.
2. **Deployment:** available only on chosen sites. The package uses
   `skipFeatureDeployment: false`, and a site owner adds the app per site.
3. **Email:** can be added later. Email digests moved to phase 3 and are not a
   dependency of anything else.
4. **Migration:** none. Work Boards starts fresh, so the Excel import and the
   monday.com migration tool were dropped from phases 1 and 3 (a general Excel
   import stays in phase 3 as a convenience).
5. **External guests:** not used, so guest access is out of scope.
6. **Name:** Work Boards.
