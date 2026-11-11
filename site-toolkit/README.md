# Site Toolkit

Fourteen SharePoint Online web parts and a site-wide alert banner, packaged as
one solution (`releases/site-toolkit.sppkg`).

- **No API permissions.** Nothing to approve in the SharePoint admin centre's
  API access page, no Azure AD app registration and no Microsoft Graph. Every
  call goes to SharePoint REST or Search, running as the signed-in user, so
  people only ever see what they can already see.
- **Built not to cause throttling.** All requests on a page go through one
  shared queue (details below).

## The web parts

| Web part | What it does | Who can use it | Requests when opened |
|---|---|---|---|
| **Storage Hog Finder** | Storage per library and folder, split into current files and older versions. Lists the largest files and can trim old versions (moved to the recycle bin). | Site owners (storage figures need manage rights) | 2, then 1 per folder opened |
| **Read & Acknowledge** | People confirm they have read each policy. A new major version asks them again. Owners see who has and who hasn't, with CSV export. | Everyone; report for owners | 2 |
| **Stale Content Radar** | Files nobody has changed for 6 months to 5 years, broken down by type and author. | Everyone | 1, then 1 per 100 more |
| **Recycle Bin Explorer** | Search the recycle bin by name, place or person, see days left, and restore. Owners can switch to the whole site collection. | Everyone (sees own deletions) | 1 |
| **Bulk Metadata Editor** | Spreadsheet-style grid for text, number, choice, yes/no and date columns. Supports pasting from Excel, fill down and undo. Saves up to 50 files per request. | Anyone who can edit the files | 3 |
| **Duplicate File Detector** | Files with the same name and size, by site or library. Can be paused and resumed. | Everyone | 0 until Scan; 1 per 500 files |
| **My Files** | Files checked out to me, and files I edited recently, in this site. | Everyone | 2 (then cached 5 min) |
| **Pulse Poll** | One-question poll with live results after voting. | Everyone | 1 |
| **Template Gallery** | Start a new document from an approved template, saved into a library you choose. | Everyone with edit rights | 2 |
| **Org Chart** | Anyone's managers, direct reports and peers, from Microsoft 365 profiles. Includes search, back, and CSV export. | Everyone | 2 (then cached 30 min) |
| **FAQ** | Searchable questions and answers kept in a list, with category filters. | Everyone | 1 (then cached 15 min) |
| **Who Has Access?** | Who has access to the site and at what level. Flags guests and "everyone". Lists files shared outside the organisation and lists with their own permissions. | Site owners | 1, then 1 per tab |
| **Broken Link Checker** | Checks the internal links on the site's pages and lists the dead ones and where they are. | Everyone | 0 until Check; 1 per page batch + 1 per distinct link |
| **Alert Banner Settings** | Switches the banner on or off and manages its alerts (severity, dates, link). | Site owners; switching needs a site collection admin | 2 |

Plus the **Alert Banner** application customizer, which shows active alerts at
the top of every page in the site collection.

### Lists the web parts create

Several web parts keep their data in a list. The first time a site owner opens
one, it shows a **Set up** button. Everyone else is asked to fetch a site owner.

| Web part | Creates |
|---|---|
| Read & Acknowledge | a `Policies` library (if missing) and the `Policy Acknowledgements` list |
| Pulse Poll | `Poll Votes` |
| Template Gallery | a `Templates` library (if missing) |
| FAQ | `FAQ`, with one example question |
| Alert Banner Settings | `Site Alerts`, on the top-level site of the site collection |

Lists are addressed by a fixed URL (`/Lists/Toolkit...`), so owners can rename
them. Acknowledgement and vote lists only let people edit their own items.

## How it avoids throttling

SharePoint documents a limit of about 3,000 requests per user per 5 minutes.
The toolkit stays far below it:

1. **One queue per page.** Every request from every Site Toolkit web part on
   the page goes through one queue (`src/common/RequestQueue.ts`). The queue
   sends one request at a time and leaves at least 500 ms between them.
2. **A hard budget.** No more than 600 requests in any rolling 5 minutes,
   whatever the web parts ask for. A batch counts as all of the operations in
   it.
3. **Throttling is handled.** On a 429 or 503 response, the whole queue pauses
   for the `Retry-After` time, then repeats the request. The gap is doubled
   for the rest of the session, and the page explains the pause instead of
   showing an error.
4. **Counts come from the search index, not from walking libraries.** Batch
   requests cover bulk edits and org chart profiles. Results are cached in the
   browser, from 1 minute (poll results) to 12 hours (profile photos).
5. **Nothing heavy runs on its own.** Scans and link checks start only when
   someone presses the button. Long scans can be stopped, and nothing polls on
   a timer.
6. **The banner is cheap.** It runs on every page, so it asks SharePoint at
   most once per person every 15 minutes. It also remembers when the alert
   list does not exist.

The Org Chart is the web part most likely to be busy, so it has a few rules
of its own:

- Profiles come from SharePoint's user profile service. Each person's profile
  is fetched once and then cached.
- Everyone around a person is fetched in one batch request (up to 50 people
  per batch).
- Each profile photo is fetched separately through the queue, at most 60 per
  page, and kept in the browser for 12 hours. Photos can be turned off in the
  web part settings.

## Install

1. Upload `releases/site-toolkit.sppkg` to the tenant app catalog, or to a
   site collection app catalog.
2. Choose **Enable app** (or **Deploy**). No API access requests appear,
   because there are none.
3. Add the web parts to pages from the **Site Toolkit** group.
4. For the banner, add **Alert Banner Settings** to a page on the site
   collection. Then switch the banner on with an account that is a site
   collection administrator.

## Build

Needs Node.js 18 (SPFx 1.18.2).

```bash
npm ci
npm run build               # gulp bundle --ship
npm run package-solution    # writes sharepoint/solution/site-toolkit.sppkg
```

The web part classes, manifests and `config/config.json` are generated from
the table in `tools/generate-webparts.js`. Change that table and run
`npm run generate`; don't edit the generated files.

## Tests

```bash
npm test               # 56 unit tests: queue, batching, search, lists, CSV, profiles, ...
npm run test:browser   # 44 browser scenarios against a simulated SharePoint
```

The browser scenarios open each web part in Chromium against an in-memory
SharePoint (`tools/harness/mock-sharepoint.js`) and click through its main
paths. A scenario fails on any console error, two requests in flight at once,
a gap shorter than the queue allows, the wrong number of requests, or list
content running as script. The browser run needs Playwright with Chromium.

**What this does not prove:** the simulated SharePoint answers in the shapes
the documentation describes. It has not been run against a real tenant. Try
the package on a test site collection before rolling it out.

## Limits worth knowing

- Search-based figures (Stale Content, Duplicates, My Files, shared files)
  follow the search index, which can lag by minutes.
- Duplicates are matched on name and size, not content.
- The Broken Link Checker cannot test links to other websites; browsers do not
  allow it. It reads up to 200 pages and checks up to 600 links per run.
- The Org Chart shows people with a Microsoft 365 profile. A change of manager
  in Entra ID can take a day to reach SharePoint's profiles.
