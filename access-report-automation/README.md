# Access report automation

**Every month, unattended: an Excel report of every site, subsite and library a person or group can reach, posted to a SharePoint folder.**

This is the **Access report** from the [SharePoint Admin Center web part](https://github.com/Charankapil/spfx-projects/tree/claude/sharepoint-admin-dashboard-p6lno9/sp-admin-center) (People & access > Access report), rebuilt to run without a browser or a signed-in admin. The web part answers "where does this person have access?" when someone clicks Start and keeps the tab open. This does the same scan on a schedule, for a whole list of people and groups, and adds what a recurring report needs: how each person gets their access, and what changed since last month.

> **Tested against a simulated tenant, not a live one.** The 68 automated tests run the real tool against a mock of Entra ID, Microsoft Graph and SharePoint REST. No Microsoft 365 tenant was available while building it, so the first real run is yours: use the `check` command and a trial run (below) before trusting the schedule. See [What has and has not been tested](#what-has-and-has-not-been-tested).

## What you get

In the folder you choose, each month:

| File | What it is |
| --- | --- |
| `Access report 2026-11.xlsx` | The report. Sheets: **Summary** (per person/group: sites, libraries, highest access, what changed), **Access** (one row per place), **Changes** (added / removed / more / less access since the previous report), **Not read** (sites that could not be read, names not found in the directory) |
| `csv/Access report 2026-11.csv` | The Access sheet as CSV (always complete, also above Excel's row limit) |
| `csv/Access report 2026-11 - changes.csv` | The Changes sheet as CSV (from the second month on) |
| `csv/Access report 2026-11 - not read.csv` | Only when something could not be read |
| `principals.csv` | **You edit this**: who to report on (see below) |
| `_state/` | Run state: a snapshot per month (for next month's comparison) and progress of an unfinished run. Leave it alone |

Columns of the Access sheet:

| Column | Meaning |
| --- | --- |
| Who | The person or group |
| Site URL, Site title | The site collection |
| Object, Name, URL | `Site`, `Subsite`, `Library` (or `List`) and its address |
| Access | `Full control`, `Design`, `Edit`, `Contribute`, `Read` or `View only`, worked out from the actual permission mask (custom levels included) |
| Permission levels | The level names as the site calls them |
| How | The route: `Direct`, `SharePoint group: X`, `Group: Finance -> SharePoint group: X`, `Everyone except external users`, `Site collection administrator`, `Inherits from site` |
| Own permissions | `Yes` when the subsite or library does not inherit |
| Note | Set when a group grants access that SharePoint's own check does not confirm (see [How access is worked out](#how-access-is-worked-out)) |

Folders, files and items with their own permissions (including sharing links) are **not** listed. The report stops at library level, as the web part does.

## How it runs

```
 schedule (GitHub Actions, days 1-5 of the month, 4x a day)
      |
      v
 access-report run ──> this month's report exists? ──yes──> stop (under a minute)
      |no
      v
 sign in as the app (certificate or federated; no user, no password)
      |
      ├─ Microsoft Graph:   who are the people and groups, which groups are they in; list of all sites
      ├─ SharePoint REST:   per site, who has what (2 requests for a site none of them can reach)
      └─ SharePoint REST:   upload workbook + CSVs, save snapshot for next month
```

- **Restartable.** Progress is saved to `_state` every 10 minutes. A run that hits its time limit, loses its connection or is cancelled is continued by the next scheduled run; no site is read twice.
- **Self-retrying.** The schedule fires several times in the first five days. The first run that completes writes the report; later ones see it and stop. A failed run needs nobody: the next one tries again.
- **Gentle.** Two requests at a time by default, `Retry-After` honoured on throttling (and every other request pauses too), identified to SharePoint with a proper `User-Agent`.
- **Quiet logs.** Addresses, e-mail addresses and ids are masked in the log and the run summary, because the Actions log of a public repository is public. `--verbose` shows them.

## Set up

About 20 minutes. You need to be able to grant admin consent in Entra ID (Global Administrator or Privileged Role Administrator).

### 1. Register the app

Follow [`setup/README.md`](setup/README.md): an app registration with four application permissions and a certificate (or a federated credential). You come back with a **tenant id**, a **client id** and a **`.secret.pem`** file.

### 2. Choose the folder

Pick (or create) a document library on a site only the right people can open, for example a `Governance` site. Copy the address of the folder you want the reports in, e.g.

`https://contoso.sharepoint.com/sites/governance/Shared Documents/Access reports`

The library must exist; the folder is created when missing. Use a plain library (no required columns, no forced check-out). For a subsite, set `report.siteUrl` and `report.folder` in the config instead of one address.

The reports say who can open what across the tenant. Treat the folder accordingly.

### 3. Tell GitHub

In the repository: **Settings > Secrets and variables > Actions**.

| Kind | Name | Value |
| --- | --- | --- |
| Variable | `ACCESS_REPORT_TENANT_ID` | Directory (tenant) id. **The schedule stays off until this is set** |
| Variable | `ACCESS_REPORT_CLIENT_ID` | Application (client) id |
| Variable | `ACCESS_REPORT_AUTH` | `certificate` (default) or `federated` |
| Secret | `ACCESS_REPORT_CERTIFICATE` | The whole content of the `.secret.pem` file (not needed for `federated`) |
| Secret | `ACCESS_REPORT_FOLDER_URL` | The folder address from step 2 |
| Secret (optional) | `ACCESS_REPORT_PRINCIPALS` | Who to report on, e.g. `user:alice@contoso.com; group:Finance team`. Usually left out in favour of `principals.csv` |
| Secret (optional) | `ACCESS_REPORT_CONFIG` | More settings as JSON, see [Settings](#settings) |

The workflow file is [`.github/workflows/access-report.yml`](../.github/workflows/access-report.yml). Scheduled workflows only run from the repository's **default branch**, so merge this branch first.

**Use a private repository for the real thing.** Logs are masked, but a private repository is the right home for a job that holds a tenant-wide credential. In a public repository GitHub also switches schedules off after 60 days without commits.

### 4. Check, then say who to report on

1. **Actions > SharePoint access report > Run workflow**, mode **check**. It tests sign-in, each permission, the folder (writes and deletes a test file) and a real permission read, and names the fix for anything that fails.
2. Run it again with mode **run**. With nobody configured yet, it creates an empty `principals.csv` in the folder and stops, telling you so.
3. Open `principals.csv` in the folder (Excel for the web works) and add one row per person or group:

   | Type | Identity | Notes |
   | --- | --- | --- |
   | user | alice@contoso.com | anything you like |
   | user | partner@fabrikam.com | guests by their e-mail |
   | group | Finance team | security or Microsoft 365 group, by name or e-mail |
   | group | 0f3c…-…  | by object id when two groups share a name |

4. Run again. The first report appears in the folder. From then on it is monthly, and `principals.csv` is read fresh each time: add or remove rows whenever you like.

A person who has left (no longer in the directory) is listed on the **Not read** sheet and skipped; the run does not fail.

## Settings

Environment variables win over the config file. All of `scan` and `sites` is optional.

| Setting | Default | Meaning |
| --- | --- | --- |
| `report.folderUrl` / `ACCESS_REPORT_FOLDER_URL` | (required) | Folder for the reports. A copied browser address of the folder view also works |
| `report.siteUrl` + `report.folder` | | The same in two parts (needed for a subsite): site address, and `Library/Folder/...` |
| `report.name` | `Access report` | Start of the file names |
| `principals` / `ACCESS_REPORT_PRINCIPALS` | | People and groups, in addition to `principals.csv` |
| `principalsFile` | `principals.csv` | Name of the list in the report folder; `""` to not use one |
| `sites.only` | | Scan exactly these addresses (no Graph site listing needed) |
| `sites.include`, `sites.exclude` | | Wildcard (`*`) patterns on the site address |
| `sites.includeOneDrive` | `false` | Also scan OneDrive personal sites |
| `scan.includeLibraries` | `true` | List libraries |
| `scan.includeLists` | `false` | List lists too |
| `scan.includeSubsites` | `true` | Walk subsites |
| `scan.expandGroups` | `true` | Read group memberships from the directory, so the report can name the group that gives access |
| `scan.thorough` | `true` | Ask SharePoint for every person's effective permission on every site. `false` skips sites where nothing links the person to the site (see below): about one request per person per site fewer |
| `scan.concurrency` / `ACCESS_REPORT_CONCURRENCY` | `2` | Requests in flight (1 to 8) |
| `scan.maxMinutes` / `ACCESS_REPORT_MAX_MINUTES` | `300` | Stop cleanly after this long; the next run continues. `0` = no limit |
| `auth.mode` / `ACCESS_REPORT_AUTH` | `certificate` | or `federated` |
| `auth.certificatePath`, `auth.privateKeyPath` | | PEM files, for local runs instead of `ACCESS_REPORT_CERTIFICATE` |
| `endpoints.login`, `endpoints.graph` | worldwide cloud | For national clouds, e.g. `https://login.microsoftonline.us` and `https://graph.microsoft.us` |

See [`config.example.json`](config.example.json).

## How access is worked out

For each site collection, things that do not depend on the person are read once and shared by everyone in the list.

1. **Is anyone on the list known to this site?** One request reads the site's entries for groups and the "everyone" claims; one more (per 10 people) looks up the people themselves. The directory says which groups each person is in (nested groups and owned Microsoft 365 groups included), so a group on the site can be matched to its members.
2. **For those who are**, the site's permissions are read once, plus the membership of the matched entries in SharePoint groups. That gives the route (the **How** column) and the permission levels.
3. **For people**, SharePoint's own effective-permission check gives the level (the **Access** column). It accounts for everything SharePoint knows, including denials.
4. Libraries and subsites that inherit get the parent's answer. Those with their own permissions are read the same way.

When the two sources disagree:

- SharePoint says access, but no route was found: the row says `Effective permission (the route was not identified)`.
- A **directory group** grants access, but SharePoint's check says none: the row is reported with the level the permissions grant and a **Note** that it is not confirmed. This is typical for someone added to a group who has not signed in since. Routes through "everyone" claims or direct grants are not second-guessed: if SharePoint says no, there is no row.

**`scan.thorough: false`** skips step 3 on sites where step 1 found no link at all. It is safe unless access comes from something other than a direct grant, a directory group or an "everyone" claim (for instance a directory role claim added to a site), which is rare.

**Cost.** A site nobody on the list can reach costs 2 requests plus, with `thorough`, one per person. A site someone can reach costs roughly 5 more, plus 1 to 2 per library or subsite with its own permissions. As a guide: 10,000 sites and 10 people is about 120,000 requests thorough (a few hours at the default pace, spread over two scheduled runs if needed) or about 25,000 with `thorough: false`.

## Permissions, and why

| API | Permission (application) | Why |
| --- | --- | --- |
| SharePoint | `Sites.FullControl.All` | Reading who has access to a site requires it; SharePoint has no read-only permission that can read permissions. Also writes the report |
| Microsoft Graph | `Sites.Read.All` | List every site collection (not needed with `sites.only`) |
| Microsoft Graph | `User.Read.All` | Look up the people |
| Microsoft Graph | `GroupMember.Read.All` | Look up groups and memberships (not needed with `scan.expandGroups: false` and no groups in the list) |

`Sites.FullControl.All` is a powerful permission. What limits the exposure:

- The tool only **reads** sites. Its only writes are files in the report folder (verified in the tests: no other write requests are made).
- No client secret exists: sign-in is by certificate, or by federated credential with nothing stored at all.
- Redirects are never followed and paging links are only followed on the same site, so the token cannot be sent elsewhere.
- Restrict who can edit the workflow and the repository's secrets. Consider a GitHub environment with required reviewers for manual runs.

To avoid the tenant-wide permission, grant the app `Sites.Selected` with Full Control on chosen sites and list them in `sites.only`.

## Running it somewhere else

It is a plain Node.js (20+) program with no runtime dependencies, so any scheduler works: Azure Automation, an Azure Function timer, Azure DevOps, cron, Task Scheduler.

```bash
npm ci && npm run build

export ACCESS_REPORT_TENANT_ID=...  ACCESS_REPORT_CLIENT_ID=...
export ACCESS_REPORT_FOLDER_URL="https://contoso.sharepoint.com/sites/governance/Shared Documents/Access reports"
export ACCESS_REPORT_CERTIFICATE="$(cat access-report.secret.pem)"

node dist/cli.js check                       # test the setup
node dist/cli.js run --no-upload --limit 20  # trial: first 20 sites, files in ./out, nothing written to SharePoint
node dist/cli.js run                         # the real thing
```

Schedule `run` a few times in the first days of the month and treat exit code `75` as "not finished, will continue": `0` done or nothing to do, `1` failed, `2` settings problem, `75` time limit reached with progress saved. Other options: `--period 2026-11`, `--force` (make it again), `--verbose`, `--json`, `--config file.json`.

## Troubleshooting

| What you see | What to do |
| --- | --- |
| `Sign-in failed (AADSTS700027)` | The certificate is not on the app registration, or has expired. Upload the `.cer` again |
| `Sign-in failed (AADSTS70021 / AADSTS700213)` | Federated credential does not match: check repository, branch or environment, and audience |
| `None of the N sites could be read` | `Sites.FullControl.All` (SharePoint, application) missing or not consented. Run `check` |
| `The list of sites could not be read from Microsoft Graph` | `Sites.Read.All` (Graph, application) missing or not consented |
| `The document library ... does not exist` | The folder address must start with the library's address name (`Shared Documents`, not `Documents`) |
| Many rows with a "not confirmed" note | People who were added to a group but have not signed in since; or groups SharePoint has not synchronised yet. The grant exists |
| Sites on the **Not read** sheet | Locked, archived or deleted sites. They are left out of the changes |
| "The sign-in certificate expires in N days" on the Summary sheet | Make a new certificate ([setup](setup/README.md)) and replace the secret |
| The run keeps ending with "Time limit reached" | Normal on large tenants: it finishes over several runs. To speed up: `scan.thorough: false`, raise `scan.concurrency` (watch for throttling in the log), or narrow `sites` |

## What has and has not been tested

Run `npm test` (needs `openssl`; uses `openpyxl` and LibreOffice when installed to open the workbook).

**Tested** (68 tests, against the mock tenant in `test/mock-tenant.mjs`, which has its own independent permission model):

- The report equals the model for every person and group: direct grants, SharePoint groups, nested security groups, Microsoft 365 members and owners, Everyone / Everyone except external users, guests, site collection administrators, Limited Access, custom levels, subsites and libraries with and without their own permissions. On a hand-built tenant and on 25 random tenants, in three modes each.
- Month-on-month changes; resume after a time limit; resume after an outage; progress from other settings not reused; a second run in the month does nothing.
- Throttling (429 / 503 with `Retry-After`), dropped connections, 500s, a token expiring mid-run.
- Certificate sign-in (the mock verifies the signature), federated sign-in, wrong certificate / tenant / permissions with the right hint.
- The workbook opens in `openpyxl` and LibreOffice with the same content; text that looks like a formula stays text; large files upload in pieces.
- Log masking; exit codes; the `check` command.
- 16 deliberate bugs were injected into the code one at a time; the tests caught each.
- 3,000 sites, 25 people and groups, 93,000 rows: under 400 MB of memory.

**Not tested:** anything against a real Microsoft 365 tenant. The requests follow the documented REST and Graph shapes and reuse the calls the web part makes, but the mock was written from the same understanding as the tool, so a wrong assumption about the real service would pass the tests. Specific things to watch on the first real run:

- `siteusers?$filter=PrincipalType eq 4` (a fallback to the unfiltered list exists if it is refused).
- Whether `getusereffectivepermissions` evaluates accounts that never visited a site (if not, the tool reports the group routes it finds and notes them as unconfirmed, and adds a warning to the Summary sheet).
- Federated sign-in to SharePoint (certificate sign-in is the long-established route).
- Graph `getAllSites` returning the site collections you expect.

`check` and `run --no-upload --limit 20` exercise all of these safely.

## Layout

```
src/
  cli.ts          commands and exit codes
  config.ts       settings, validation
  auth.ts         certificate / federated sign-in (client assertion), token cache
  http.ts         retries, throttling, concurrency ceiling
  graph.ts        sites, people, groups, memberships
  sharepoint.ts   REST reads with paging; file upload / download
  permissions.ts  permission masks, claims, role assignment matching (pure)
  scanner.ts      one site collection: who can reach what, and how
  runner.ts       the run: resume, time limit, retries, report
  report.ts       workbook, CSVs, snapshot, changes
  xlsx.ts, csv.ts writers (no dependencies)
  store.ts        report folder in SharePoint, or a local folder for trial runs
  check.ts        the check command
test/             mock tenant, end-to-end, unit and random-tenant tests
setup/            app registration guide, certificate script
```
