# Code review: SharePoint Experience Suite 1.0.0 → 1.0.1

This document lists the issues found when the supplied 1.0.0 source was first built with the official SPFx 1.24.1 toolchain and reviewed line by line, and what was changed. Items under **Remaining concerns** were not fixed in code because they need a design or tenant decision.

## Result

| Check | 1.0.0 as supplied | 1.0.1 |
|---|---|---|
| `npm install` | **Fails** (dependency conflict) | Succeeds |
| ESLint (official SPFx React profile) | **37 errors, 22 warnings** | 0 errors, 0 warnings |
| TypeScript | Compiles | Compiles |
| Jest | **0 tests ran** (file outside `src/`) | 38 tests run, 38 pass |
| `.sppkg` | Not produced | Produced and verified (`releases/`) |

## Blocking build issues

1. **`npm install` failed.** `@rushstack/heft` was pinned to `1.1.4`, but `@microsoft/spfx-heft-plugins@1.24.1` requires `^1.1.5` (ERESOLVE). This is almost certainly why the original "install did not complete". *Fix:* the dependency set now matches the official SPFx 1.24.1 generator (heft 1.3.2, `@types/jest`, `css-loader`, `tslib`, `sp-component-base`, ESLint and npm `overrides`).
2. **The unit test never ran.** `tests/searchMapping.test.ts` was outside `src/`, so it was never compiled into `lib-commonjs/`. Heft reported "No tests found" and still exited with success, which made it look like it passed. *Fix:* tests moved to `src/webparts/experienceSuite/test/`.
3. **Lint failures.** The project had no ESLint config. With the standard SPFx profile, the build fails with 37 errors: components used before they were defined, plus the test file outside the TS program. There were also 22 warnings (`any`, `==`, `void` operator, unused variables). *Fix:* added `eslint.config.js`, moved sub-components into `components/Parts.tsx`, and typed the code properly.

## Functional bugs (would have failed on a real tenant)

4. **Critical: every list read returned HTTP 400 once the lists existed.** All four lists were queried with the same `$select`, which included `RequesterEmail`, `RequesterName` and `ApproverEmail`. The lists don't all have those columns (for example, Announcements has none of them). SharePoint rejects a `$select` naming a column that doesn't exist with a 400, and 400 was not treated as "empty". So the whole suite went to the "Couldn't load your SharePoint profile" error **as soon as the provisioning script had been run**. It only appeared to work while the lists were missing. *Fix:* per-list `$select` in `LIST_FIELDS`, with a unit test.
5. **PeopleConnect always returned no results.** It searched `contentclass:urn:content-class:SPSPeople` against the default result source. The default "Local SharePoint Results" source excludes people. *Fix:* query the **Local People Results** source (`sourceid=b09a7990-…`).
6. **ETag concurrency never worked.** Items were read with `odata=nometadata`, which removes `@odata.etag`, so every update was sent with `IF-MATCH: *`. The "stale ETag" protection described in the docs and test plan was silently off. *Fix:* list reads use `odata.metadata=minimal`, the ETag is kept on each item, and a 412 conflict shows a clear message.
7. **Second task update could be rejected.** Once ETags worked, the task row kept its old ETag after an update. *Fix:* `updateListItem` returns the new ETag from the response and the UI stores it.
8. **One failing list broke the whole app.** `Promise.all` over five calls meant a 403 on, for example, ES Approvals replaced the entire UI with "Couldn't load your SharePoint profile". *Fix:* each list loads independently. Failures show as a named warning and the other modules keep working.
9. **TeamPulse showed everyone's tasks.** The UI and docs say "tasks assigned to you", but all tasks were listed and editable from the UI. *Fix:* filtered by `Owner` = current user's email. The home tile now counts *your open* tasks.
10. **Unassigned approvals could be approved by anyone.** `setApprovalStatus` skipped the assignment check when `ApproverEmail` and `Owner` were both blank. *Fix:* the item must be assigned to the actor.
11. **Missing-list detection depended on the build target.** With the ES5 compile target, `instanceof` on a custom `Error` subclass is unreliable. *Fix:* errors carry a numeric `status`, and the prototype is restored. A unit test caught this.
12. **Office location never showed.** `Office` is not a standard SharePoint Online people managed property. *Fix:* uses `BaseOfficeLocation`, falling back to `OfficeNumber`.

## UI / UX bugs

13. Multi-word statuses (`In Progress`, `Not Started`) produced two CSS classes (`in`, `progress`). `Rejected`, `Submitted` and `Assigned` had no style at all. *Fix:* `statusClass()` produces `in-progress` etc., and the missing styles were added.
14. Turning off PeopleConnect still showed it in the navigation and quick links, and showed a working search box on top of a "disabled" banner. *Fix:* the module is hidden, and an open PeopleConnect view returns to MySpace.
15. Success messages ("Your request was submitted…") followed the user into every other module. *Fix:* cleared on navigation.
16. Every approve or submit action blanked the whole UI with the loading spinner. *Fix:* reloads after an action happen in the background, and stale responses from overlapping reloads are dropped.
17. Pressing Enter in a search box ignored the busy state, so overlapping searches could finish out of order. *Fix:* guarded.
18. "0 results" was shown before any search had run, and "No results" was never shown after one. *Fix:* the results panel distinguishes the two.
19. Request title was limited to 120 characters in the UI while the column allows 255. *Fix:* 255.
20. Due dates were never shown. *Fix:* shown on items that have one.

## Security hardening

21. Result, profile and picture URLs from Search went straight into `href` and `src`. *Fix:* `safeUrl()` only allows `http(s):` and `mailto:` (tested with `javascript:`). Links also use `rel="noopener noreferrer"`.
22. List titles are now URL-encoded as well as quote-escaped in `getbytitle('…')`.

## Platform / maintainability

23. `ReactDom.render` on React 18 runs React in legacy (React 17) mode and was removed in React 19. *Fix:* `createRoot` and `root.unmount()`, as in the official 1.24.1 template.
24. The localized strings file existed but was never used, so the property pane was hard-coded. *Fix:* the property pane uses `ExperienceSuiteWebPartStrings`.
25. `config/serve.json` used the old `serveConfigurations.pageUrl` format. `.gitignore` missed Heft outputs (`lib-commonjs`, `jest-output`, `.heft`, `release`). *Fix:* both replaced with the official template versions.
26. The 140-line JSX file had each module on a single line of up to 1,500 characters. *Fix:* reformatted and split into `ExperienceSuite.tsx` and `Parts.tsx`, with pure helpers in `utils/`.

## Provisioning script

27. **`Connect-PnPOnline -Interactive` no longer works without a client ID.** PnP retired its shared multi-tenant app in September 2024. *Fix:* new `-ClientId` parameter (or `ENTRAID_APP_ID`) and a clear error message. **Note:** the docs say "no app registration is needed". That is true for the web part, but the provisioning script now needs a PnP sign-in app.
28. Fields added with `Add-PnPFieldFromXml` were not added to any view, so list owners couldn't see them. *Fix:* each new field is added to the default view.
29. Status choice columns had no default value. Items created through list forms or flows could have a blank status and would never appear as "Pending". *Fix:* the first choice is the default.
30. A column whose *display* name matched but whose internal name didn't (for example, created by hand as "Approver Email") was silently accepted. *Fix:* the script warns.

> The PowerShell changes were reviewed but **not executed** in this environment (no PowerShell or tenant available). Run them against a test site first.

## Remaining concerns (need a decision, not fixed in code)

- **Client-side checks are not security.** Any user with Edit on ES Approvals can change `Status`, `ApproverEmail` or `DecisionBy` directly through the list UI or REST. Likewise, `RequesterEmail` is written by the browser and can be spoofed. Use item-level permissions, or move decisions into Power Automate or another server-side process. Consider filtering requests by `Author` (Created By, set by the server) rather than `RequesterEmail`.
- **Text email columns instead of Person columns.** `Owner` and `ApproverEmail` are free text. Typos and changed UPNs or aliases silently hide items. Person or Group columns are more robust, but this needs a schema change.
- **Reads are capped at the 100 most recently modified items**, and the user-based filtering happens in the browser. For example, an approver whose item is #101 will not see it. For larger lists, filter on the server (`$filter=ApproverEmail eq '…'`) and index the filtered columns.
- **Theme:** the manifest declares `supportsThemeVariants: true`, but the CSS is fixed light-theme colours. On a dark section the web part stays light. It is readable, but it doesn't follow the theme.
- **Global CSS:** the stylesheet is plain global CSS (`es-` prefixed) rather than CSS modules, so page styles can still clash with it. Several font sizes are 9–10px, which is below accessibility guidance.
- **No telemetry/logging** and no retry/back-off on 429/503 throttling responses.
- **Not yet validated on a tenant.** The build and automated tests are clean, but tenant smoke tests in `TEST-PLAN.md` (especially PeopleConnect fields and permission scenarios) are still required.
