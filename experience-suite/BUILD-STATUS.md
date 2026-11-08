# Build and package status

## 1.0.1 (current)

Built with Node.js 22.23 and npm 10.9 using the official SPFx 1.24.1 Heft toolchain:

```bash
npm install
npm run build   # heft test --clean --production && heft package-solution --production
```

| Step | Result |
|---|---|
| `npm install` | OK |
| ESLint (`@microsoft/eslint-config-spfx` React profile) | 0 errors, 0 warnings |
| TypeScript 5.8 | OK |
| Jest | 4 suites, 38 tests, all passing |
| Webpack production bundle | OK (~46 KB) |
| `package-solution --production` | `sharepoint/solution/sharepoint-experience-suite.sppkg` |

A copy of the package is in `releases/sharepoint-experience-suite.sppkg`.

### Test suites

- `helpers.test.ts`: email matching, assignment, status classes, URL allow-list, KQL building, dates.
- `searchMapping.test.ts`: Search row → property mapping.
- `SharePointService.test.ts`: per-list `$select`, ETag mapping and If-Match, 404 vs. other errors, approval guards, 412 handling, people result source, KQL escaping. Uses a fake `spHttpClient`.
- `ExperienceSuite.test.tsx`: renders the whole web part in jsdom against a fake SharePoint and covers MySpace counts, ApprovalFlow filtering and approve, TeamPulse filtering, link sanitising, partial list failure, the PeopleConnect toggle and request submission.

## Not yet done

- Tenant validation (see `TEST-PLAN.md`).
- The PowerShell provisioning script has not been executed in this environment.

See `REVIEW.md` for the issues found in 1.0.0 and the fixes applied.
