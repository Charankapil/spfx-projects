# Local preview

SPFx 1.18 has no local workbench, so this folder bundles the compiled
web part code with a small in-memory SharePoint site (`fakeSharePoint.js`)
that answers the REST calls Formwright makes. Use it to work on the
designer and the forms without a tenant.

```bash
npx gulp build        # compile src/ to lib/
npm run preview       # bundle lib/ + the fake site to preview/out/
```

Open `preview/out/index.html` in a browser:

| URL                                          | Shows                                                |
| -------------------------------------------- | ---------------------------------------------------- |
| `index.html`                                 | The designer, as "Customize forms" opens it on a "Purchase Requests" list |
| `index.html?sample=1`                        | The designer with the sample purchase-request form   |
| `index.html?view=form&mode=New&sample=1`     | The New form                                         |
| `index.html?view=form&mode=Display&item=1&sample=1` | Item 1 as an approver (Approve button)        |
| `index.html?view=form&mode=Edit&item=1&sample=1`    | Item 1 in edit mode                           |

The fake site lives in memory, so a reload starts from scratch. Every write
is recorded in `window.__fake.log`, which is the easiest way to check the
exact REST payloads a form sends. The signed-in user is a member of
"Finance Approvers", so approval rules in the sample apply.
