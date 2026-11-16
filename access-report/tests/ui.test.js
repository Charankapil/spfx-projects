// Drives the real web part UI (built into lib/, bundled by tests/harness) in Chromium on the
// fake tenant. Run: npm run build && node tests/harness/build.js && node tests/ui.test.js
// Needs Playwright with Chromium; it is not a dependency of the project.
require('./register');
const assert = require('assert');
const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright');
const pkg = require('../src/webparts/accessReport/flow/package');
const { parseReport } = require('../src/webparts/accessReport/services/reportCsv');

const PAGE = 'file://' + path.join(__dirname, 'harness', 'index.html');
const SHOTS = process.env.SHOTS_DIR;
const HOME = 'https://contoso.sharepoint.com/sites/it';

async function shot(page, name) {
  if (SHOTS) {
    fs.mkdirSync(SHOTS, { recursive: true });
    await page.screenshot({ path: path.join(SHOTS, name + '.png'), fullPage: true });
  }
}
const tab = (page, name) => page.getByRole('tab', { name }).click();
const noOverflow = async (page, label) => {
  const o = await page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth);
  assert.ok(o <= 1, `${label}: the page scrolls sideways by ${o}px`);
};

async function freshSiteJourney(browser) {
  const context = await browser.newContext({ viewport: { width: 1280, height: 900 }, acceptDownloads: true });
  const page = await context.newPage();
  const consoleErrors = [];
  page.on('console', (m) => m.type() === 'error' && !/Access Report\]|Failed to load resource|fabric-icons|Download the React DevTools/.test(m.text()) && consoleErrors.push(m.text()));
  await page.goto(PAGE + '?s=fresh');

  // 1. Set-up
  await page.getByText('Set up this site for the Access Report').waitFor();
  assert.ok(await page.getByText('Still to create: Access Report Subjects, Access Report Settings, Access Reports.').isVisible());
  await shot(page, '01-setup');
  await page.getByRole('button', { name: 'Set up' }).click();
  await page.getByRole('tab', { name: 'Run report' }).waitFor();
  assert.deepStrictEqual(await page.evaluate(() => Object.keys(window.__tenant.lists).sort()), ['/sites/it/AccessReports', '/sites/it/Lists/AccessReportSettings', '/sites/it/Lists/AccessReportSubjects']);

  // 2. Nothing to run yet
  assert.ok(await page.getByRole('button', { name: 'Run report', exact: true }).isDisabled());
  assert.ok(await page.getByText('Add at least one person or group').isVisible());

  // 3. Who and where
  await tab(page, 'Who and where');
  await page.getByText('Nobody yet.').waitFor();
  await page.getByLabel('Name or e-mail').fill('a');
  await page.getByLabel('Name or e-mail').press('Enter');
  await page.getByText('Ann Example <ann@contoso.com>').waitFor();
  assert.strictEqual(await page.locator('ul li').filter({ hasText: 'IT Members' }).count(), 0, 'SharePoint groups are not offered');
  for (const name of ['Ann Example', 'Marketing', 'Finance Auditors']) {
    await page.locator('li').filter({ hasText: name }).getByRole('button', { name: 'Add' }).click();
    await page.locator('li').filter({ hasText: name }).getByRole('button', { name: 'Added' }).waitFor();
  }
  await page.getByRole('button', { name: 'Add “Everyone except external users”' }).click();
  await page.getByRole('cell', { name: 'Everyone except external users', exact: true }).waitFor();
  assert.ok(await page.getByRole('button', { name: 'Add “Everyone except external users”' }).isDisabled(), 'cannot be added twice');
  // by hand: a bad value is refused, a good one is added
  await page.getByLabel('E-mail address').fill('not an address');
  await page.getByRole('button', { name: 'Add', exact: true }).last().click();
  await page.getByText('Enter the sign-in name (e-mail address) of the person.').waitFor();
  await page.getByLabel('E-mail address').fill("O'Brien_fabrikam.com#EXT#@contoso.onmicrosoft.com");
  await page.getByLabel('Name in the report (optional)').fill('Pat (guest)');
  await page.getByRole('button', { name: 'Add', exact: true }).last().click();
  await page.getByRole('cell', { name: 'Pat (guest)', exact: true }).waitFor();
  assert.ok(await page.getByText('5 of 5 included').isVisible());
  let subjects = await page.evaluate(() => window.__tenant.lists['/sites/it/Lists/AccessReportSubjects'].items);
  assert.deepStrictEqual(subjects.map((s) => s.ARLogin), [
    'i:0#.f|membership|ann@contoso.com',
    'c:0o.c|federateddirectoryclaimprovider|22222222-2222-2222-2222-222222222222',
    'c:0t.c|tenant|11111111-1111-1111-1111-111111111111',
    'c:0-.f|rolemanager|spo-grid-all-users/33333333-3333-3333-3333-333333333333',
    "i:0#.f|membership|o'brien_fabrikam.com#ext#@contoso.onmicrosoft.com"
  ]);
  // switch one off, remove one (asks first)
  await page.getByRole('switch', { name: 'Include Everyone except external users' }).click();
  await page.getByText('4 of 5 included').waitFor();
  await page.getByRole('button', { name: 'Remove Everyone except external users' }).click();
  await page.getByRole('button', { name: 'Keep' }).click();
  assert.ok(await page.getByText('4 of 5 included').isVisible());
  await page.getByRole('button', { name: 'Remove Everyone except external users' }).click();
  await page.getByRole('button', { name: 'Remove', exact: true }).click();
  await page.getByText('4 of 4 included').waitFor();
  // options: a bad folder is refused, a missing one is caught, a good change is saved
  const folder = page.getByLabel('Report folder (server-relative path)');
  await folder.fill('AccessReports');
  await page.getByText('Start with /').waitFor();
  assert.ok(await page.getByRole('button', { name: 'Save' }).isDisabled());
  await folder.fill('/sites/it/Nowhere');
  await page.getByRole('button', { name: 'Save' }).click();
  await page.getByText('/sites/it/Nowhere does not exist.').waitFor();
  await page.getByRole('button', { name: 'Undo changes' }).click();
  await page.getByLabel('Only sites whose title or address contains (optional)').fill('fin');
  await page.getByRole('button', { name: 'Save' }).click();
  await page.getByText('Saved. The next run').waitFor();
  assert.strictEqual(await page.evaluate(() => window.__tenant.settingsItem().ARSiteFilter), 'fin');
  await page.getByLabel('Only sites whose title or address contains (optional)').fill('');
  await page.getByRole('button', { name: 'Save' }).click();
  await page.waitForFunction(() => window.__tenant.settingsItem().ARSiteFilter === '');
  await shot(page, '02-who-and-where');
  await noOverflow(page, 'who and where');

  // 4. Run
  await tab(page, 'Run report');
  await page.getByText('4 included').waitFor();
  await page.getByRole('button', { name: 'Run report', exact: true }).click();
  await page.getByText(/Finished: \d+ rows across 9 site collections \(subsites included\), saved as access-report-/).waitFor({ timeout: 60000 });
  await page.getByText('9 of 9 site collections checked (finished)').waitFor();
  const files = await page.evaluate(() => window.__tenant.files);
  const names = Object.keys(files).sort();
  assert.strictEqual(names.length, 2, names.join(', '));
  assert.ok(/^\/sites\/it\/AccessReports\/access-report-\d{4}-\d\d-\d\d-\d{4}-unread\.csv$/.test(names[0]) && names[1] === names[0].replace('-unread', ''));
  const rows = parseReport(files[names[1]]);
  assert.ok(rows.length > 20 && rows.some((r) => r[0] === 'Pat (guest)' && r[4] === 'Payroll' && r[6] === 'Edit / contribute'));
  assert.ok(await page.getByRole('heading', { name: `${rows.length} rows`, exact: true }).isVisible());
  assert.ok(await page.getByRole('button', { name: /Retry \d unread/ }).isVisible());
  await page.getByPlaceholder('Filter by who, site, library or how').fill('payroll');
  const payroll = rows.filter((r) => r.join(' ').toLowerCase().indexOf('payroll') >= 0).length;
  assert.ok(payroll > 0 && payroll < rows.length);
  await page.getByRole('heading', { name: `${payroll} ${payroll === 1 ? 'row' : 'rows'} match`, exact: true }).waitFor();
  await page.getByPlaceholder('Filter by who, site, library or how').fill('');
  await shot(page, '03-run-finished');
  await noOverflow(page, 'run');
  const dl = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download CSV' }).click();
  const csvPath = await (await dl).path();
  assert.strictEqual(fs.readFileSync(csvPath, 'utf8'), files[names[1]].replace(/^﻿/, '﻿'), 'the download is the saved file');

  // 5. Saved reports and what changed
  const earlier = rows.filter((r) => !(r[0] === 'Pat (guest)' && r[4] === 'Payroll')).map((r) => (r[1].endsWith('/sites/legal') ? r.slice(0, 6).concat(['Read'], r.slice(7)) : r));
  earlier.push(['Ann Example <ann@contoso.com>', HOME + '/../old', 'Old site', 'Site', 'Old site', 'https://contoso.sharepoint.com/sites/old', 'Read', 'Effective permission', '']);
  await page.evaluate(
    ([header, body]) => {
      window.__tenant.files['/sites/it/AccessReports/access-report-2026-10-01-0300.csv'] = '﻿' + [header].concat(body).map((r) => r.map((c) => '"' + c.replace(/"/g, '""') + '"').join(',') + '\r\n').join('');
    },
    [['Who', 'Site URL', 'Site title', 'Object', 'Name', 'URL', 'Access', 'How', 'Own permissions'], earlier]
  );
  await tab(page, 'Saved reports');
  await page.getByRole('link', { name: 'access-report-2026-10-01-0300.csv' }).waitFor();
  assert.strictEqual(await page.locator('tbody tr').count(), 3);
  await page.locator('tbody tr').first().getByRole('button', { name: 'View' }).click();
  await page.getByText(names[1].split('/').pop() + ' ·').waitFor();
  await page.getByRole('combobox', { name: 'Compare with an earlier report' }).click();
  await page.getByRole('option', { name: 'access-report-2026-10-01-0300.csv' }).click();
  await page.getByText('What changed in').waitFor();
  const fact = async (label) => (await page.locator('dl > div').filter({ hasText: label }).locator('dd').innerText()).trim();
  assert.strictEqual(await fact('New access'), '1');
  assert.strictEqual(await fact('Access level changed'), '1');
  assert.strictEqual(await fact('Access removed'), '1');
  assert.ok(await page.getByText('Read → Design / manage').isVisible());
  await shot(page, '04-saved-reports-changes');
  await noOverflow(page, 'saved reports');
  await page.getByRole('button', { name: 'Show the report' }).click();
  await page.getByText(names[1].split('/').pop() + ' ·').waitFor();

  // 6. Monthly schedule: the downloads are exactly the packages built for this site
  await tab(page, 'Monthly schedule');
  await page.getByText('The flow has not run yet.').waitFor();
  for (const [button, build, file] of [
    ['Download the flow package', pkg.legacyPackage, pkg.LEGACY_FILE],
    ['Download as a solution', pkg.solutionPackage, pkg.SOLUTION_FILE]
  ]) {
    const d = page.waitForEvent('download');
    await page.getByRole('button', { name: button }).click();
    const got = await d;
    assert.strictEqual(got.suggestedFilename(), file);
    assert.ok(fs.readFileSync(await got.path()).equals(Buffer.from(build(HOME))), `${file} downloaded from the page equals the package built for ${HOME}`);
  }
  await page.evaluate(() => (window.__tenant.settingsItem().ARLastRun = 'Finished 2026-12-01 03:41 UTC: 9 sites, 4 people and groups, 30 rows, 12 not read. File: access-report-2026-12-01-0341.csv'));
  await page.getByRole('button', { name: 'Refresh' }).click();
  await page.getByText('Finished 2026-12-01 03:41 UTC: 9 sites').waitFor();
  await shot(page, '05-monthly-schedule');
  await noOverflow(page, 'schedule');

  // 7. Narrow screen: nothing spills sideways
  await page.setViewportSize({ width: 390, height: 800 });
  for (const t of ['Run report', 'Who and where', 'Saved reports', 'Monthly schedule']) {
    const visible = await page.getByRole('tab', { name: t }).isVisible();
    if (!visible) {
      await page.getByRole('tab', { name: /more/i }).or(page.locator('button[aria-haspopup="true"]')).first().click();
      await page.getByRole('menuitem', { name: t }).click();
    } else {
      await tab(page, t);
    }
    await page.waitForTimeout(150);
    await noOverflow(page, `${t} at 390px`);
  }
  await shot(page, '06-narrow');

  assert.deepStrictEqual(await page.evaluate(() => window.__errors), [], 'no uncaught errors in the page');
  assert.deepStrictEqual(consoleErrors, [], 'no console errors');
  await context.close();
}

async function stopAndResume(browser) {
  const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
  const page = await context.newPage();
  await page.goto(PAGE + '?s=rich&delay=60');
  await page.getByRole('button', { name: 'Run report', exact: true }).click();
  await page.getByText(/[2-5] of 9 site collections checked/).waitFor({ timeout: 60000 });
  await page.getByRole('button', { name: 'Stop (resume later)' }).click();
  await page.getByText('Stopped. Progress is saved in this browser').waitFor({ timeout: 30000 });
  const stoppedAt = Number(/(\d) of 9 site collections checked \(stopped\)/.exec(await page.locator('body').innerText())[1]);
  assert.ok(stoppedAt >= 2 && stoppedAt < 9, `stopped at ${stoppedAt}`);
  await shot(page, '07-stopped');
  // Closing and reopening the page keeps the progress.
  await page.reload();
  await page.getByText(`${stoppedAt} of 9 site collections checked (stopped)`).waitFor();
  await page.getByText('Unfinished run').waitFor();
  await page.getByRole('button', { name: 'Resume' }).click();
  await page.getByText('9 of 9 site collections checked (finished)').waitFor({ timeout: 60000 });
  await page.getByText(/Finished: \d+ rows across 9 site collections/).waitFor();
  const log = await page.evaluate(() => window.__tenant.log.map((l) => l.url));
  // Each site collection starts by listing its top-level site's subsites: count those.
  const collections = await page.evaluate(() => window.__tenant.sites.filter((x) => !x.parent).map((x) => x.url.toLowerCase()));
  const scanned = new Set(log.filter((u) => /\/_api\/web\/webs\?/.test(u)).map((u) => u.split('/_api/')[0].toLowerCase()).filter((u) => collections.indexOf(u) >= 0));
  assert.strictEqual(scanned.size, 9 - stoppedAt, 'after the reload only the remaining site collections were scanned');
  assert.strictEqual((await page.evaluate(() => Object.keys(window.__tenant.files))).length, 2);
  // Discard clears it for good.
  await page.getByRole('button', { name: 'Discard' }).click();
  await page.reload();
  await page.getByRole('button', { name: 'Run report', exact: true }).waitFor();
  assert.strictEqual(await page.getByText('Last run').count(), 0);
  assert.deepStrictEqual(await page.evaluate(() => window.__errors), []);
  await context.close();
}

(async () => {
  const browser = await chromium.launch();
  let failed = 0;
  for (const [name, fn] of [
    ['a new site: set up, choose who, run, review, compare, download the flow', freshSiteJourney],
    ['a run can be stopped, survives a reload, resumes and is discarded', stopAndResume]
  ]) {
    const t0 = Date.now();
    try {
      await fn(browser);
      console.log(`  ok   ${name} (${Date.now() - t0}ms)`);
    } catch (e) {
      failed++;
      console.log(`  FAIL ${name}\n       ${e && e.stack ? e.stack.split('\n').slice(0, 10).join('\n       ') : e}`);
    }
  }
  await browser.close();
  console.log(`\n${2 - failed}/2 passed`);
  process.exit(failed ? 1 : 0);
})();
