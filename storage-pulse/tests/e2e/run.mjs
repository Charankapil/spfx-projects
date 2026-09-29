// Browser tests: run `gulp bundle` in the project first (they use lib/).
// Uses the Chromium at CHROMIUM_PATH, or Playwright's default install.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import path from 'path';
import { fileURLToPath } from 'url';
import { chromium } from 'playwright-core';
import { buildPage } from './build.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const pageFile = await buildPage();
const browser = await chromium.launch(process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : {});
const shots = process.env.SCREENSHOTS_DIR;

async function open(scenario, viewport = { width: 1240, height: 900 }) {
  const page = await browser.newPage({ viewport });
  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  await page.addInitScript((s) => {
    window.__scenario = s;
    if (s.props && s.props.theme && s.props.theme.isDark) {
      document.addEventListener('DOMContentLoaded', () => document.body.classList.add('dark'));
    }
  }, scenario);
  await page.goto('file://' + pageFile);
  return { page, errors };
}

async function scanToEnd(page) {
  await page.getByRole('button', { name: /Start scan|Run new scan/ }).first().click();
  await page.getByText('Where the inactive storage is').waitFor({ timeout: 120000 });
  await page.waitForTimeout(1500); // count-up animations
}

const savedFileKey = '/sites/finance/SiteAssets/storage-pulse-site-collection.json';

test('owner scan: totals match, results are saved, "saved by" comes from SharePoint', async () => {
  const { page, errors } = await open({ mock: { bigItems: 60000 } });
  await scanToEnd(page);
  const info = await page.evaluate(() => {
    const figs = [...document.querySelectorAll('[class*=figureValue]')].map((e) => e.getAttribute('title'));
    const bytes = figs.slice(0, 2).map((t) => Number(String(t).replace(/[^0-9]/g, '')));
    return { bytes: bytes[0] + bytes[1], expected: window.__sp.expected({}), saved: Object.keys(window.__sp.files) };
  });
  assert.equal(info.bytes, info.expected.bytes);
  assert.ok(info.saved.indexOf('/sites/finance/SiteAssets/storage-pulse-site-collection.json') >= 0);
  if (shots) {
    await page.locator('.ms-MessageBar button').first().click();
    await page.waitForTimeout(300);
    await page.screenshot({ path: path.join(shots, 'dashboard-light.png'), fullPage: true });
  }
  await page.evaluate(() => window.__remount(false));
  await page.getByText(/by Megan Bowen/).waitFor({ timeout: 10000 });
  assert.deepEqual(errors, []);
  await page.close();
});

test('failures are listed with their reason, and Retry failed fixes them', async () => {
  const { page, errors } = await open({ mock: { flaky: true } });
  await scanToEnd(page);
  await page.getByRole('heading', { name: 'Scan issues' }).waitFor();
  assert.ok(await page.getByText(/Flaky Library/).first().isVisible());
  assert.ok(await page.getByText(/server error or timed out/).first().isVisible());
  // The summary no longer blames access.
  assert.equal(await page.getByText(/with your access/).count(), 0);
  if (shots) {
    await page.locator('section[aria-label="Scan issues"]').screenshot({ path: path.join(shots, 'scan-issues.png') });
  }
  await page.evaluate(() => window.__sp.heal());
  await page.getByRole('button', { name: /Retry failed/ }).click();
  await page.getByText(/Retried the parts that failed/).waitFor({ timeout: 60000 });
  // Only the restricted HR site is left (it really is inaccessible).
  assert.ok(await page.getByText(/1 still could not be read/).isVisible());
  assert.equal(await page.locator('section[aria-label="Scan issues"] li').count(), 1);
  assert.deepEqual(errors, []);
  await page.close();
});

test('throttling shows a calm status, never an error, and the scan finishes with exact totals', async () => {
  const { page, errors } = await open({ mock: { tenantLimit: { perSecond: 5, retryAfterSec: 2 }, extraLibraries: 8 } });
  await page.getByRole('button', { name: /Start scan/ }).first().click();
  await page.getByText('SharePoint asked the scan to slow down').waitFor({ timeout: 60000 });
  assert.ok(await page.getByText(/Nothing is lost|request\(s\) at a time|Slowed down/).first().isVisible());
  if (shots) {
    await page.locator('section[aria-label="Scanning…"]').screenshot({ path: path.join(shots, 'scan-throttled.png') });
  }
  await page.getByText('Where the inactive storage is').waitFor({ timeout: 180000 });
  await page.waitForTimeout(1500);
  assert.equal(await page.getByRole('heading', { name: 'Scan issues' }).count(), 1); // only the restricted HR site
  assert.equal(await page.locator('section[aria-label="Scan issues"] li').count(), 1);
  const bytes = await page.evaluate(() => {
    const figs = [...document.querySelectorAll('[class*=figureValue]')].map((e) => e.getAttribute('title'));
    return Number(String(figs[0]).replace(/[^0-9]/g, '')) + Number(String(figs[1]).replace(/[^0-9]/g, ''));
  });
  assert.equal(bytes, await page.evaluate(() => window.__sp.expected({}).bytes));
  assert.deepEqual(errors, []);
  await page.close();
});

test('long throttling pauses with progress saved; a later visit offers Resume and finishes with exact totals', async () => {
  const { page, errors } = await open({
    mock: { outage: { afterRequests: 25, forMs: 6000 }, extraLibraries: 10 },
    props: { tuning: { throttlePatienceMs: 1500, checkpointEveryMs: 0 } }
  });
  await page.getByRole('button', { name: /Start scan/ }).first().click();
  await page.getByText(/SharePoint is throttling requests, so the scan paused/).waitFor({ timeout: 60000 });
  // A pause is a warning with a way forward, not an error.
  assert.equal(await page.getByText(/scan failed/i).count(), 0);
  assert.ok(await page.getByRole('button', { name: 'Resume scan' }).isVisible());
  const saved = await page.evaluate(() => Object.keys(window.__sp.files));
  assert.ok(saved.some((k) => k.endsWith('-inprogress.json')), `checkpoint not saved: ${saved}`);
  if (shots) {
    await page.locator('.ms-MessageBar').first().screenshot({ path: path.join(shots, 'scan-paused.png') });
  }

  // "Closing the tab": the page is rebuilt from what is saved in Site Assets.
  await page.evaluate(() => window.__sp.endOutage());
  await page.evaluate(() => window.__remount(true));
  await page.getByText(/An unfinished scan from/).waitFor({ timeout: 15000 });
  await page.getByRole('button', { name: 'Resume scan' }).click();
  await page.getByText('Where the inactive storage is').waitFor({ timeout: 120000 });
  await page.waitForTimeout(1500);
  const bytes = await page.evaluate(() => {
    const figs = [...document.querySelectorAll('[class*=figureValue]')].map((e) => e.getAttribute('title'));
    return Number(String(figs[0]).replace(/[^0-9]/g, '')) + Number(String(figs[1]).replace(/[^0-9]/g, ''));
  });
  assert.equal(bytes, await page.evaluate(() => window.__sp.expected({}).bytes));
  // The checkpoint is cleaned up once the scan is complete and saved.
  const after = await page.evaluate(() => Object.keys(window.__sp.files));
  assert.ok(!after.some((k) => k.endsWith('-inprogress.json')), `checkpoint left behind: ${after}`);
  assert.ok(after.some((k) => k.endsWith('storage-pulse-site-collection.json')));
  assert.deepEqual(errors, []);
  await page.close();
});

test('an unfinished scan can be discarded, and visitors are not offered it', async () => {
  const { page } = await open({
    mock: { outage: { afterRequests: 25, forMs: 6000 }, extraLibraries: 6 },
    props: { tuning: { throttlePatienceMs: 1500, checkpointEveryMs: 0 } }
  });
  await page.getByRole('button', { name: /Start scan/ }).first().click();
  await page.getByRole('button', { name: 'Resume scan' }).waitFor({ timeout: 60000 });
  await page.evaluate(() => window.__sp.endOutage());
  // A visitor (not an owner) does not see the owner's unfinished scan.
  await page.evaluate(() => window.__remount(false));
  await page.getByText(/Ask a site owner|No scan yet|How much of your storage/).first().waitFor({ timeout: 15000 });
  assert.equal(await page.getByText(/An unfinished scan from/).count(), 0);
  // The owner can discard it.
  await page.evaluate(() => window.__remount(true));
  await page.getByText(/An unfinished scan from/).waitFor({ timeout: 15000 });
  await page.getByRole('button', { name: 'Discard' }).click();
  await page.getByText(/An unfinished scan from/).waitFor({ state: 'detached', timeout: 5000 });
  await page.waitForTimeout(300);
  assert.ok(!(await page.evaluate(() => Object.keys(window.__sp.files))).some((k) => k.endsWith('-inprogress.json')));
  await page.close();
});

test('SharePoint answering 406 does not abandon libraries, and the dashboard says so', async () => {
  const { page, errors } = await open({ mock: { notAcceptable: 'orderby' } });
  await scanToEnd(page);
  assert.ok(await page.getByText(/rejected the fastest query form \(HTTP 406\) for \d+ libraries/).isVisible());
  // Only the truly restricted HR site is listed as an issue, and no library shows a 406.
  assert.equal(await page.getByText(/406/).count(), 1); // just the note above
  const totals = await page.evaluate(() => {
    const figs = [...document.querySelectorAll('[class*=figureValue]')].map((e) => e.getAttribute('title'));
    return Number(String(figs[0]).replace(/[^0-9]/g, '')) + Number(String(figs[1]).replace(/[^0-9]/g, ''));
  });
  assert.equal(totals, await page.evaluate(() => window.__sp.expected({}).bytes));
  assert.deepEqual(errors, []);
  await page.close();
});

test('a quick scan marks libraries measured as a whole and approximate periods', async () => {
  const { page } = await open({ mock: {} });
  await scanToEnd(page);
  assert.ok(await page.getByText(/Quick scan: \d+ libraries/).isVisible());
  assert.ok(await page.getByText(/measured as a whole/).first().isVisible());
  assert.ok(await page.getByText(/Where the site collection storage goes/).isVisible());
  await page.getByRole('combobox').click();
  assert.ok(await page.getByRole('option', { name: '2 years (approx.)' }).isVisible());
  assert.ok(await page.getByRole('option', { name: '6 months', exact: true }).isVisible());
  await page.keyboard.press('Escape');
  await page.close();
});

test('a tampered results file cannot inject script links', async () => {
  const { page: seed } = await open({ mock: {} });
  await scanToEnd(seed);
  const saved = await seed.evaluate((k) => window.__sp.files[k], savedFileKey);
  await seed.close();
  const doc = JSON.parse(saved);
  doc.result.libraries[0].url = 'javascript:alert(document.domain)';
  doc.result.webs[0].url = 'javascript:alert(1)';
  doc.result.largestOldFiles[0].serverRelativeUrl = 'javascript:alert(2)';
  doc.result.scannedBy = 'Forged Name';
  const { page, errors } = await open({ mock: { savedFiles: { [savedFileKey]: JSON.stringify(doc) } } });
  await page.getByText('Where the inactive storage is').waitFor({ timeout: 15000 });
  const bad = await page.evaluate(() =>
    [...document.querySelectorAll('a[href]')].filter((a) => !/^https:\/\/contoso\.sharepoint\.com\//.test(a.getAttribute('href') || '')).length
  );
  assert.equal(bad, 0);
  // Who saved it comes from SharePoint, not from the file.
  assert.ok(await page.getByText(/by Megan Bowen/).isVisible());
  assert.deepEqual(errors, []);
  await page.close();
});

test('a corrupt results file shows a warning instead of breaking the page', async () => {
  const { page, errors } = await open({ mock: { savedFiles: { [savedFileKey]: '{"schemaVersion":2,"result":{"libraries":[{"histogram":1}]}' } } });
  await page.getByText(/could not be loaded/).waitFor({ timeout: 15000 });
  assert.ok(await page.getByRole('button', { name: 'Start scan' }).first().isVisible());
  assert.deepEqual(errors, []);
  await page.close();
});

test('visitors cannot scan when only owners may', async () => {
  const { page } = await open({ mock: { owner: false } });
  await page.getByText(/Ask a site owner/).waitFor({ timeout: 15000 });
  assert.equal(await page.getByRole('button', { name: 'Start scan' }).count(), 0);
  await page.close();
});

test('visitors can scan for themselves when allowed, without saving', async () => {
  const { page } = await open({ mock: { owner: false }, props: { scanPermission: 'everyone' } });
  await scanToEnd(page);
  assert.ok(await page.getByText(/shown to you only/).isVisible());
  assert.equal(await page.evaluate(() => Object.keys(window.__sp.files).length), 0);
  await page.close();
});

test('old results show the stale banner', async () => {
  const { page: seed } = await open({ mock: {} });
  await scanToEnd(seed);
  const saved = JSON.parse(await seed.evaluate((k) => window.__sp.files[k], savedFileKey));
  await seed.close();
  saved.result.scanStartedAt = '2025-01-01T09:00:00.000Z';
  saved.result.scanCompletedAt = '2025-01-01T09:05:00.000Z';
  const { page } = await open({ mock: { savedFiles: { [savedFileKey]: JSON.stringify(saved) } } });
  await page.getByText(/days old/).waitFor({ timeout: 15000 });
  await page.close();
});

test('Cancel stops a scan straight away during throttling', async () => {
  const { page } = await open({ mock: { alwaysThrottle: true } });
  await page.getByRole('button', { name: 'Start scan' }).first().click();
  // Wait until the always-throttled library has been asked for items (the scan is now in its 120 s wait).
  await page.waitForFunction(
    () => window.__sp.log.some((l) => l.indexOf("dddddddd-0000-0000-0000-000000000001')/items?$select=Id,FSObjType") >= 0),
    null,
    { timeout: 30000 }
  );
  await page.waitForTimeout(500);
  const t0 = Date.now();
  await page.getByRole('button', { name: 'Cancel' }).click();
  await page.getByText('Scan cancelled. Any previous results are still shown.').waitFor({ timeout: 5000 });
  assert.ok(Date.now() - t0 < 2000);
  // A cancel is not a failure: no half-finished results or "issues" appear.
  assert.equal(await page.getByRole('heading', { name: 'Scan issues' }).count(), 0);
  if (shots) {
    // (the scan panel itself is captured in the dark-mode test)
  }
  await page.close();
});

test('dark mode and phone width render without horizontal scrolling', async () => {
  const { page, errors } = await open({ mock: { bigItems: 150000, extraLibraries: 20 }, props: { theme: { isDark: true } } });
  await page.getByRole('button', { name: 'Start scan' }).first().click();
  await page.waitForFunction(() => /Items read/i.test(document.body.innerText), null, { timeout: 30000 });
  await page.waitForTimeout(700);
  if (shots) {
    await page.locator('section[aria-label="Scanning…"]').screenshot({ path: path.join(shots, 'scan-panel-dark.png') });
  }
  await page.getByText('Where the inactive storage is').waitFor({ timeout: 120000 });
  await page.waitForTimeout(1500);
  if (shots) {
    await page.locator('.ms-MessageBar button').first().click();
    await page.waitForTimeout(300);
    await page.screenshot({ path: path.join(shots, 'dashboard-dark.png'), fullPage: true });
  }
  await page.setViewportSize({ width: 390, height: 800 });
  await page.waitForTimeout(300);
  assert.ok((await page.evaluate(() => document.documentElement.scrollWidth)) <= 390);
  if (shots) {
    await page.screenshot({ path: path.join(shots, 'phone-dark.png'), fullPage: true });
  }
  assert.deepEqual(errors, []);
  await page.close();
});

test.after(() => browser.close());
