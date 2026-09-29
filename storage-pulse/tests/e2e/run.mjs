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
  await page.waitForFunction(() => /Busy Library/.test(document.body.innerText), null, { timeout: 30000 });
  await page.waitForTimeout(500);
  const t0 = Date.now();
  await page.getByRole('button', { name: 'Cancel' }).click();
  await page.getByText('Scan cancelled').waitFor({ timeout: 5000 });
  assert.ok(Date.now() - t0 < 2000);
  if (shots) {
    // (the scan panel itself is captured in the dark-mode test)
  }
  await page.close();
});

test('dark mode and phone width render without horizontal scrolling', async () => {
  const { page, errors } = await open({ mock: { bigItems: 60000 }, props: { theme: { isDark: true } } });
  await page.getByRole('button', { name: 'Start scan' }).first().click();
  await page.waitForFunction(() => /Items read/i.test(document.body.innerText), null, { timeout: 30000 });
  await page.waitForTimeout(2500);
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
