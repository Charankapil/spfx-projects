import assert from 'node:assert/strict';
import { test } from 'node:test';

import { splitByThreshold, totalHistogram } from '../../src/webparts/storagePulse/services/activity';
import { sanitizeIndex } from '../../src/webparts/storagePulse/services/centralIndex';
import { ResultsStore, siteFileStem } from '../../src/webparts/storagePulse/services/ResultsStore';
import { checkSite, findSites } from '../../src/webparts/storagePulse/services/siteDirectory';
import { IScanOptions, StorageScanService } from '../../src/webparts/storagePulse/services/StorageScanService';
// @ts-ignore - plain JS test double
import { createMockSharePoint } from '../mock/mockSharePoint.js';

const ORIGIN = 'https://contoso.sharepoint.com';
const LEGAL = `${ORIGIN}/sites/legal`;
const FINANCE = `${ORIGIN}/sites/finance`;
const BASE: IScanOptions = {
  scope: 'siteCollection',
  includeHidden: false,
  excludeSystemLibraries: false,
  excludedLibraries: [],
  speed: 'fast',
  autoRetryDelayMs: 0,
  profileOverride: { minGapMs: 0 }
};
const INDEX = '/sites/finance/SiteAssets/storage-pulse-central-index.json';

test('a page on one site collection scans another one, exactly, and only that one', async () => {
  const sp = createMockSharePoint({});
  const result = await new StorageScanService(sp.context as never, LEGAL).scan(BASE, () => undefined);
  assert.equal(result.rootUrl, LEGAL);
  assert.equal(result.rootTitle, 'Legal');
  assert.deepEqual(result.libraries.map((l) => l.title).sort(), ['Closed Matters', 'Documents']);
  assert.ok(result.libraries.every((l) => l.url.indexOf(`${LEGAL}/`) === 0));
  const expected = sp.expected({ site: '/sites/legal' });
  const split = splitByThreshold(totalHistogram(result.libraries), 12);
  assert.deepEqual({ files: split.totalFiles, bytes: split.totalBytes }, { files: expected.files, bytes: expected.bytes });
  // Storage used is the scanned site collection's, not the page's.
  assert.equal(result.siteStorageBytes, Math.round(180 * 1024 * 1024 * 1024));
  assert.ok(!sp.log.some((l: string) => l.indexOf('/sites/finance/_api/web/lists') >= 0), 'the page’s own site was scanned');

  // The site map and instant estimates work on the other site collection too.
  const map = await new StorageScanService(sp.context as never, LEGAL).mapSite(BASE, undefined, () => undefined);
  assert.equal(map.libraries.length, 2);
  assert.ok(map.libraries.every((l) => l.unscanned && l.metrics));
});

test('results for other site collections are kept on the central site, one file each, with an overview index', async () => {
  const sp = createMockSharePoint({});
  const legal = await new StorageScanService(sp.context as never, LEGAL).scan(BASE, () => undefined);
  await new ResultsStore(sp.context as never, LEGAL).save(legal);
  const stem = siteFileStem(LEGAL);
  assert.ok(sp.files[`/sites/finance/SiteAssets/${stem}.json`], 'results not saved on the central site');
  assert.ok(!sp.files['/sites/finance/SiteAssets/storage-pulse-site-collection.json'], 'the page’s own results were overwritten');

  const finance = await new StorageScanService(sp.context as never, FINANCE).scan(BASE, () => undefined);
  await new ResultsStore(sp.context as never, FINANCE).save(finance);
  await new ResultsStore(sp.context as never, LEGAL).save(legal);

  const index = await new ResultsStore(sp.context as never).loadIndex();
  assert.deepEqual(index.sites.map((s) => s.title), ['Finance', 'Legal']);
  const line = index.sites.find((s) => s.title === 'Legal')!;
  const exact = splitByThreshold(totalHistogram(legal.libraries), 24);
  assert.equal(splitByThreshold(line.histogram, 24).inactiveBytes, exact.inactiveBytes);
  assert.equal(line.libraries, 2);
  assert.equal(line.librariesRead, 2);
  assert.equal(line.siteStorageBytes, legal.siteStorageBytes);

  // Each site's results load back on their own.
  const loaded = await new ResultsStore(sp.context as never, LEGAL).load('siteCollection');
  assert.equal(loaded!.result.rootUrl, LEGAL);

  // Forgetting a site removes its line and its file, and leaves the others.
  await new ResultsStore(sp.context as never).forgetSite(LEGAL);
  assert.deepEqual((await new ResultsStore(sp.context as never).loadIndex()).sites.map((s) => s.title), ['Finance']);
  assert.equal(sp.files[`/sites/finance/SiteAssets/${stem}.json`], undefined);

  // A normal (single-site) save does not touch the index.
  const before = sp.files[INDEX];
  await new ResultsStore(sp.context as never).save(finance);
  assert.equal(sp.files[INDEX], before);
});

test('file names are stable per site collection and never collide', () => {
  assert.equal(siteFileStem(LEGAL), siteFileStem(`${LEGAL.toUpperCase()}/`));
  assert.notEqual(siteFileStem(LEGAL), siteFileStem(`${ORIGIN}/teams/legal`));
  assert.notEqual(siteFileStem(`${ORIGIN}/sites/a-b`), siteFileStem(`${ORIGIN}/sites/a_b`));
  assert.match(siteFileStem(`${ORIGIN}/sites/Ünïcode site`), /^storage-pulse-sc-[a-z0-9-]+-[0-9a-f]{8}$/);
});

test('the index is checked when read back: other hosts, scripts and junk are dropped', () => {
  const good = {
    url: LEGAL,
    title: 'Legal',
    scanStartedAt: '2026-01-01T00:00:00Z',
    scanCompletedAt: '2026-01-01T01:00:00Z',
    scannedBy: 'Alex',
    histogram: { counts: new Array(121).fill(1), bytes: new Array(121).fill(10) },
    versionBytes: 5,
    libraries: 2,
    librariesRead: 2,
    librariesEstimated: 0,
    failures: 0
  };
  const index = sanitizeIndex(
    {
      sites: [
        good,
        { ...good, url: 'javascript:alert(1)' },
        { ...good, url: 'https://evil.example.com/sites/x' },
        { ...good, histogram: { counts: [1], bytes: [1] } },
        { ...good, scanStartedAt: 'yesterday' },
        null
      ]
    },
    ORIGIN
  );
  assert.deepEqual(index.sites.map((s) => s.url), [LEGAL]);
  assert.deepEqual(sanitizeIndex('nonsense', ORIGIN).sites, []);
});

test('site discovery lists each site collection once, from search, and leaves out OneDrive', async () => {
  const sp = createMockSharePoint({});
  const sites = await findSites(sp.context as never);
  assert.deepEqual(
    sites.map((s) => [s.title, s.url]),
    [
      ['Board', `${ORIGIN}/sites/board`],
      ['Finance', FINANCE],
      ['Legal', LEGAL]
    ]
  );
});

test('a pasted address is resolved to its site collection, and access is checked first', async () => {
  const sp = createMockSharePoint({});
  const ok = await checkSite(sp.context as never, `${LEGAL}/Shared%20Documents/Forms/AllItems.aspx`);
  assert.deepEqual(ok, { ok: true, url: LEGAL, title: 'Legal' });
  assert.deepEqual(await checkSite(sp.context as never, `${ORIGIN}/sites/board`), { ok: false, reason: 'access' });
  assert.deepEqual(await checkSite(sp.context as never, `${ORIGIN}/sites/nowhere`), { ok: false, reason: 'notfound' });
  assert.deepEqual(await checkSite(sp.context as never, 'https://contoso-my.sharepoint.com/personal/alex'), { ok: false, reason: 'host' });
  assert.deepEqual(await checkSite(sp.context as never, 'not a url'), { ok: false, reason: 'invalid' });
  assert.deepEqual(await checkSite(sp.context as never, 'http://contoso.sharepoint.com/sites/legal'), { ok: false, reason: 'invalid' });
});
