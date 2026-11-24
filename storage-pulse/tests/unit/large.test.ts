import assert from 'node:assert/strict';
import { test } from 'node:test';

import { IScanResult } from '../../src/webparts/storagePulse/models/IScanResult';
import { AGE_BANDS, bandTotals, splitByThreshold, totalHistogram } from '../../src/webparts/storagePulse/services/activity';
import { folderRows, inactiveFolders } from '../../src/webparts/storagePulse/services/folders';
import { sanitizeResult } from '../../src/webparts/storagePulse/services/safeData';
import { forDisplay, libraryKey } from '../../src/webparts/storagePulse/services/siteMap';
import { IScanOptions, ScanPausedError, StorageScanService } from '../../src/webparts/storagePulse/services/StorageScanService';
// @ts-ignore - plain JS test double
import { createMockSharePoint } from '../mock/mockSharePoint.js';

const ORIGIN = 'https://contoso.sharepoint.com';
const BASE: IScanOptions = {
  scope: 'siteCollection',
  includeHidden: false,
  excludeSystemLibraries: false,
  excludedLibraries: [],
  speed: 'fast',
  autoRetryDelayMs: 0,
  profileOverride: { minGapMs: 0 }
};
const BIG_ID = 'cccccccc-0000-0000-0000-000000000001';

function totals(result: IScanResult): { files: number; bytes: number } {
  const split = splitByThreshold(totalHistogram(result.libraries), 12);
  return { files: split.totalFiles, bytes: split.totalBytes };
}

function bigReads(log: string[]): string[] {
  return log.filter((l) => l.indexOf(`${BIG_ID}')/items?`) >= 0 && l.indexOf('$filter=Id gt') >= 0);
}

test('a large library paused part-way carries on from where it stopped, with exact totals', async () => {
  // Only the large library, so the pause lands inside it.
  const only = { ...BASE, excludedLibraries: ['Documents', 'Site Assets', 'Audit Archive 2016', 'Empty Library', 'Site Pages'] };
  // Find where reading the library starts (after listing, metrics and the folder breakdown), then pause a few pages in.
  const probe = createMockSharePoint({ bigItems: 130000 });
  await new StorageScanService(probe.context as never).scan(only, () => undefined);
  const firstPage = probe.log.findIndex((l: string) => bigReads([l]).length > 0);
  assert.ok(firstPage > 0);
  const sp = createMockSharePoint({ bigItems: 130000, outage: { afterRequests: firstPage + 8, forMs: 60000 } });
  const checkpoints: IScanResult[] = [];
  let paused: ScanPausedError | undefined;
  try {
    await new StorageScanService(sp.context as never).scan({ ...only, throttlePatienceMs: 500, checkpointEveryMs: 0 }, () => undefined, (s) => {
      checkpoints.push(s);
    });
    assert.fail('the scan should have paused');
  } catch (err) {
    assert.ok(err instanceof ScanPausedError, String(err));
    paused = err as ScanPausedError;
  }
  // Checkpoints are taken inside the library, not only between libraries.
  assert.ok(
    checkpoints.some((c) => c.libraries.some((l) => l.id === BIG_ID && l.pending && l.readState && l.readState.seen > 0)),
    'no checkpoint was taken while the large library was being read'
  );
  const snapshot = sanitizeResult(JSON.parse(JSON.stringify(paused!.snapshot)), ORIGIN)!;
  const half = snapshot.libraries.find((l) => l.id === BIG_ID)!;
  assert.ok(half.pending && half.readState, 'the half-read library lost its read state');
  assert.ok(half.files > 0 && half.readState!.seen > 0);
  const progressed = half.readState!.ranges.filter((r) => r.cursor > r.from);
  assert.ok(progressed.length > 0);

  sp.endOutage();
  const before = sp.log.length;
  const resumed = await new StorageScanService(sp.context as never).resume(snapshot, only, () => undefined);
  const after = bigReads(sp.log.slice(before));
  // Every range that had made progress continued from its cursor, not from its start.
  for (const r of progressed) {
    const start = `$filter=Id gt ${r.from}${r.to !== undefined ? ` and Id le ${r.to}` : ''}&`;
    assert.ok(!after.some((l) => l.indexOf(start) >= 0), `range ${r.from}-${r.to} was read again from its start`);
  }
  const expected = sp.expected(only);
  assert.deepEqual(totals(resumed), { files: expected.files, bytes: expected.bytes });
  const big = resumed.libraries.find((l) => l.id === BIG_ID)!;
  assert.equal(big.readState, undefined);
  assert.equal(big.error, undefined);
  assert.equal(big.unreadItems, expected.hidden);
  assert.equal(
    big.fileTypes!.reduce((s, t) => s + t.counts.reduce((a, b) => a + b, 0), 0),
    big.files,
    'file types out of step with the files counted'
  );
});

test('a batch that keeps failing is read again on its own, not with the whole library', async () => {
  const sp = createMockSharePoint({ bigItems: 130000, failAfter: { libraryId: BIG_ID, requests: 6, failures: 5 } });
  const result = await new StorageScanService(sp.context as never).scan(BASE, () => undefined);
  const big = result.libraries.find((l) => l.id === BIG_ID)!;
  assert.equal(big.error, undefined, big.error);
  assert.equal(big.readState, undefined);
  const expected = sp.expected(BASE);
  assert.deepEqual(totals(result), { files: expected.files, bytes: expected.bytes });
  // Each range started once: the retry asked only for the skipped batch.
  const reads = bigReads(sp.log);
  assert.equal(reads.filter((l) => l.indexOf('$filter=Id gt 0 and Id le 50000&') >= 0).length, 1);
  assert.ok(reads.length < 60, `${reads.length} page requests for a 26-page library`);
});

test('large libraries get a folder breakdown from storage metrics: exact, biggest first, Forms left out', async () => {
  const sp = createMockSharePoint({ bigItems: 60000 });
  const result = await new StorageScanService(sp.context as never).scan(BASE, () => undefined);
  const big = result.libraries.find((l) => l.id === BIG_ID)!;
  const years = big.folders!.filter((f) => f.depth === 1);
  assert.equal(years.length, 8);
  assert.ok(!big.folders!.some((f) => f.path === 'Forms'));
  // Exact per folder (SharePoint's figures count every file, including ones hidden from this user).
  const files = sp.libs.big.fileCache as { dir: string; size: number }[];
  for (const y of years) {
    const inside = files.filter((f) => f.dir === y.path || f.dir.indexOf(y.path + '/') === 0);
    assert.equal(y.files, inside.length, y.path);
    assert.equal(y.bytes, inside.reduce((s, f) => s + f.size, 0), y.path);
  }
  // Big folders are opened: their boxes are listed too.
  assert.ok(big.folders!.some((f) => f.depth === 2 && /^Year \d+\/Box \d+$/.test(f.path)));
  assert.deepEqual(
    big.folders!.map((f) => f.totalSize),
    big.folders!.map((f) => f.totalSize).sort((a, b) => b - a)
  );

  // A library with every file at the top level says so.
  const flat = result.libraries.find((l) => l.title === 'Documents' && l.webTitle === 'Projects')!;
  assert.deepEqual(flat.folders, []);
  assert.match(flat.foldersNote || '', /No folders/);

  // Whole folders nobody has touched are reported as inactive, exactly.
  const rows = folderRows(result, 24);
  const audit = result.libraries.find((l) => l.title === 'Audit Archive 2016')!;
  const idle = inactiveFolders(rows).folders.filter((r) => r.libraryTitle === 'Audit Archive 2016');
  assert.equal(idle.length, 3);
  assert.equal(idle.reduce((s, r) => s + r.files, 0), audit.files);
  assert.equal(idle.reduce((s, r) => s + r.bytes, 0), audit.bytes);
  // Nested inactive folders are not counted twice.
  const nested = inactiveFolders(rows.concat(rows.filter((r) => r.libraryTitle === 'Audit Archive 2016').map((r) => ({ ...r, key: r.key + '/x', path: r.path + '/x', depth: 2 }))));
  assert.equal(nested.folders.filter((r) => r.libraryTitle === 'Audit Archive 2016').length, 3);

  // Small libraries do not spend requests on folders.
  assert.equal(result.libraries.find((l) => l.title === 'Site Assets')!.folders, undefined);
});

test('folder breakdown works where SharePoint will not expand metrics, or will not list a big folder', async () => {
  const plain = await new StorageScanService(createMockSharePoint({ bigItems: 60000 }).context as never).scan(BASE, () => undefined);
  const sp = createMockSharePoint({ bigItems: 60000, noFolderExpand: true });
  const fallback = await new StorageScanService(sp.context as never).scan(BASE, () => undefined);
  const a = plain.libraries.find((l) => l.id === BIG_ID)!.folders!;
  const b = fallback.libraries.find((l) => l.id === BIG_ID)!.folders!;
  // Each folder is asked for its own figures instead (within the request budget).
  assert.ok(sp.log.some((l: string) => /GetFolderByServerRelativePath\(decodedurl='.+'\)\/StorageMetrics/.test(l)));
  for (const f of b) {
    const same = a.find((x) => x.path === f.path)!;
    assert.ok(same, f.path);
    assert.equal(f.bytes, same.bytes);
    assert.equal(f.files, same.files);
  }

  const limited = await new StorageScanService(createMockSharePoint({ bigItems: 60000, folderListLimit: 1000 }).context as never).scan(
    BASE,
    () => undefined
  );
  const big = limited.libraries.find((l) => l.id === BIG_ID)!;
  const years = big.folders!.filter((f) => f.depth === 1);
  assert.equal(years.length, 8);
  assert.ok(years.every((y) => y.childrenUnavailable), 'year folders too big to list should be marked');
  assert.ok(!big.folders!.some((f) => f.depth === 2));
  assert.equal(big.error, undefined, 'a folder that cannot be listed must not fail the library');
});

async function mapAndEstimate(mockOptions: Record<string, unknown>, titles: string[]) {
  const sp = createMockSharePoint(mockOptions);
  const service = new StorageScanService(sp.context as never);
  const map = await service.mapSite(BASE, undefined, () => undefined);
  const keys = map.libraries.filter((l) => titles.indexOf(l.title) >= 0).map(libraryKey);
  const before = sp.log.length;
  const estimated = await service.estimateSelected(map, keys, BASE, () => undefined);
  return { sp, map, estimated, requests: sp.log.slice(before) };
}

test('search estimates: counts by age and type match a full read; sizes are scaled to storage metrics', async () => {
  const { sp, estimated, requests } = await mapAndEstimate({ bigItems: 130000 }, ['Scanned Records', 'Documents']);
  const exact = await new StorageScanService(sp.context as never).scan(BASE, () => undefined);

  for (const lib of estimated.libraries.filter((l) => l.estimate)) {
    const e = lib.estimate!;
    assert.equal(e.error, undefined, e.error);
    const truth = exact.libraries.find((l) => libraryKey(l) === libraryKey(lib))!;
    assert.equal(e.files, truth.files, `${lib.title}: files`);
    // Every age band holds exactly the files a full read finds there.
    for (const band of AGE_BANDS) {
      assert.equal(bandTotals(e.histogram, band).files, bandTotals(truth.histogram, band).files, `${lib.title} ${band.label}`);
    }
    // File types: counts are search's own.
    for (const t of truth.fileTypes!) {
      const est = e.fileTypes.find((x) => x.extension === t.extension);
      assert.ok(est, `${lib.title}: missing type ${t.extension}`);
      assert.equal(est!.counts.reduce((a, b) => a + b, 0), t.counts.reduce((a, b) => a + b, 0), `${lib.title} ${t.extension}`);
    }
    // Bytes: the library's average file size from storage metrics, applied to the files search found.
    const m = lib.metrics!;
    assert.ok(Math.abs(e.bytes - (m.fileStreamSize / m.fileCount!) * e.files) <= 1, `${lib.title}: calibrated bytes`);
    // Close to the truth for the active / inactive split.
    for (const threshold of [6, 12, 24]) {
      const share = (h: IScanResult['libraries'][0]['histogram']): number => {
        const s = splitByThreshold(h, threshold);
        return s.inactiveBytes / s.totalBytes;
      };
      const gap = Math.abs(share(e.histogram) - share(truth.histogram));
      assert.ok(gap < 0.05, `${lib.title} ${threshold} months: inactive share off by ${(gap * 100).toFixed(1)} points`);
    }
    // Largest old files: exact sizes from search, at least three months old.
    assert.ok(e.largest.length > 0);
    assert.ok(e.largest.every((f) => f.ageMonths >= 3 && f.serverRelativeUrl.indexOf(decodeURIComponent(new URL(lib.url).pathname) + '/') === 0));
    assert.ok(Math.abs(e.indexCoverage! - truth.files / m.fileCount!) < 1e-9);
  }
  // A few dozen requests for 140,000 files.
  const searches = requests.filter((l) => l.indexOf('/_api/search/query') >= 0);
  assert.ok(searches.length <= 3 * 25, `${searches.length} search requests`);
  assert.ok(!requests.some((l) => l.indexOf("/items?") >= 0), 'an estimate must not read items');

  // The dashboard shows estimated libraries; others not read stay out.
  const view = forDisplay(estimated);
  assert.deepEqual(view.libraries.map((l) => l.title).sort(), ['Documents', 'Documents', 'Scanned Records']);
  assert.ok(view.libraries.every((l) => l.fromEstimate));
  assert.equal(view.siteStorageBytes, undefined);
  assert.ok(view.largestOldFiles.length > 0);
  // Saved and loaded back: the estimate and folders survive; the display flag is never saved.
  const restored = sanitizeResult(JSON.parse(JSON.stringify(estimated)), ORIGIN)!;
  const big = restored.libraries.find((l) => l.id === BIG_ID)!;
  assert.ok(big.estimate && big.estimate.files > 0 && big.folders && big.folders.length > 0);
  assert.deepEqual(forDisplay(restored).libraries.map((l) => l.files), view.libraries.map((l) => l.files));
});

test('search estimates cope with verbose answers, a ListId-blind index, partial indexing and libraries left out of search', async () => {
  const plain = (await mapAndEstimate({}, ['Audit Archive 2016'])).estimated.libraries.find((l) => l.estimate)!.estimate!;
  const verbose = (await mapAndEstimate({ searchShape: 'verbose' }, ['Audit Archive 2016'])).estimated.libraries.find((l) => l.estimate)!.estimate!;
  assert.deepEqual(verbose.histogram, plain.histogram);
  assert.deepEqual(verbose.fileTypes, plain.fileTypes);

  const strict = (await mapAndEstimate({ notAcceptable: 'search' }, ['Audit Archive 2016'])).estimated.libraries.find((l) => l.estimate)!.estimate!;
  assert.equal(strict.error, undefined, strict.error);
  assert.deepEqual(strict.histogram, plain.histogram);

  const byPath = await mapAndEstimate({ listIdUnsupported: true }, ['Audit Archive 2016']);
  const e = byPath.estimated.libraries.find((l) => l.estimate)!.estimate!;
  assert.equal(e.error, undefined);
  assert.equal(e.files, plain.files);
  assert.ok(byPath.requests.some((l: string) => l.indexOf('path:"https://contoso.sharepoint.com/sites/finance/Audit Archive 2016"') >= 0));

  const partial = (await mapAndEstimate({ searchIndexed: 0.8 }, ['Audit Archive 2016'])).estimated.libraries.find((l) => l.estimate)!.estimate!;
  assert.ok(partial.indexCoverage! > 0.7 && partial.indexCoverage! < 0.9, String(partial.indexCoverage));

  const off = await mapAndEstimate({ searchOff: ['aaaaaaaa-0000-0000-0000-000000000003'] }, ['Audit Archive 2016']);
  const audit = off.estimated.libraries.find((l) => l.title === 'Audit Archive 2016')!;
  assert.match(audit.estimate!.error || '', /search found no files/);
  // Its folder figures are still there, and it stays off the dashboard.
  assert.ok(audit.folders && audit.folders.length === 3);
  assert.equal(forDisplay(off.estimated).libraries.length, 0);
});

test('reading an estimated library file by file replaces the estimate', async () => {
  const { sp, estimated } = await mapAndEstimate({}, ['Audit Archive 2016']);
  const service = new StorageScanService(sp.context as never);
  const audit = estimated.libraries.find((l) => l.title === 'Audit Archive 2016')!;
  const read = await service.scanSelected(estimated, [libraryKey(audit)], BASE, () => undefined);
  const after = read.libraries.find((l) => l.title === 'Audit Archive 2016')!;
  assert.equal(after.estimate, undefined);
  assert.equal(after.unscanned, undefined);
  assert.ok(after.folders && after.folders.length === 3, 'the folder breakdown is taken again');
  const shown = forDisplay(read).libraries.find((l) => l.title === 'Audit Archive 2016')!;
  assert.equal(shown.fromEstimate, undefined);
  assert.equal(shown.files, after.files);
});
