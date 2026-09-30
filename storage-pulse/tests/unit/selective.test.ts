import assert from 'node:assert/strict';
import { test } from 'node:test';

import { ILibraryResult, IScanResult } from '../../src/webparts/storagePulse/models/IScanResult';
import { canRunScan, isAllowedPerson, parseAllowedPeople } from '../../src/webparts/storagePulse/services/access';
import { sanitizeResult } from '../../src/webparts/storagePulse/services/safeData';
import { buildMap, coverage, isLarge, libraryKey, selectionTotals } from '../../src/webparts/storagePulse/services/siteMap';
import { splitByThreshold, totalHistogram } from '../../src/webparts/storagePulse/services/activity';
import { IScanOptions, StorageScanService } from '../../src/webparts/storagePulse/services/StorageScanService';
// @ts-ignore - plain JS test double
import { createMockSharePoint } from '../mock/mockSharePoint.js';

const BASE: IScanOptions = {
  scope: 'siteCollection',
  includeHidden: false,
  excludeSystemLibraries: false,
  excludedLibraries: [],
  speed: 'fast',
  autoRetryDelayMs: 0,
  profileOverride: { minGapMs: 0 }
};
const ORIGIN = 'https://contoso.sharepoint.com';

function totals(libraries: ILibraryResult[]): { files: number; bytes: number } {
  const split = splitByThreshold(totalHistogram(libraries), 12);
  return { files: split.totalFiles, bytes: split.totalBytes };
}

function itemRequests(log: string[], since = 0): string[] {
  return log.slice(since).filter((l) => l.indexOf('/items') >= 0);
}

test('the map lists every library with SharePoint\'s own size and count, without reading any files', async () => {
  const sp = createMockSharePoint({ extraLibraries: 8 });
  const mapped = await new StorageScanService(sp.context as never).mapSite(BASE, undefined, () => undefined);
  assert.ok(mapped.libraries.length > 10);
  assert.ok(mapped.libraries.every((l) => l.unscanned && l.files === 0 && l.bytes === 0));
  assert.equal(itemRequests(sp.log).length, 0, 'the map must not read items');
  const docs = mapped.libraries.find((l) => l.title === 'Documents' && l.webTitle === 'Finance')!;
  assert.ok(docs.metrics && docs.metrics.totalSize > 0 && (docs.metrics.fileCount || 0) > 0);
  assert.deepEqual(coverage(mapped).scanned, 0);
  // Sites come in discovery order, with subsites indented below their parent.
  const tree = buildMap(mapped);
  assert.equal(tree[0].depth, 0);
  assert.ok(tree.some((n) => n.depth === 1));
  assert.equal(tree[0].all.length, mapped.libraries.length);
});

test('scanning in steps adds up to exactly what one full scan finds', async () => {
  const sp = createMockSharePoint({ extraLibraries: 8 });
  const full = await new StorageScanService(sp.context as never).scan(BASE, () => undefined);

  const service = new StorageScanService(sp.context as never);
  const mapped = await service.mapSite(BASE, undefined, () => undefined);
  const first = mapped.libraries.filter((l) => l.title === 'Documents' || l.title === 'Site Pages');
  const keys = new Set(first.map(libraryKey));
  const before = sp.log.length;
  const step1 = await service.scanSelected(mapped, Array.from(keys), BASE, () => undefined);

  // Only the chosen libraries were read; the others stayed on the map.
  const asked = itemRequests(sp.log, before).join('\n');
  for (const lib of step1.libraries) {
    const read = asked.indexOf(`lists(guid'${lib.id}')/items`) >= 0;
    assert.equal(read, keys.has(libraryKey(lib)), `${lib.title} read=${read}`);
  }
  assert.equal(coverage(step1).scanned, first.length);
  assert.equal(step1.libraries.filter((l) => l.unscanned).length, mapped.libraries.length - first.length);
  const expectedFirst = totals(full.libraries.filter((l) => keys.has(libraryKey(l))));
  assert.deepEqual(totals(step1.libraries.filter((l) => !l.unscanned)), expectedFirst);
  assert.ok(step1.libraries.filter((l) => !l.unscanned).every((l) => !!l.scannedAt));

  // A later visit: everything still on the map.
  const step2 = await service.scanSelected(step1, step1.libraries.filter((l) => l.unscanned).map(libraryKey), BASE, () => undefined);
  assert.equal(coverage(step2).scanned, step2.libraries.length);
  assert.deepEqual(totals(step2.libraries), totals(full.libraries));
  assert.equal(step2.scanStartedAt, mapped.scanStartedAt, 'ages stay measured from the first scan');
  assert.equal(step2.partial, undefined);
});

test('choosing a library that was read before reads it again without double counting', async () => {
  const sp = createMockSharePoint({});
  const service = new StorageScanService(sp.context as never);
  const full = await service.scan(BASE, () => undefined);
  const docs = full.libraries.find((l) => l.title === 'Documents' && l.webTitle === 'Finance')!;
  const again = await service.scanSelected(full, [libraryKey(docs)], BASE, () => undefined);
  assert.deepEqual(totals(again.libraries), totals(full.libraries));
  assert.equal(again.largestOldFiles.length, full.largestOldFiles.length);
});

test('refreshing the map keeps what was read and adds new libraries as unscanned', async () => {
  const sp = createMockSharePoint({});
  const service = new StorageScanService(sp.context as never);
  const full = await service.scan(BASE, () => undefined);
  const grown = createMockSharePoint({ extraLibraries: 3 });
  const refreshed = await new StorageScanService(grown.context as never).mapSite(BASE, full, () => undefined);
  const added = refreshed.libraries.filter((l) => l.unscanned);
  assert.equal(added.length, 3);
  assert.deepEqual(totals(refreshed.libraries.filter((l) => !l.unscanned)), totals(full.libraries));
});

test('map results survive the sanitiser and huge libraries are flagged', async () => {
  const sp = createMockSharePoint({ extraLibraries: 2 });
  const mapped = await new StorageScanService(sp.context as never).mapSite(BASE, undefined, () => undefined);
  const restored = sanitizeResult(JSON.parse(JSON.stringify(mapped)), ORIGIN) as IScanResult;
  assert.equal(restored.libraries.filter((l) => l.unscanned).length, mapped.libraries.length);
  assert.ok(restored.libraries.some((l) => l.lastUserChange));
  const totalsAll = selectionTotals(restored, new Set(restored.libraries.map(libraryKey)));
  assert.equal(totalsAll.count, restored.libraries.length);
  const huge = { ...restored.libraries[0], metrics: { totalSize: 1, fileStreamSize: 1, fileCount: 300000 } };
  assert.ok(isLarge(huge));
  assert.ok(!isLarge({ ...restored.libraries[0], metrics: { totalSize: 1, fileStreamSize: 1, fileCount: 10 } }));
});

test('who can scan: owners, everyone, or only the people named', () => {
  const allowed = parseAllowedPeople('Ada@contoso.com;\n  grace@contoso.com , ');
  assert.deepEqual(allowed, ['ada@contoso.com', 'grace@contoso.com']);
  assert.ok(isAllowedPerson({ email: 'ADA@contoso.com' }, allowed));
  assert.ok(isAllowedPerson({ email: 'other@x.com', loginName: 'i:0#.f|membership|grace@contoso.com' }, allowed));
  assert.ok(!isAllowedPerson({ email: 'eve@contoso.com', loginName: 'i:0#.f|membership|eve@contoso.com' }, allowed));
  assert.ok(!isAllowedPerson({}, allowed));
  assert.ok(!isAllowedPerson({ email: 'ada@contoso.com' }, []));

  assert.equal(canRunScan('owners', true, false), true);
  assert.equal(canRunScan('owners', false, true), false);
  assert.equal(canRunScan('everyone', false, false), true);
  // Naming people replaces ownership: an owner who is not listed cannot scan.
  assert.equal(canRunScan('people', true, false), false);
  assert.equal(canRunScan('people', false, true), true);
});
