import assert from 'node:assert/strict';
import { test } from 'node:test';

import { IScanResult } from '../../src/webparts/storagePulse/models/IScanResult';
import { splitByThreshold, splitFileType, THRESHOLD_OPTIONS, totalFileTypes, totalHistogram } from '../../src/webparts/storagePulse/services/activity';
import { describeError } from '../../src/webparts/storagePulse/services/httpErrors';
import { ScanCancelledError, StorageScanService, IScanOptions } from '../../src/webparts/storagePulse/services/StorageScanService';
// @ts-ignore - plain JS test double
import { createMockSharePoint } from '../mock/mockSharePoint.js';

const BASE: IScanOptions = { scope: 'siteCollection', includeHidden: false, excludeSystemLibraries: false, excludedLibraries: [] };

async function scan(mockOptions: Record<string, unknown>, options: Partial<IScanOptions> = {}) {
  const sp = createMockSharePoint(mockOptions);
  const service = new StorageScanService(sp.context as never);
  const result: IScanResult = await service.scan({ ...BASE, ...options }, () => undefined);
  return { sp, result };
}

function totals(result: IScanResult): { files: number; bytes: number } {
  const split = splitByThreshold(totalHistogram(result.libraries), 12);
  return { files: split.totalFiles, bytes: split.totalBytes };
}

test('a full scan counts every visible file exactly, with a large parallel library and no next links', async () => {
  const { sp, result } = await scan({ bigItems: 130000, noNextLink: true, throttleRate: 0.03 });
  const expected = sp.expected(BASE);
  assert.deepEqual(totals(result), { files: expected.files, bytes: expected.bytes });

  // Items hidden by item-level permissions are reported, not silently lost.
  const big = result.libraries.find((l) => l.title === 'Scanned Records')!;
  assert.equal(big.unreadItems, expected.hidden);

  // Large library read as parallel ID ranges; lean responses requested.
  assert.ok(sp.log.some((l: string) => l.indexOf('$filter=Id gt 50000 and Id le 100000') >= 0));
  assert.ok(sp.acceptHeaders.has('application/json;odata.metadata=nometadata'));

  // File types add up to the same totals for every threshold.
  const types = totalFileTypes(result.libraries)!;
  for (const threshold of THRESHOLD_OPTIONS) {
    const overall = splitByThreshold(totalHistogram(result.libraries), threshold);
    const byType = types.map((t) => splitFileType(t, threshold)).reduce((s, t) => s + t.inactiveBytes, 0);
    assert.equal(byType, overall.inactiveBytes, `threshold ${threshold}`);
  }
});

test('a timed-out library is re-read with smaller pages', async () => {
  const { sp, result } = await scan({});
  const audit = result.libraries.find((l) => l.title === 'Audit Archive 2016')!;
  assert.equal(audit.error, undefined);
  assert.ok(audit.files > 0);
  assert.ok(sp.log.some((l: string) => l.indexOf('aaaaaaaa-0000-0000-0000-000000000003') >= 0 && l.indexOf('$top=2500') >= 0));
});

test('an inaccessible subsite is recorded and the scan carries on', async () => {
  const { result } = await scan({});
  const hr = result.webs.find((w) => w.title === 'HR (restricted)')!;
  assert.match(hr.error || '', /Access denied/);
  assert.ok(result.libraries.length > 3);
});

test('hidden and catalog libraries are skipped unless asked for; catalogs always', async () => {
  const off = await scan({});
  assert.ok(!off.result.libraries.some((l) => l.title === 'Preservation Hold Library'));
  const on = await scan({}, { includeHidden: true });
  assert.ok(on.result.libraries.some((l) => l.title === 'Preservation Hold Library'));
  assert.ok(!on.result.libraries.some((l) => l.title === 'Master Page Gallery'));
  assert.deepEqual(totals(on.result), (({ files, bytes }) => ({ files, bytes }))(on.sp.expected({ includeHidden: true })));
});

test('system and named libraries can be excluded', async () => {
  const opts = { excludeSystemLibraries: true, excludedLibraries: ['audit archive 2016'] };
  const { sp, result } = await scan({}, opts);
  const titles = result.libraries.map((l) => l.title);
  assert.ok(titles.indexOf('Site Pages') < 0 && titles.indexOf('Site Assets') < 0 && titles.indexOf('Audit Archive 2016') < 0);
  assert.deepEqual(totals(result), (({ files, bytes }) => ({ files, bytes }))(sp.expected(opts)));
});

test('library links are encoded', async () => {
  const { result } = await scan({});
  const audit = result.libraries.find((l) => l.title === 'Audit Archive 2016')!;
  assert.equal(audit.url, 'https://contoso.sharepoint.com/sites/finance/Audit%20Archive%202016');
  assert.ok(result.libraries.every((l) => l.url.indexOf(' ') < 0));
});

test('Cancel ends the scan at once, even during a long throttling wait', async () => {
  const sp = createMockSharePoint({ alwaysThrottle: true });
  const service = new StorageScanService(sp.context as never);
  let busySeen = false;
  const running = service.scan(BASE, (p) => {
    if (decodeURIComponent(p.currentItem).indexOf('Busy Library') >= 0) {
      busySeen = true;
    }
  });
  const started = Date.now();
  while (!busySeen && Date.now() - started < 20000) {
    await new Promise((r) => setTimeout(r, 50));
  }
  if (!busySeen) {
    service.cancel();
    await running.catch(() => undefined);
    assert.fail('scan never reached the throttled library');
  }
  await new Promise((r) => setTimeout(r, 300)); // now inside the 120 s Retry-After wait
  const cancelledAt = Date.now();
  service.cancel();
  await assert.rejects(running, ScanCancelledError);
  assert.ok(Date.now() - cancelledAt < 1000, 'cancel took too long');
});

test('SharePoint error messages are read from nometadata responses', async () => {
  const body = JSON.stringify({ 'odata.error': { code: 'x', message: { lang: 'en-US', value: 'Access denied.' } } });
  const response = { text: () => Promise.resolve(body), statusText: 'Forbidden' };
  assert.equal(await describeError(response as never), 'Access denied.');
});
