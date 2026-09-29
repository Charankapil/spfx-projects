import assert from 'node:assert/strict';
import { test } from 'node:test';

import { IScanResult } from '../../src/webparts/storagePulse/models/IScanResult';
import { sanitizeResult } from '../../src/webparts/storagePulse/services/safeData';
import { splitByThreshold, splitFileType, THRESHOLD_OPTIONS, totalFileTypes, totalHistogram } from '../../src/webparts/storagePulse/services/activity';
import { describeError } from '../../src/webparts/storagePulse/services/httpErrors';
import {
  describeFailure,
  HttpError,
  IScanOptions,
  ScanCancelledError,
  ScanPausedError,
  StorageScanService
} from '../../src/webparts/storagePulse/services/StorageScanService';
// @ts-ignore - plain JS test double
import { createMockSharePoint } from '../mock/mockSharePoint.js';

// Tests run without real-world pacing: the fast profile with no gap between requests, and no wait before the automatic second pass.
const BASE: IScanOptions = {
  scope: 'siteCollection',
  includeHidden: false,
  excludeSystemLibraries: false,
  excludedLibraries: [],
  speed: 'fast',
  autoRetryDelayMs: 0,
  profileOverride: { minGapMs: 0 }
};

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

test('a detailed scan counts every visible file exactly, with a large parallel library and throttling', async () => {
  const { sp, result } = await scan({ bigItems: 130000, throttleRate: 0.03 });
  const expected = sp.expected(BASE);
  assert.deepEqual(totals(result), { files: expected.files, bytes: expected.bytes });

  // Items hidden by item-level permissions are reported, not silently lost.
  const big = result.libraries.find((l) => l.title === 'Scanned Records')!;
  assert.equal(big.unreadItems, expected.hidden);

  // Large library read as parallel ID ranges; lean responses requested.
  assert.ok(sp.log.some((l: string) => l.indexOf('$filter=Id gt 50000 and Id le 100000') >= 0));
  // Never more requests in flight than the limit, even with several libraries and readers at once.
  assert.ok(sp.maxInFlight() <= 6, `max in flight ${sp.maxInFlight()}`);
  assert.ok(sp.maxInFlight() >= 3, 'libraries were not read in parallel');
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

test('SharePoint answering 406 to the fast query falls back and still reads every file', async () => {
  for (const mode of ['orderby', 'both']) {
    const { sp, result } = await scan({ notAcceptable: mode, bigItems: 90000, extraLibraries: 40 });
    const expected = sp.expected(BASE);
    assert.deepEqual(totals(result), { files: expected.files, bytes: expected.bytes }, `mode ${mode}`);
    // Only the restricted HR site fails: no library is abandoned because of a 406.
    assert.deepEqual(
      result.libraries.filter((l) => l.error).map((l) => l.title),
      [],
      `mode ${mode}`
    );
    assert.equal(result.paging && result.paging.level, mode === 'both' ? 2 : 1, `mode ${mode}`);
    // Once a few libraries have learned it, later ones start at the working level instead of failing
    // first: the wasted requests stay a small constant however many libraries there are.
    assert.ok(result.libraries.length > 45);
    assert.ok(sp.notAcceptableCount() <= 20, `${sp.notAcceptableCount()} 406s in mode ${mode}`);
    // The big library still gets its item IDs (default header) or is read as one range; either way it is complete.
    const big = result.libraries.find((l) => l.title === 'Scanned Records')!;
    assert.equal(big.unreadItems, expected.hidden, `mode ${mode}`);
  }
});

test('a scan with no 406 reports no compatibility paging', async () => {
  const { result } = await scan({});
  assert.equal(result.paging, undefined);
});

test('quick scan measures dormant libraries as a whole and agrees with a detailed scan', async () => {
  const detailed = await scan({});
  const quick = await scan({}, { quickAfterMonths: 12 });
  const audit = quick.result.libraries.find((l) => l.title === 'Audit Archive 2016')!;
  assert.ok(audit.measuredAsWhole, 'dormant library was not measured as a whole');
  assert.ok(audit.measuredAsWhole!.dormantMonths >= 40);
  // No file-by-file reads for it.
  assert.ok(!quick.sp.log.some((l: string) => l.indexOf("aaaaaaaa-0000-0000-0000-000000000003')/items?$select=Id,FSObjType") >= 0));
  // Active libraries are still read file by file.
  assert.ok(!quick.result.libraries.find((l) => l.title === 'Documents' && l.webTitle === 'Finance')!.measuredAsWhole);
  // Same totals, and the same split for every period up to the quick threshold.
  assert.deepEqual(totals(quick.result), totals(detailed.result));
  for (const threshold of [3, 6, 12]) {
    const a = splitByThreshold(totalHistogram(quick.result.libraries), threshold);
    const b = splitByThreshold(totalHistogram(detailed.result.libraries), threshold);
    assert.equal(a.inactiveBytes, b.inactiveBytes, `threshold ${threshold}`);
    assert.equal(a.inactiveFiles, b.inactiveFiles, `threshold ${threshold}`);
  }
  assert.equal(quick.result.quickAfterMonths, 12);
  assert.ok(quick.sp.log.length < detailed.sp.log.length, 'quick scan did not save requests');
});

test('Retry failed re-reads only what failed and merges it in', async () => {
  const sp = createMockSharePoint({ flaky: true });
  const service = new StorageScanService(sp.context as never);
  const first = await service.scan(BASE, () => undefined);
  const flaky = first.libraries.find((l) => l.title === 'Flaky Library')!;
  assert.match(flaky.error || '', /server error or timed out/);
  assert.ok(first.webs.find((w) => w.title === 'HR (restricted)')!.error);

  sp.heal();
  const before = sp.log.length;
  const second = await service.retryFailed(first, BASE, () => undefined);
  const again = second.libraries.find((l) => l.title === 'Flaky Library')!;
  assert.equal(again.error, undefined);
  assert.ok(again.files > 0);
  // Only the failed parts were requested again (not the 12,400-item Documents library).
  const retried = sp.log.slice(before);
  assert.ok(!retried.some((l: string) => l.indexOf("aaaaaaaa-0000-0000-0000-000000000001')/items") >= 0));
  assert.deepEqual(totals(second), (({ files, bytes }) => ({ files, bytes }))(sp.expected(BASE)));
  assert.equal(second.libraries.length, first.libraries.length);
  assert.equal(second.scanStartedAt, first.scanStartedAt);
});

test('one unreadable item skips its batch, not the rest of the library', async () => {
  const { result } = await scan({ poisonItem: true });
  const lib = result.libraries.find((l) => l.title === 'Poisoned Library')!;
  assert.ok(lib.partial, 'library should be partly read');
  assert.match(lib.error || '', /1 batch\(es\) of up to 500 items could not be read/);
  // 3,000 items with a folder every 25th: 2,880 files. At most one 500-ID batch is lost.
  assert.ok(lib.files >= 2880 - 500 && lib.files < 2880, `files ${lib.files}`);
});

test('failures are described by cause, not blamed on access', () => {
  assert.match(describeFailure(new HttpError(500, '500: Server is busy.')), /server error or timed out/);
  assert.match(describeFailure(new HttpError(429, '429: throttled')), /throttling/);
  assert.match(describeFailure(new HttpError(403, '403: Access denied.')), /Access denied/);
  assert.match(describeFailure(new HttpError(500, '500: exceeds the list view threshold')), /list view threshold/);
});

test('a busy tenant: throttling never fails a library, and nothing is sent while a Retry-After is active', async () => {
  const sp = createMockSharePoint({ tenantLimit: { perSecond: 6, retryAfterSec: 1 }, extraLibraries: 12 });
  const service = new StorageScanService(sp.context as never);
  const concurrency: number[] = [];
  const result = await service.scan(BASE, (p) => {
    if (p.throttle) {
      concurrency.push(p.throttle.concurrency);
    }
  });
  const expected = sp.expected(BASE);
  assert.deepEqual(totals(result), { files: expected.files, bytes: expected.bytes });
  assert.deepEqual(result.libraries.filter((l) => l.error).map((l) => l.title), []);
  assert.ok(sp.throttles() >= 1, 'the simulated tenant never throttled');
  // The scan slowed itself down after being throttled ...
  assert.ok(concurrency.some((c) => c < 6), `concurrency stayed at ${Math.max(...concurrency)}`);
  // ... and did not keep sending: only requests already on their way can arrive during a hold.
  assert.ok(sp.violations() <= sp.throttles() * 6, `${sp.violations()} requests sent during ${sp.throttles()} holds`);
});

test('long throttling pauses the scan with its progress saved, and it resumes to exact totals', async () => {
  const sp = createMockSharePoint({ outage: { afterRequests: 25, forMs: 6000 }, extraLibraries: 10 });
  const service = new StorageScanService(sp.context as never);
  const checkpoints: IScanResult[] = [];
  const options: IScanOptions = { ...BASE, throttlePatienceMs: 1500, checkpointEveryMs: 0 };
  let paused: ScanPausedError | undefined;
  try {
    await service.scan(options, () => undefined, (snapshot) => {
      checkpoints.push(snapshot);
    });
    assert.fail('the scan should have paused');
  } catch (err) {
    assert.ok(err instanceof ScanPausedError, `expected a pause, got ${String(err)}`);
    paused = err as ScanPausedError;
  }
  const snapshot = paused!.snapshot as IScanResult;
  assert.ok(snapshot && snapshot.partial, 'the pause carried no snapshot');
  assert.ok(snapshot.partial!.librariesDone > 0 && snapshot.partial!.librariesDone < snapshot.partial!.librariesTotal);
  assert.ok(checkpoints.length >= 1, 'no checkpoint was handed out');
  assert.equal(snapshot.libraries.filter((l) => l.pending).length, snapshot.partial!.librariesTotal - snapshot.partial!.librariesDone);
  // Nothing half-read leaks into the snapshot: pending libraries are empty.
  assert.ok(snapshot.libraries.filter((l) => l.pending).every((l) => l.files === 0 && l.bytes === 0));

  // Saved to Site Assets and loaded back the way a later visit would: through the sanitiser.
  const restored = sanitizeResult(JSON.parse(JSON.stringify(snapshot)), 'https://contoso.sharepoint.com')!;
  assert.ok(restored && restored.partial);

  sp.endOutage();
  const resumed = await new StorageScanService(sp.context as never).resume(restored, BASE, () => undefined);
  const expected = sp.expected(BASE);
  assert.deepEqual(totals(resumed), { files: expected.files, bytes: expected.bytes });
  assert.equal(resumed.partial, undefined);
  assert.ok(resumed.libraries.every((l) => !l.pending && !l.error));
  assert.equal(resumed.scanStartedAt, snapshot.scanStartedAt, 'ages must stay measured from the original start');
  assert.equal(resumed.libraries.length, snapshot.libraries.length);
});

test('a library that failed for a transient reason is retried automatically', async () => {
  const sp = createMockSharePoint({ flaky: 6 }); // fails its first six item requests, then recovers
  const service = new StorageScanService(sp.context as never);
  const result = await service.scan(BASE, () => undefined);
  const flaky = result.libraries.find((l) => l.title === 'Flaky Library')!;
  assert.equal(flaky.error, undefined, flaky.error);
  assert.ok(flaky.files > 0);
  const expected = sp.expected(BASE);
  assert.deepEqual(totals(result), { files: expected.files, bytes: expected.bytes });
  // The failed first attempt did not leave a second copy of its files behind.
  assert.equal(flaky.files, sp.libs.flaky.items.filter((i: { folder?: boolean }) => !i.folder).length);
});

test('slower profiles really send fewer requests per second', async () => {
  const sp = createMockSharePoint({ extraLibraries: 6, latencyMs: 1 });
  const rate = async (options: Partial<IScanOptions>): Promise<number> => {
    const fresh = createMockSharePoint({ extraLibraries: 6, latencyMs: 1 });
    const t0 = Date.now();
    await new StorageScanService(fresh.context as never).scan({ ...BASE, ...options, quickAfterMonths: undefined }, () => undefined);
    return fresh.log.length / ((Date.now() - t0) / 1000);
  };
  const gentle = await rate({ speed: 'gentle', profileOverride: { minGapMs: 40 } });
  const fast = await rate({ speed: 'fast', profileOverride: { minGapMs: 5 } });
  assert.ok(gentle < fast, `gentle ${gentle.toFixed(1)}/s vs fast ${fast.toFixed(1)}/s`);
  assert.ok(gentle <= 1000 / 40 + 1, `gentle ${gentle.toFixed(1)} requests/s exceeds its pacing`);
  assert.ok(sp.log.length === 0);
});
