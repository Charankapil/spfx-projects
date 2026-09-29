import assert from 'node:assert/strict';
import { test } from 'node:test';

import { AGE_BUCKET_COUNT } from '../../src/webparts/storagePulse/models/IScanResult';
import {
  AGE_BANDS,
  ageInMonths,
  bandIndex,
  emptyFileTypeStat,
  emptyHistogram,
  extensionOf,
  newestFileAge,
  splitByThreshold,
  splitFileType,
  THRESHOLD_OPTIONS,
  totalFileTypes
} from '../../src/webparts/storagePulse/services/activity';
import { format } from '../../src/webparts/storagePulse/components/text';

test('ageInMonths counts whole calendar months', () => {
  assert.equal(ageInMonths(new Date(2025, 8, 30), new Date(2026, 8, 29)), 11);
  assert.equal(ageInMonths(new Date(2025, 8, 29), new Date(2026, 8, 29)), 12);
  assert.equal(ageInMonths(new Date(2026, 0, 31), new Date(2026, 1, 28)), 0);
  assert.equal(ageInMonths(new Date(2026, 0, 31), new Date(2026, 2, 1)), 1);
});

test('ageInMonths clamps future dates to 0 and very old ones to the last bucket', () => {
  assert.equal(ageInMonths(new Date(2030, 0, 1), new Date(2026, 0, 1)), 0);
  assert.equal(ageInMonths(new Date(1990, 0, 1), new Date(2026, 0, 1)), AGE_BUCKET_COUNT - 1);
});

test('every threshold on offer is an age-band boundary', () => {
  const starts = AGE_BANDS.map((b) => b.from);
  for (const t of THRESHOLD_OPTIONS) {
    assert.ok(starts.indexOf(t) >= 0, `threshold ${t} is not a band start`);
  }
});

test('bandIndex puts each month in the band that contains it', () => {
  for (let m = 0; m < AGE_BUCKET_COUNT; m++) {
    const band = AGE_BANDS[bandIndex(m)];
    assert.ok(m >= band.from && (band.to === undefined || m < band.to), `month ${m}`);
  }
});

test('splitByThreshold and splitFileType agree for every threshold', () => {
  const h = emptyHistogram();
  const t = emptyFileTypeStat('pdf');
  for (let m = 0; m < AGE_BUCKET_COUNT; m++) {
    const files = (m % 7) + 1;
    const bytes = files * 1000 + m;
    h.counts[m] += files;
    h.bytes[m] += bytes;
    t.counts[bandIndex(m)] += files;
    t.bytes[bandIndex(m)] += bytes;
  }
  for (const threshold of THRESHOLD_OPTIONS) {
    const a = splitByThreshold(h, threshold);
    const b = splitFileType(t, threshold);
    assert.equal(a.inactiveBytes, b.inactiveBytes);
    assert.equal(a.activeFiles, b.activeFiles);
    assert.equal(a.totalBytes, a.activeBytes + a.inactiveBytes);
  }
});

test('newestFileAge finds the youngest non-empty month', () => {
  const h = emptyHistogram();
  assert.equal(newestFileAge(h), undefined);
  h.counts[14] = 2;
  h.counts[40] = 1;
  assert.equal(newestFileAge(h), 14);
});

test('extensionOf', () => {
  assert.equal(extensionOf('Report.DOCX'), 'docx');
  assert.equal(extensionOf('archive.tar.gz'), 'gz');
  assert.equal(extensionOf('README'), '(none)');
  assert.equal(extensionOf('.gitignore'), '(none)');
  assert.equal(extensionOf('trailing.'), '(none)');
  assert.equal(extensionOf('Minutes v2. final draft'), '(none)');
  assert.equal(extensionOf('x.averyveryverylongext'), '(none)');
});

test('totalFileTypes adds libraries together and is undefined for pre-1.1 scans', () => {
  const a = emptyFileTypeStat('pdf');
  a.bytes[0] = 5;
  const b = emptyFileTypeStat('pdf');
  b.bytes[0] = 7;
  const lib = (fileTypes?: typeof a[]) => ({ fileTypes } as never);
  assert.equal(totalFileTypes([lib(undefined)]), undefined);
  const total = totalFileTypes([lib([a]), lib([b])]);
  assert.equal(total && total[0].bytes[0], 12);
});

test('format fills placeholders and leaves unknown ones', () => {
  assert.equal(format('{a} of {b} {c}', { a: 1, b: 'x' }), '1 of x {c}');
});
