import assert from 'node:assert/strict';
import { test } from 'node:test';

import { encodePath, safeHref, sanitizeResult, serverRelativeHref } from '../../src/webparts/storagePulse/services/safeData';
import { emptyHistogram } from '../../src/webparts/storagePulse/services/activity';

const ORIGIN = 'https://contoso.sharepoint.com';

test('safeHref only allows https links on the tenant host', () => {
  assert.equal(safeHref('javascript:alert(1)', ORIGIN), undefined);
  assert.equal(safeHref('JaVaScRiPt:alert(1)', ORIGIN), undefined);
  assert.equal(safeHref('data:text/html,<script>1</script>', ORIGIN), undefined);
  assert.equal(safeHref('http://contoso.sharepoint.com/sites/a', ORIGIN), undefined);
  assert.equal(safeHref('https://evil.example.com/sites/a', ORIGIN), undefined);
  assert.equal(safeHref('https://contoso.sharepoint.com.evil.com/x', ORIGIN), undefined);
  assert.equal(safeHref('https://contoso.sharepoint.com/sites/a b', ORIGIN), 'https://contoso.sharepoint.com/sites/a%20b');
});

test('serverRelativeHref needs a single leading slash', () => {
  assert.equal(serverRelativeHref('//evil.example.com/x', ORIGIN), undefined);
  assert.equal(serverRelativeHref('javascript:alert(1)', ORIGIN), undefined);
  assert.equal(
    serverRelativeHref('/sites/a/Shared Documents/Q1 #1 100%.pdf', ORIGIN),
    'https://contoso.sharepoint.com/sites/a/Shared%20Documents/Q1%20%231%20100%25.pdf'
  );
});

test('encodePath does not double-encode', () => {
  assert.equal(encodePath('/sites/a%20b/c d'), '/sites/a%20b/c%20d');
});

function validResult(): Record<string, unknown> {
  return {
    scope: 'siteCollection',
    rootUrl: `${ORIGIN}/sites/finance`,
    rootTitle: 'Finance',
    webs: [{ title: 'Finance', url: `${ORIGIN}/sites/finance` }],
    libraries: [
      {
        id: '1',
        title: 'Documents',
        webTitle: 'Finance',
        webUrl: `${ORIGIN}/sites/finance`,
        url: `${ORIGIN}/sites/finance/Shared%20Documents`,
        itemCount: 3,
        files: 3,
        bytes: 30,
        histogram: emptyHistogram()
      }
    ],
    largestOldFiles: [],
    scanStartedAt: '2026-09-01T00:00:00Z',
    scanCompletedAt: '2026-09-01T00:05:00Z',
    scannedBy: 'Alex'
  };
}

test('sanitizeResult accepts a scan this web part wrote', () => {
  const clean = sanitizeResult(validResult(), ORIGIN);
  assert.ok(clean);
  assert.equal(clean!.libraries.length, 1);
});

test('sanitizeResult strips dangerous links instead of rendering them', () => {
  const raw = validResult();
  const libs = raw.libraries as Record<string, unknown>[];
  libs[0].url = 'javascript:alert(document.cookie)';
  libs[0].webUrl = 'https://evil.example.com';
  (raw.webs as Record<string, unknown>[])[0].url = 'javascript:alert(1)';
  raw.largestOldFiles = [
    { name: 'x', serverRelativeUrl: 'javascript:alert(1)', bytes: 1, modified: '', ageMonths: 20 },
    { name: 'y', serverRelativeUrl: '/sites/finance/y.pdf', bytes: 1, modified: '2020-01-01', ageMonths: 20 }
  ];
  const clean = sanitizeResult(raw, ORIGIN)!;
  assert.equal(clean.libraries[0].url, '');
  assert.equal(clean.libraries[0].webUrl, '');
  assert.equal(clean.webs[0].url, '');
  assert.deepEqual(
    clean.largestOldFiles.map((f) => f.name),
    ['y']
  );
});

test('sanitizeResult rejects or repairs malformed data', () => {
  assert.equal(sanitizeResult(null, ORIGIN), undefined);
  assert.equal(sanitizeResult({ libraries: 'x' }, ORIGIN), undefined);
  const noDates = validResult();
  noDates.scanStartedAt = 'yesterday';
  assert.equal(sanitizeResult(noDates, ORIGIN), undefined);
  const otherHost = validResult();
  otherHost.rootUrl = 'https://evil.example.com/sites/finance';
  assert.equal(sanitizeResult(otherHost, ORIGIN), undefined);

  const bad = validResult();
  const libs = bad.libraries as Record<string, unknown>[];
  libs.push({ ...libs[0], histogram: { counts: [1, 2], bytes: [1, 2] } }); // wrong length: dropped
  libs.push({ ...libs[0], bytes: -5, files: 'lots', title: { evil: true } }); // bad numbers / text: repaired
  const clean = sanitizeResult(bad, ORIGIN)!;
  assert.equal(clean.libraries.length, 2);
  assert.equal(clean.libraries[1].bytes, 0);
  assert.equal(clean.libraries[1].files, 0);
  assert.equal(clean.libraries[1].title, '(untitled)');
});
