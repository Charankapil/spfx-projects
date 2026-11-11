import assert from 'node:assert/strict';
import { test } from 'node:test';

import { alertStatus, loadAlerts } from '../src/common/alerts';
import { buildBatchBody, parseBatchResponse } from '../src/common/batch';
import { cacheGet, cacheSet } from '../src/common/cache';
import { csvEscape, toCsv } from '../src/common/csv';
import { extensionOf, formatAge, formatBytes, fromDateInput, safeHref, toDateInput } from '../src/common/format';
import { field, ListService } from '../src/common/lists';
import { httpStatusOf } from '../src/common/RequestQueue';
import { fileNameOf, folderOf, search } from '../src/common/search';
import { fakeSp, response, useFakeClock } from './helpers';

const WEB = 'https://contoso.sharepoint.com/sites/team';

test('getAll follows next links and stops at the cap', async () => {
  const clock = useFakeClock();
  const { sp, calls } = fakeSp((call) => {
    const page = Number(/page=(\d+)/.exec(call.url)?.[1] || 0);
    return response(200, { value: [page * 2, page * 2 + 1], '@odata.nextLink': `${WEB}/_api/x?page=${page + 1}` });
  }, clock);
  const result = await sp.getAll<number>(`${WEB}/_api/x?page=0`, 5);
  assert.deepEqual(result.items, [0, 1, 2, 3, 4]);
  assert.equal(result.more, true);
  assert.equal(calls.length, 3);
});

test('getAll reports no more when the collection ends', async () => {
  const clock = useFakeClock();
  const { sp } = fakeSp(() => response(200, { value: [1, 2] }), clock);
  const result = await sp.getAll<number>(`${WEB}/_api/x`, 100);
  assert.deepEqual(result, { items: [1, 2], more: false });
});

test('errors carry the HTTP status and SharePoint’s own message', async () => {
  const clock = useFakeClock();
  const { sp } = fakeSp(() => response(404, { error: { code: 'x', message: "List 'Nope' does not exist." } }), clock);
  await assert.rejects(sp.getJson(`${WEB}/_api/x`), (err: Error) => {
    assert.equal(httpStatusOf(err), 404);
    assert.match(err.message, /does not exist/);
    return true;
  });
});

test('verbose error bodies are understood too', async () => {
  const clock = useFakeClock();
  const { sp } = fakeSp(() => response(400, { 'odata.error': { message: { value: 'Bad refiner.' } } }), clock);
  await assert.rejects(sp.getJson(`${WEB}/_api/x`), /Bad refiner/);
});

test('a post with no content returns undefined, and merge sends the update headers', async () => {
  const clock = useFakeClock();
  const { sp, calls } = fakeSp(() => response(204), clock);
  assert.equal(await sp.post(`${WEB}/_api/x`), undefined);
  await sp.merge(`${WEB}/_api/x`, { Title: 'a' });
  assert.equal(calls[1].headers['X-HTTP-Method'], 'MERGE');
  assert.equal(calls[1].headers['IF-MATCH'], '*');
  assert.equal(calls[1].body, '{"Title":"a"}');
});

test('paths with apostrophes and spaces are escaped in file addresses', () => {
  const clock = useFakeClock();
  const { sp } = fakeSp(() => response(200), clock);
  assert.equal(
    sp.fileApi("/sites/team/Shared Documents/Bob's plan.docx"),
    `${WEB}/_api/web/GetFileByServerRelativePath(decodedurl='%2Fsites%2Fteam%2FShared%20Documents%2FBob''s%20plan.docx')`
  );
  assert.equal(sp.toServerRelative(`${WEB}/Shared Documents/100% done.docx`), '/sites/team/Shared Documents/100% done.docx');
  assert.equal(sp.toServerRelative(`${WEB}/SitePages/My%20Page.aspx`, true), '/sites/team/SitePages/My Page.aspx');
});

test('search builds a safely quoted query and reads both response shapes', async () => {
  const clock = useFakeClock();
  const cells = (path: string, size: string): unknown => ({ Cells: [{ Key: 'Path', Value: path }, { Key: 'Size', Value: size }, { Key: 'Empty', Value: null }] });
  const modern = {
    PrimaryQueryResult: {
      RelevantResults: { TotalRows: 2, Table: { Rows: [cells('https://x/a.docx', '10'), cells('https://x/b.docx', '20')] } },
      RefinementResults: { Refiners: [{ Name: 'FileType', Entries: [{ RefinementName: 'docx', RefinementCount: '2' }, { RefinementName: 'pdf', RefinementCount: '1,204' }] }] }
    }
  };
  const verbose = {
    d: {
      query: {
        PrimaryQueryResult: {
          RelevantResults: { TotalRows: '1', Table: { Rows: { results: [{ Cells: { results: [{ Key: 'Path', Value: 'https://x/c.docx' }] } }] } } },
          RefinementResults: null
        }
      }
    }
  };
  let shape: unknown = modern;
  const { sp, calls } = fakeSp(() => response(200, shape), clock);

  const first = await search(sp, { kql: `Author:"O'Brien" Path:"${WEB}/*"`, select: ['Path', 'Size'], rowLimit: 9999, sort: 'Size:descending', refiners: 'FileType' });
  assert.equal(first.total, 2);
  assert.deepEqual(first.rows[1], { Path: 'https://x/b.docx', Size: '20', Empty: '' });
  assert.deepEqual(first.refiners.FileType, [{ name: 'pdf', count: 1204 }, { name: 'docx', count: 2 }]);
  const url = decodeURIComponent(calls[0].url);
  assert.ok(url.includes(`querytext='Author:"O''Brien"`), url);
  assert.ok(url.includes('rowlimit=500'), 'the row limit is capped at 500');
  assert.ok(url.includes("sortlist='Size:descending'"));
  assert.equal(calls[0].config.name, 'search-v3', 'search must use the OData 3 configuration');

  shape = verbose;
  const second = await search(sp, { kql: 'x', select: [], rowLimit: 10 });
  assert.equal(second.total, 1);
  assert.deepEqual(second.rows, [{ Path: 'https://x/c.docx' }]);
  assert.deepEqual(second.refiners, {});
});

test('file and folder names are taken from a path', () => {
  assert.equal(fileNameOf('https://x/sites/a/Docs/My file.docx'), 'My file.docx');
  assert.equal(folderOf('https://x/sites/a/Docs/My file.docx'), 'https://x/sites/a/Docs');
});

test('a batch body has one change set per update', () => {
  const body = buildBatchBody(
    [
      { url: `${WEB}/_api/web/lists(guid'1')/items(1)`, body: { Title: 'a' } },
      { url: `${WEB}/_api/web/lists(guid'1')/items(2)`, body: { Title: 'b "quoted"' } }
    ],
    'abc'
  );
  assert.equal(body.split('PATCH ').length - 1, 2);
  assert.equal(body.split('--changeset_abc_').length - 1, 4);
  assert.ok(body.endsWith('--batch_abc--\r\n'));
  assert.ok(!/[^\r]\n/.test(body), 'every line must end with CRLF');
  assert.ok(body.includes('{"Title":"b \\"quoted\\""}'));
});

test('a batch response gives one result per update, in order', () => {
  const text = [
    '--batchresponse_1',
    'Content-Type: application/http',
    '',
    'HTTP/1.1 204 No Content',
    'CONTENT-TYPE: application/json',
    '',
    '',
    '--batchresponse_1',
    'Content-Type: application/http',
    '',
    'HTTP/1.1 400 Bad Request',
    'CONTENT-TYPE: application/json',
    '',
    '{"error":{"code":"-1","message":"The value is not valid for column Status."}}',
    '--batchresponse_1',
    'Content-Type: application/http',
    '',
    'HTTP/1.1 429 Too Many Requests',
    'Retry-After: 12',
    '',
    '',
    '--batchresponse_1--'
  ].join('\r\n');
  const results = parseBatchResponse(text);
  assert.equal(results.length, 3);
  assert.deepEqual(results[0], { ok: true, status: 204 });
  assert.equal(results[1].error, 'The value is not valid for column Status.');
  assert.equal(results[2].retryAfter, 12);
});

test('batchMerge repeats the whole batch when one update inside it is throttled', async () => {
  const clock = useFakeClock();
  const part = (status: string, extra = ''): string => `--batchresponse_1\r\nContent-Type: application/http\r\n\r\nHTTP/1.1 ${status}\r\n${extra}\r\n\r\n`;
  let attempt = 0;
  const { sp, calls } = fakeSp(() => {
    attempt++;
    return attempt === 1
      ? response(200, part('204 No Content') + part('429 Too Many Requests', 'Retry-After: 9'))
      : response(200, part('204 No Content') + part('204 No Content'));
  }, clock);
  const results = await sp.batchMerge([
    { url: `${WEB}/_api/a(1)`, body: { A: 1 } },
    { url: `${WEB}/_api/a(2)`, body: { A: 2 } }
  ]);
  assert.deepEqual(results.map((r) => r.ok), [true, true]);
  assert.equal(calls.length, 2);
  assert.ok(calls[1].at - calls[0].at >= 9000, 'the retry did not wait for Retry-After');
  assert.match(calls[0].headers['Content-Type'], /^multipart\/mixed; boundary="batch_[0-9a-f]{32}"$/);
  assert.ok(calls[0].url.endsWith('/_api/$batch'));
});

test('batchMerge refuses an answer that does not cover every update', async () => {
  const clock = useFakeClock();
  const { sp } = fakeSp(() => response(200, '--batchresponse_1\r\n\r\nHTTP/1.1 204 No Content\r\n\r\n'), clock);
  await assert.rejects(sp.batchMerge([{ url: 'a', body: {} }, { url: 'b', body: {} }]), /incomplete/);
});

test('ensure creates a missing list with its columns, and only the missing columns of an existing one', async () => {
  const clock = useFakeClock();
  const definition = {
    urlName: 'ToolkitThings',
    title: 'Things & "stuff"',
    description: 'd',
    fields: [field.text('Colour', 'Colour', true), field.choice('Kind', 'Kind <b>', ['A & B', 'C'])],
    indexAuthor: true,
    ownItemsOnly: true
  };

  let exists = false;
  const first = fakeSp((call) => {
    if (call.method === 'GET') {
      return exists ? response(200, { Id: 'g1', value: [] }) : response(404, { error: { message: 'File Not Found.' } });
    }
    if (call.url.endsWith('/_api/web/lists')) {
      exists = true;
      return response(201, { Id: 'g1' });
    }
    return response(204);
  }, clock);
  await ListService.forWeb(first.sp).ensure(definition);
  const posts = first.calls.filter((c) => c.method === 'POST');
  assert.equal(JSON.parse(posts[0].body || '{}').Title, 'ToolkitThings', 'created under its URL name');
  assert.equal(JSON.parse(posts[1].body || '{}').Title, 'Things & "stuff"', 'then given its display title');
  const xml = posts.filter((c) => c.url.includes('createfieldasxml')).map((c) => JSON.parse(c.body || '{}').parameters.SchemaXml as string);
  assert.equal(xml.length, 2);
  assert.ok(xml[0].includes('Indexed="TRUE"'));
  assert.ok(xml[1].includes('DisplayName="Kind &lt;b&gt;"') && xml[1].includes('<CHOICE>A &amp; B</CHOICE>'), xml[1]);
  assert.ok(posts.some((c) => c.body === '{"WriteSecurity":2}'));
  assert.ok(first.calls[0].url.includes("GetList('%2Fsites%2Fteam%2FLists%2FToolkitThings')"));

  const second = fakeSp((call) => {
    if (call.url.includes('/fields?')) {
      return response(200, { value: [{ InternalName: 'Title' }, { InternalName: 'Colour' }] });
    }
    return call.method === 'GET' ? response(200, { Id: 'g1' }) : response(204);
  }, clock);
  await ListService.forWeb(second.sp).ensure(definition);
  const added = second.calls.filter((c) => c.url.includes('createfieldasxml'));
  assert.equal(added.length, 1);
  assert.ok((added[0].body || '').includes('Name=\\"Kind\\"'));
  assert.ok(!second.calls.some((c) => c.url.endsWith('/_api/web/lists')), 'an existing list is not created again');
});

test('ensure does not hide a real failure behind "missing"', async () => {
  const clock = useFakeClock();
  const { sp } = fakeSp(() => response(403, { error: { message: 'Access denied.' } }), clock);
  await assert.rejects(ListService.forWeb(sp).ensure({ urlName: 'X', title: 'X', description: '', fields: [] }), /Access denied/);
});

test('CSV values are quoted, and formulas are defused', () => {
  assert.equal(csvEscape('=HYPERLINK("http://evil")'), `"'=HYPERLINK(""http://evil"")"`);
  assert.equal(csvEscape('plain'), 'plain');
  assert.equal(csvEscape('a,b'), '"a,b"');
  assert.equal(toCsv(['Name', 'Change'], [['x', -5], ['-cmd', undefined]]), "Name,Change\r\nx,-5\r\n'-cmd,");
});

test('formatting helpers', () => {
  assert.equal(formatBytes(0), '0 B');
  assert.equal(formatBytes(1536), '1.5 KB');
  assert.equal(formatBytes(NaN), '0 B');
  assert.equal(extensionOf('Report.Final.DOCX'), 'docx');
  assert.equal(extensionOf('noextension'), '');
  assert.equal(formatAge(new Date(2024, 0, 1), new Date(2025, 2, 5)), '1 yr 2 mo ago');
  assert.equal(formatAge(new Date(2025, 2, 1), new Date(2025, 2, 5)), '4 days ago');
  assert.equal(toDateInput(fromDateInput('2024-02-29')), '2024-02-29');
  assert.equal(fromDateInput('29/02/2024'), undefined);
});

test('only ordinary web addresses become links', () => {
  assert.equal(safeHref('https://contoso.com/a'), 'https://contoso.com/a');
  assert.equal(safeHref('/sites/hr/SitePages/x.aspx'), '/sites/hr/SitePages/x.aspx');
  assert.equal(safeHref(' javascript:alert(1)'), undefined);
  assert.equal(safeHref('//evil.example'), undefined);
  assert.equal(safeHref('data:text/html,x'), undefined);
  assert.equal(safeHref(undefined), undefined);
});

test('cached values expire', () => {
  cacheSet('session', 'k1', { a: 1 }, 10);
  assert.deepEqual(cacheGet('session', 'k1'), { a: 1 });
  cacheSet('session', 'k2', 'old', -1);
  assert.equal(cacheGet('session', 'k2'), undefined);
  sessionStorage.setItem('k3', 'not json');
  assert.equal(cacheGet('session', 'k3'), undefined);
});

test('alert status follows its dates', () => {
  const now = new Date('2025-06-15T12:00:00Z');
  assert.equal(alertStatus({ Id: 1, Title: 'a' }, now), 'active');
  assert.equal(alertStatus({ Id: 1, Title: 'a', StartDate: '2025-06-16T00:00:00Z' }, now), 'scheduled');
  assert.equal(alertStatus({ Id: 1, Title: 'a', StartDate: '2025-06-01T00:00:00Z', EndDate: '2025-06-14T00:00:00Z' }, now), 'expired');
  assert.equal(alertStatus({ Id: 1, Title: 'a', StartDate: '2025-06-01T00:00:00Z', EndDate: '2025-06-30T00:00:00Z' }, now), 'active');
});

test('the banner asks SharePoint once, however many times it is drawn', async () => {
  const clock = useFakeClock();
  localStorage.clear();
  const { sp, calls } = fakeSp(() => response(200, { value: [{ Id: 1, Title: 'Maintenance' }] }), clock);
  const together = await Promise.all([loadAlerts(sp), loadAlerts(sp), loadAlerts(sp)]);
  assert.equal(calls.length, 1, 'three draws at page load must share one request');
  assert.equal(together[2]?.[0].Title, 'Maintenance');
  await loadAlerts(sp);
  await loadAlerts(sp);
  assert.equal(calls.length, 1, 'later draws must come from the browser cache');
  await loadAlerts(sp, true);
  assert.equal(calls.length, 2, 'a forced refresh asks again');
});

test('a site with no alerts list is remembered, not asked on every page', async () => {
  const clock = useFakeClock();
  localStorage.clear();
  const { sp, calls } = fakeSp(() => response(404, { error: { message: 'File Not Found.' } }), clock);
  assert.equal(await loadAlerts(sp), undefined);
  assert.equal(await loadAlerts(sp), undefined);
  assert.equal(calls.length, 1);
});

test('a failed alert request is not cached and can be retried', async () => {
  const clock = useFakeClock();
  localStorage.clear();
  let fail = true;
  const { sp, calls } = fakeSp(() => (fail ? response(500, { error: { message: 'boom' } }) : response(200, { value: [] })), clock);
  await assert.rejects(loadAlerts(sp), /boom/);
  fail = false;
  assert.deepEqual(await loadAlerts(sp), []);
  assert.equal(calls.length, 2);
});
