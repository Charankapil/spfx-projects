require('./register');
const assert = require('assert');
const path = require('path');
const src = (p) => require(path.join(__dirname, '..', 'src', 'webparts', 'adminCenter', p));

const { SPClient, MIN_GAP_MS, MAX_CONCURRENCY, rowsOf, entityOf, nextLinkOf, odataString, CancelledError } = src('core/SPClient');
const { parseSearch, kqlQuote, SearchApi } = src('services/SearchApi');
const { AdminApi, classifyUser, toLogin, permissionLevel, changeToken, describeChange } = src('services/AdminApi');
const { evaluateHealth } = src('services/HealthEngine');
const { csvCell, toCsv } = src('services/exportCsv');
const { formatBytes } = src('services/format');
const { parseAddresses } = src('services/addresses');

const ORIGIN = 'https://contoso.sharepoint.com';
const tests = [];
const test = (name, fn) => tests.push([name, fn]);
const json = (body, status = 200, headers = {}) => new Response(JSON.stringify(body), { status, headers: Object.assign({ 'content-type': 'application/json' }, headers) });
const { SPHttpClient } = require('./mocks/sp-http');
const make = (handler) => new SPClient(new SPHttpClient(handler), ORIGIN);

// ---------------------------------------------------------------- SPClient
test('GET returns json and counts the request', async () => {
  const c = make(async () => json({ value: [1] }));
  assert.deepStrictEqual(await c.get(ORIGIN + '/_api/web'), { value: [1] });
  assert.strictEqual(c.stats.requests, 1);
});

test('identical GETs are served from cache and in-flight ones are shared', async () => {
  let calls = 0;
  const c = make(async () => {
    calls++;
    await new Promise((r) => setTimeout(r, 50));
    return json({ n: calls });
  });
  const [a, b] = await Promise.all([c.get(ORIGIN + '/_api/a'), c.get(ORIGIN + '/_api/a')]);
  assert.strictEqual(calls, 1, 'in-flight GET shared');
  assert.deepStrictEqual(a, b);
  await c.get(ORIGIN + '/_api/a');
  assert.strictEqual(calls, 1, 'second GET from cache');
  assert.ok(c.stats.cacheHits >= 1);
});

test('a write clears the read cache', async () => {
  let calls = 0;
  const c = make(async (m) => {
    if (m === 'GET') calls++;
    return json({ ok: 1 });
  });
  await c.get(ORIGIN + '/_api/a');
  await c.post(ORIGIN + '/_api/b', { x: 1 });
  await c.get(ORIGIN + '/_api/a');
  assert.strictEqual(calls, 2);
});

test('never exceeds MAX_CONCURRENCY and spaces request starts by MIN_GAP_MS', async () => {
  let active = 0;
  let peak = 0;
  const starts = [];
  const c = make(async () => {
    active++;
    peak = Math.max(peak, active);
    starts.push(Date.now());
    await new Promise((r) => setTimeout(r, 120));
    active--;
    return json({});
  });
  await Promise.all(Array.from({ length: 8 }, (_, i) => c.get(ORIGIN + '/_api/x' + i)));
  assert.ok(peak <= MAX_CONCURRENCY, 'peak concurrency ' + peak);
  for (let i = 1; i < starts.length; i++) {
    assert.ok(starts[i] - starts[i - 1] >= MIN_GAP_MS - 15, `gap ${starts[i] - starts[i - 1]}ms`);
  }
});

test('429 with Retry-After pauses the whole queue and retries', async () => {
  let n = 0;
  const times = [];
  const c = make(async () => {
    times.push(Date.now());
    n++;
    return n === 1 ? json({}, 429, { 'Retry-After': '1' }) : json({ ok: true });
  });
  const t0 = Date.now();
  const r = await c.get(ORIGIN + '/_api/t');
  assert.deepStrictEqual(r, { ok: true });
  assert.ok(Date.now() - t0 >= 950, 'waited for Retry-After');
  assert.strictEqual(c.stats.throttled, 1);
  assert.strictEqual(c.stats.retries, 1);
});

test('other queued requests wait during a throttle pause (global back-off)', async () => {
  let first = true;
  const starts = {};
  const c = make(async (m, url) => {
    starts[url] = (starts[url] || []).concat(Date.now());
    if (url.endsWith('/slow') && first) {
      first = false;
      return json({}, 503, { 'Retry-After': '1' });
    }
    return json({ url });
  });
  const t0 = Date.now();
  await Promise.all([c.get(ORIGIN + '/_api/slow'), c.get(ORIGIN + '/_api/other')]);
  // 'other' was queued behind the throttled request and must not have hit SharePoint before the pause ended
  const otherAt = starts[ORIGIN + '/_api/other'][0] - t0;
  assert.ok(otherAt >= 900, 'queued request must wait out the pause, started at ' + otherAt + 'ms');
  const secondSlow = starts[ORIGIN + '/_api/slow'][1] - t0;
  assert.ok(secondSlow >= 950);
});

test('gives up after repeated throttling and surfaces a friendly error', async () => {
  const c = make(async () => json({}, 429, { 'Retry-After': '0' }));
  await assert.rejects(() => c.get(ORIGIN + '/_api/never'), (e) => /throttling/i.test(e.message) && e.status === 429);
  assert.ok(c.stats.requests <= 5, 'bounded retries: ' + c.stats.requests);
});

test('RateLimit headers nearly exhausted slow the queue down before any 429', async () => {
  let n = 0;
  const c = make(async () => {
    n++;
    return json({}, 200, n === 1 ? { 'RateLimit-Limit': '100', 'RateLimit-Remaining': '5', 'RateLimit-Reset': '1' } : {});
  });
  await c.get(ORIGIN + '/_api/a');
  assert.ok(c.stats.pausedUntil > Date.now() - 50);
  const t = Date.now();
  await c.get(ORIGIN + '/_api/b');
  assert.ok(Date.now() - t >= 800, 'second request waited for the window');
});

test('blocks requests to any other origin', async () => {
  const c = make(async () => json({}));
  await assert.rejects(() => c.get('https://evil.example.com/_api/web'), /different origin/);
  await assert.rejects(() => c.get('https://contoso-admin.sharepoint.com/_api/web'), /different origin/);
  await assert.rejects(() => c.post(ORIGIN + '/_api/web', {}, {}, 'https://evil.example.com/'), /different origin/);
  assert.throws(() => c.assertSameOrigin('http://contoso.sharepoint.com/_api/web'));
});

test('cancelPending rejects queued requests but not finished ones', async () => {
  const c = make(async () => {
    await new Promise((r) => setTimeout(r, 100));
    return json({});
  });
  const ps = Array.from({ length: 6 }, (_, i) => c.get(ORIGIN + '/_api/q' + i).then(() => 'ok', (e) => (e instanceof CancelledError ? 'cancelled' : 'err')));
  await new Promise((r) => setTimeout(r, 20));
  c.cancelPending();
  const res = await Promise.all(ps);
  assert.ok(res.includes('cancelled'));
  assert.ok(res.includes('ok'), 'in-flight finished');
});

test('errors carry SharePoint message and 403 is explained', async () => {
  const c = make(async () => json({ error: { message: { value: 'Access denied.' } } }, 403));
  await assert.rejects(() => c.get(ORIGIN + '/_api/x'), (e) => e.status === 403 && /Access denied/.test(e.message) && /higher permissions/.test(e.message));
});

test('POST sends X-HTTP-Method + IF-MATCH for MERGE and passes webUrl', async () => {
  let seen;
  const c = make(async (m, url, cfg, opts) => {
    seen = { m, url, opts };
    return new Response(null, { status: 204 });
  });
  const r = await c.post(ORIGIN + '/sites/a/_api/web', { Title: 'x' }, { method: 'MERGE' }, ORIGIN + '/sites/a');
  assert.strictEqual(r, undefined);
  assert.strictEqual(seen.opts.headers['X-HTTP-Method'], 'MERGE');
  assert.strictEqual(seen.opts.headers['IF-MATCH'], '*');
  assert.strictEqual(seen.opts.webUrl, ORIGIN + '/sites/a');
  assert.strictEqual(seen.opts.body, '{"Title":"x"}');
});

test('response helpers handle v3 verbose, v4 and bare shapes', () => {
  assert.deepStrictEqual(rowsOf({ d: { results: [1, 2] } }), [1, 2]);
  assert.deepStrictEqual(rowsOf({ value: [3] }), [3]);
  assert.deepStrictEqual(rowsOf([4]), [4]);
  assert.deepStrictEqual(rowsOf(undefined), []);
  assert.deepStrictEqual(entityOf({ d: { a: 1 } }), { a: 1 });
  assert.deepStrictEqual(entityOf({ a: 1 }), { a: 1 });
  assert.strictEqual(nextLinkOf({ 'odata.nextLink': 'u' }), 'u');
  assert.strictEqual(nextLinkOf({ '@odata.nextLink': 'v' }), 'v');
  assert.strictEqual(nextLinkOf({ d: { __next: 'w' } }), 'w');
  assert.strictEqual(odataString("O'Brien"), "O''Brien");
});

// ---------------------------------------------------------------- search
const v3Search = {
  d: {
    query: {
      PrimaryQueryResult: {
        RelevantResults: {
          TotalRows: 2,
          Table: { Rows: { results: [{ Cells: { results: [{ Key: 'Title', Value: 'Site A' }, { Key: 'Path', Value: 'https://contoso.sharepoint.com/sites/a/' }, { Key: 'GroupId', Value: null }] } }] } }
        },
        RefinementResults: { Refiners: { results: [{ Name: 'FileType', Entries: { results: [{ RefinementName: 'docx', RefinementCount: '5' }, { RefinementName: 'pdf', RefinementCount: '9' }] } }] } }
      }
    }
  }
};
const v4Search = {
  PrimaryQueryResult: {
    RelevantResults: { TotalRows: 1, Table: { Rows: [{ Cells: [{ Key: 'Title', Value: 'B' }] }] } },
    RefinementResults: { Refiners: [{ Name: 'FileType', Entries: [{ RefinementName: 'xlsx', RefinementCount: '3' }] }] }
  }
};

test('parseSearch understands OData v3 verbose and v4 shapes', () => {
  const a = parseSearch(v3Search.d ? v3Search : v3Search);
  assert.strictEqual(a.total, 2);
  assert.strictEqual(a.rows[0].Title, 'Site A');
  assert.strictEqual(a.rows[0].GroupId, '');
  assert.deepStrictEqual(a.refiners.FileType, [{ name: 'docx', count: 5 }, { name: 'pdf', count: 9 }]);
  const b = parseSearch(v4Search);
  assert.strictEqual(b.rows[0].Title, 'B');
  assert.strictEqual(b.refiners.FileType[0].count, 3);
  assert.deepStrictEqual(parseSearch({}), { rows: [], total: 0, refiners: {} });
});

test('search requests use OData v3, quote paths and never enumerate beyond rowlimit', async () => {
  const urls = [];
  const cfgs = [];
  const c = make(async (m, url, cfg) => {
    urls.push(decodeURIComponent(url));
    cfgs.push(cfg);
    return json(v3Search);
  });
  const s = new SearchApi(c);
  const ft = await s.fileTypes(ORIGIN + '/sites/a/');
  assert.strictEqual(cfgs[0].v3, true, 'search must be OData 3');
  assert.ok(urls[0].includes(`querytext='IsDocument:1 Path:"${ORIGIN}/sites/a/*"'`), urls[0]);
  assert.ok(urls[0].includes('rowlimit=1'));
  assert.ok(urls[0].includes("refiners='FileType(filter=500/0/*)'"));
  assert.deepStrictEqual(ft.types.map((t) => t.type), ['pdf', 'docx'], 'sorted by count');
  assert.strictEqual(kqlQuote('a"b'), '"ab"', 'quotes cannot break out of KQL phrase');
});

test('fileTypes falls back to default refiner when the option is rejected (but not on 403)', async () => {
  let n = 0;
  const c = make(async (m, url) => {
    n++;
    return decodeURIComponent(url).includes('filter=500') ? json({ error: { message: 'bad' } }, 500) : json(v4Search);
  });
  const r = await new SearchApi(c).fileTypes(ORIGIN + '/sites/a');
  assert.strictEqual(r.topTenOnly, true);
  assert.strictEqual(r.types[0].type, 'xlsx');
  const denied = make(async () => json({ error: { message: 'no' } }, 403));
  await assert.rejects(() => new SearchApi(denied).fileTypes(ORIGIN + '/sites/a'), /Access denied/);
});

test('listSites honours maxSites (truncated) and marks sites on other domains as not manageable', async () => {
  const { c } = docIdSearch(3000);
  const r = await new SearchApi(c).listSites(ORIGIN, 1000);
  assert.strictEqual(r.sites.length, 1000);
  assert.strictEqual(r.truncated, true);
  const od = make(async () => json({ PrimaryQueryResult: { RelevantResults: { TotalRows: 1, Table: { Rows: [{ Cells: [{ Key: 'Title', Value: 'Bob' }, { Key: 'Path', Value: 'https://contoso-my.sharepoint.com/personal/bob' }, { Key: 'DocId', Value: '7' }, { Key: 'GroupId', Value: '00000000-0000-0000-0000-000000000000' }] }] } } } }));
  const r2 = await new SearchApi(od).listSites(ORIGIN);
  assert.strictEqual(r2.sites[0].manageable, false);
  assert.strictEqual(r2.sites[0].groupConnected, false);
});

// ---------------------------------------------------------------- admin api
test('classifyUser', () => {
  assert.strictEqual(classifyUser({ LoginName: 'i:0#.f|membership|bob@contoso.com', PrincipalType: 1 }), 'Member');
  assert.strictEqual(classifyUser({ LoginName: 'i:0#.f|membership|ext_x_gmail.com#ext#@contoso.onmicrosoft.com', PrincipalType: 1 }), 'Guest');
  assert.strictEqual(classifyUser({ LoginName: 'c:0-.f|rolemanager|spo-grid-all-users/abc', PrincipalType: 4 }), 'OrgWide');
  assert.strictEqual(classifyUser({ LoginName: 'c:0(.s|true', PrincipalType: 4 }), 'OrgWide');
  assert.strictEqual(classifyUser({ LoginName: 'c:0t.c|tenant|xyz', PrincipalType: 4 }), 'Group');
  assert.strictEqual(classifyUser({ LoginName: 'SHAREPOINT\\system', PrincipalType: 1 }), 'System');
  assert.strictEqual(classifyUser({ LoginName: 'x', PrincipalType: 1, IsShareByEmailGuestUser: true }), 'Guest');
});

test('toLogin builds claims for e-mail and leaves claims alone', () => {
  assert.strictEqual(toLogin(' a@b.com '), 'i:0#.f|membership|a@b.com');
  assert.strictEqual(toLogin('c:0t.c|tenant|1'), 'c:0t.c|tenant|1');
});

test('permissionLevel maps base permission masks', () => {
  assert.strictEqual(permissionLevel(4294967295, 2147483647), 'Full control');
  assert.strictEqual(permissionLevel(0, 0), 'No access');
  assert.strictEqual(permissionLevel(1, 0), 'Read');
  assert.strictEqual(permissionLevel(0x1 | 0x2 | 0x4 | 0x8, 0), 'Edit / contribute');
  const design = 0x1 | 0x2 | 0x4 | 0x8 | (2 ** 18) | (2 ** 24);
  assert.strictEqual(permissionLevel(design, 0), 'Design / manage');
});

test('changeToken ticks are exact (2020-01-01 = 637134336000000000)', () => {
  assert.strictEqual(changeToken('S', Date.UTC(2020, 0, 1)), '1;1;S;637134336000000000;-1');
  assert.strictEqual(changeToken('S', 0), '1;1;S;621355968000000000;-1');
});

test('describeChange maps types and categories', () => {
  const r = describeChange({ __metadata: { type: 'SP.ChangeGroup' }, ChangeType: 13, GroupId: 5, Time: '2025-01-02T03:04:05Z' });
  assert.deepStrictEqual([r.changeType, r.objectKind, r.objectId, r.severity], ['Member added', 'Group', '5', 'security']);
  assert.strictEqual(describeChange({ __metadata: { type: 'SP.ChangeList' }, ChangeType: 3, ListId: 'abc' }).severity, 'structure');
  assert.strictEqual(describeChange({ ChangeType: 2 }).severity, 'info');
  assert.strictEqual(describeChange({ ChangeType: 99 }).changeType, 'Change 99');
});

test('getUsers pages through nextLink with a hard cap of 5 pages', async () => {
  let calls = 0;
  const c = make(async (m, url) => {
    calls++;
    return json({ value: [{ Id: calls, Title: 'U' + calls, LoginName: 'i:0#.f|membership|u' + calls + '@c.com', PrincipalType: 1 }], 'odata.nextLink': ORIGIN + '/_api/web/siteusers?$skiptoken=' + calls });
  });
  const r = await new AdminApi(c).getUsers(ORIGIN);
  assert.strictEqual(calls, 5);
  assert.strictEqual(r.users.length, 5);
  assert.strictEqual(r.truncated, true);
});

test('getRecycleBin falls back to default ordering when $orderby is rejected, then sorts newest first', async () => {
  const urls = [];
  const c = make(async (m, url) => {
    urls.push(decodeURIComponent(url));
    if (url.includes('orderby')) return json({ error: { message: 'x' } }, 400);
    return json({ value: [{ Id: 'a', LeafName: 'old.docx', ItemState: 1, Size: '10', DeletedDate: '2024-01-01T00:00:00Z' }, { Id: 'b', LeafName: 'new.docx', ItemState: 2, Size: '5', DeletedDate: '2025-01-01T00:00:00Z' }] });
  });
  const items = await new AdminApi(c).getRecycleBin(ORIGIN, 300);
  assert.strictEqual(urls.length, 2);
  assert.ok(urls[0].includes('$orderby=DeletedDate desc') && !urls[1].includes('orderby'));
  assert.deepStrictEqual(items.map((i) => i.id), ['b', 'a']);
  assert.strictEqual(items[0].stage, 2);
  assert.ok(urls[0].includes('$top=300'));
});

test('ensureUser/addUserToGroup/removeUserFromGroup/setSiteAdmin build the right calls', async () => {
  const calls = [];
  const c = make(async (m, url, cfg, opts) => {
    calls.push({ m, url: decodeURIComponent(url), opts });
    return url.includes('ensureuser') ? json({ Id: 12, LoginName: 'i:0#.f|membership|a@b.com', Title: 'A' }) : new Response(null, { status: 204 });
  });
  const api = new AdminApi(c);
  const u = await api.ensureUser(ORIGIN + '/sites/a', 'a@b.com');
  assert.strictEqual(u.id, 12);
  assert.strictEqual(JSON.parse(calls[0].opts.body).logonName, 'i:0#.f|membership|a@b.com');
  await api.addUserToGroup(ORIGIN + '/sites/a', 7, u.loginName);
  assert.ok(calls[1].url.endsWith('/sites/a/_api/web/sitegroups/getbyid(7)/users'));
  assert.strictEqual(JSON.parse(calls[1].opts.body).LoginName, u.loginName);
  await api.removeUserFromGroup(ORIGIN + '/sites/a', 7, 12);
  assert.ok(calls[2].url.endsWith('/users/removebyid(12)'));
  await api.setSiteAdmin(ORIGIN + '/sites/a', 12, false);
  assert.strictEqual(calls[3].opts.headers['X-HTTP-Method'], 'MERGE');
  assert.deepStrictEqual(JSON.parse(calls[3].opts.body), { IsSiteAdmin: false });
});

test('getChanges posts a site-scope query with verbose metadata and no item/file changes', async () => {
  let sent;
  const c = make(async (m, url, cfg, opts) => {
    sent = { url, cfg, body: JSON.parse(opts.body) };
    return json({ d: { results: [{ __metadata: { type: 'SP.ChangeUser' }, ChangeType: 1, UserId: 3, Time: '2025-01-01T00:00:00Z' }, { __metadata: { type: 'SP.ChangeGroup' }, ChangeType: 13, GroupId: 4, Time: '2025-01-02T00:00:00Z' }] } });
  });
  const rows = await new AdminApi(c).getChanges(ORIGIN, 'SITEID', 7);
  assert.strictEqual(sent.cfg.v3, true);
  assert.strictEqual(sent.body.query.Item, false);
  assert.strictEqual(sent.body.query.File, false);
  assert.ok(/^1;1;SITEID;\d{18};-1$/.test(sent.body.query.ChangeTokenStart.StringValue), sent.body.query.ChangeTokenStart.StringValue);
  assert.strictEqual(rows[0].objectKind, 'Group', 'newest first');
});

test('getLists parses numbers delivered as strings and builds absolute library URLs', async () => {
  const c = make(async () => json({ value: [{ Id: 'g', Title: 'Docs', BaseType: 1, ItemCount: '5200', EnableVersioning: false, Hidden: false, NoCrawl: true, HasUniqueRoleAssignments: false, RootFolder: { ServerRelativeUrl: '/sites/a/Shared Documents' } }] }));
  const [l] = await new AdminApi(c).getLists(ORIGIN + '/sites/a');
  assert.strictEqual(l.itemCount, 5200);
  assert.strictEqual(l.kind, 'Library');
  assert.strictEqual(l.url, ORIGIN + '/sites/a/Shared Documents');
  assert.strictEqual(l.noCrawl, true);
});

// ---------------------------------------------------------------- health, csv, misc
const U = (o) => Object.assign({ id: 1, title: 'U', email: '', loginName: 'x', principalType: 1, isSiteAdmin: false, kind: 'Member' }, o);
const L = (o) => Object.assign({ id: 'i', title: 'L', description: '', kind: 'Library', baseTemplate: 101, itemCount: 10, versioning: true, minorVersions: false, moderation: false, hidden: false, noCrawl: false, uniquePermissions: false, isSystem: false, url: '', rootFolder: '' }, o);

test('health: clean site scores 100 with only good findings', () => {
  const h = evaluateHealth({
    site: { storageFraction: 0.2, readOnly: false },
    web: { lastModified: new Date() },
    users: [U({ isSiteAdmin: true }), U({ id: 2, isSiteAdmin: true })],
    lists: [L()]
  });
  assert.strictEqual(h.score, 100);
  assert.ok(h.findings.every((f) => f.severity === 'good'));
});

test('health: problems are found and penalised', () => {
  const h = evaluateHealth({
    site: { storageFraction: 0.95, readOnly: true },
    web: { lastModified: new Date(Date.now() - 400 * 86400000) },
    users: [U({ isSiteAdmin: true }), U({ id: 2, kind: 'Guest' }), U({ id: 3, kind: 'OrgWide', title: 'Everyone except external users' })],
    lists: [L({ itemCount: 6000, title: 'Big' }), L({ id: 'j', title: 'NoVer', versioning: false }), L({ id: 'k', noCrawl: true })]
  });
  const ids = h.findings.filter((f) => f.severity !== 'good').map((f) => f.id);
  ['storage', 'readonly', 'admins', 'orgwide', 'threshold', 'versioning', 'nocrawl', 'stale', 'guests'].forEach((id) => assert.ok(ids.includes(id), 'missing ' + id));
  assert.strictEqual(h.findings[0].severity, 'critical', 'sorted by severity');
  assert.ok(h.score < 60 && h.score >= 0, 'score ' + h.score);
});

test('health: hidden and system lists are ignored; score never negative', () => {
  const h = evaluateHealth({ lists: Array.from({ length: 30 }, (_, i) => L({ id: 'x' + i, itemCount: 9000, versioning: false })), site: { storageFraction: 1 }, users: [] });
  assert.ok(h.score >= 0);
  const h2 = evaluateHealth({ lists: [L({ hidden: true, itemCount: 99999 }), L({ isSystem: true, versioning: false })] });
  assert.ok(h2.findings.every((f) => f.severity === 'good'));
});

test('csv: formula injection is neutralised and quoting is correct', () => {
  assert.strictEqual(csvCell('=HYPERLINK("x")'), '"\'=HYPERLINK(""x"")"');
  assert.strictEqual(csvCell('+1'), "'+1");
  assert.strictEqual(csvCell('-2'), "'-2");
  assert.strictEqual(csvCell('@cmd'), "'@cmd");
  assert.strictEqual(csvCell('a,b'), '"a,b"');
  assert.strictEqual(csvCell(null), '');
  assert.strictEqual(csvCell(new Date(0)), '1970-01-01T00:00:00.000Z');
  assert.strictEqual(toCsv(['a', 'b'], [[1, 'x']]), 'a,b\r\n1,x');
});

test('format and address parsing', () => {
  assert.strictEqual(formatBytes(0), '0 B');
  assert.strictEqual(formatBytes(1536), '1.5 KB');
  assert.strictEqual(formatBytes(5 * 1024 ** 4), '5.0 TB');
  assert.deepStrictEqual(parseAddresses('a@b.com, A@B.com;c@d.com\nnot-an-email\n\n e@f.com '), ['a@b.com', 'c@d.com', 'e@f.com']);
});

// ---------------------------------------------------------------- storage growth
const GE = src('services/GrowthEngine');
const { GrowthStore, HISTORY_FILE } = src('services/GrowthStore');
const GBYTES = 1073741824;
const NOW = 1_800_000_000;
const pts = (gbList, stepDays = 1, frac = 0) => gbList.map((g, i) => [NOW - (gbList.length - 1 - i) * stepDays * 86400, Math.round(g * GBYTES), frac]);
const analyze = (gbList, step, frac, settings) => GE.analyzeSite('u', 't', pts(gbList, step, frac), Object.assign({}, GE.DEFAULT_SETTINGS, settings));

test('growth: single snapshot or <20h apart is a baseline, not an anomaly', () => {
  assert.strictEqual(analyze([10]).status, 'baseline');
  const close = GE.analyzeSite('u', 't', [[NOW - 3600, 1e9, 0], [NOW, 9e10, 0]], GE.DEFAULT_SETTINGS);
  assert.strictEqual(close.status, 'baseline');
});

test('growth: steady growth is ok', () => {
  const r = analyze([100, 100.2, 100.4, 100.6, 100.8, 101, 101.2, 101.4]);
  assert.strictEqual(r.status, 'ok');
  assert.deepStrictEqual(r.reasons, []);
});

test('growth: +40% and 8 GB in a week is flagged; small sites with big % are not', () => {
  const r = analyze([20, 21, 22, 23, 24, 25, 26, 28]); // 20 -> 28 GB in 7 days
  assert.ok(r.status === 'warning' || r.status === 'critical', r.status);
  assert.ok(/Grew 8\.0 GB \(\+40%\) in 7\.0 days/.test(r.reasons[0]), r.reasons[0]);
  const tiny = analyze([0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8]); // +700% but under 5 GB
  assert.strictEqual(tiny.status, 'ok');
});

test('growth: huge absolute growth is critical even if the site is big (percentage small)', () => {
  const r = analyze([1000, 1010, 1020, 1030, 1040, 1050, 1060, 1100]); // +100 GB / week = 10%
  assert.strictEqual(r.status, 'critical');
});

test('growth: shrinking sites are never flagged for growth', () => {
  assert.strictEqual(analyze([100, 90, 80, 70, 60, 50, 40, 30]).status, 'ok');
});

test('growth: a spike far above the usual daily rate is flagged', () => {
  const r = analyze([50, 50.1, 50.2, 50.3, 50.4, 50.5, 50.6, 50.7, 50.8, 56.8], 1, 0, { gbPerWeek: 5, pctPerWeek: 500 });
  assert.ok(r.reasons.some((x) => /Latest growth/.test(x)), JSON.stringify(r.reasons));
  assert.strictEqual(r.status, 'warning');
});

test('growth: quota forecast warns at <=30 days and is critical at <=14 days', () => {
  // 100 GB used = 80% of a 125 GB quota, +1 GB/day -> 25 days to full
  const warn = analyze([93, 94, 95, 96, 97, 98, 99, 100], 1, 800, { gbPerWeek: 500, hugeGbPerWeek: 500 });
  assert.strictEqual(warn.status, 'warning');
  assert.ok(Math.round(warn.daysToFull) === 25);
  // +2 GB/day -> 12.5 days
  const crit = analyze([86, 88, 90, 92, 94, 96, 98, 100], 1, 800, { gbPerWeek: 500, hugeGbPerWeek: 500 });
  assert.strictEqual(crit.status, 'critical');
  const calm = analyze([99, 99.1, 99.2, 99.3, 99.4, 99.5, 99.6, 99.7], 1, 800);
  assert.strictEqual(calm.status, 'ok');
});

test('growth: irregular intervals are normalised per day', () => {
  // 2 snapshots 14 days apart, +14 GB (1 GB/day = 7 GB/week on 28 GB = 25%)
  const r = GE.analyzeSite('u', 't', [[NOW - 14 * 86400, 28 * GBYTES, 0], [NOW, 42 * GBYTES, 0]], GE.DEFAULT_SETTINGS);
  assert.ok(Math.abs(r.growth7Bytes - 7 * GBYTES) < GBYTES * 0.01);
  assert.strictEqual(r.status, 'warning');
});

test('growth: addPoint replaces within 6h, appends after, sorts and caps at 90', () => {
  let p = [[1000, 1, 0]];
  p = GE.addPoint(p, [1000 + 3600, 2, 0]);
  assert.deepStrictEqual(p, [[4600, 2, 0]]);
  p = GE.addPoint(p, [4600 + 7 * 3600, 3, 0]);
  assert.strictEqual(p.length, 2);
  let big = [];
  for (let i = 0; i < 120; i++) big = GE.addPoint(big, [i * 86400, i, 0]);
  assert.strictEqual(big.length, GE.MAX_POINTS);
  assert.strictEqual(big[big.length - 1][1], 119);
});

test('growth: normalizeDoc survives junk and keeps valid data', () => {
  assert.deepStrictEqual(GE.normalizeDoc(null).sites, {});
  assert.deepStrictEqual(GE.normalizeDoc('x').sites, {});
  const d = GE.normalizeDoc({ settings: { pctPerWeek: 40, gbPerWeek: -1, bogus: 1 }, sites: { a: { title: 'A', s: [[2, 5, 0], ['x', 1], [1, 4, 10], [3, -5, 0]] }, b: 'nope', c: { s: 'nope' } }, lastCapture: '77' });
  assert.strictEqual(d.settings.pctPerWeek, 40);
  assert.strictEqual(d.settings.gbPerWeek, GE.DEFAULT_SETTINGS.gbPerWeek, 'invalid value falls back');
  assert.deepStrictEqual(Object.keys(d.sites), ['a']);
  assert.deepStrictEqual(d.sites.a.s, [[1, 4, 10], [2, 5, 0]], 'sorted, junk dropped');
  assert.strictEqual(d.lastCapture, 77);
});

test('growth: analyzeAll sorts critical first, then by weekly growth', () => {
  const doc = GE.emptyDoc();
  doc.sites.calm = { title: 'calm', s: pts([10, 10, 10, 10, 10, 10, 10, 10]) };
  doc.sites.bad = { title: 'bad', s: pts([100, 110, 120, 130, 140, 150, 160, 200]) };
  const all = GE.analyzeAll(doc);
  assert.deepStrictEqual(all.map((a) => a.url), ['bad', 'calm']);
  assert.strictEqual(GE.anomaliesOf(all).length, 1);
});

// A tiny in-memory SharePoint for the history file and per-site usage.
function growthEnv(opts = {}) {
  const env = { file: opts.file === undefined ? undefined : opts.file, writes: 0, usageCalls: [], usage: opts.usage || {}, failUsage: opts.failUsage || (() => false), clock: 0 };
  const client = make(async (m, url, cfg, o) => {
    const u = decodeURIComponent(url);
    if (u.includes('GetFileByServerRelativeUrl')) {
      assert.ok(u.includes(`/sites/admin/SiteAssets/${HISTORY_FILE}`), u);
      return env.file === undefined ? json({ error: { message: 'File Not Found.' } }, 404) : new Response(env.file, { status: 200, headers: { 'content-type': 'application/json' } });
    }
    if (u.includes('Files/add')) {
      assert.ok(u.includes(`GetFolderByServerRelativeUrl('/sites/admin/SiteAssets')`) && u.includes(`url='${HISTORY_FILE}',overwrite=true`), u);
      assert.strictEqual(typeof o.body, 'string');
      env.file = o.body;
      env.writes++;
      return json({ Name: HISTORY_FILE });
    }
    const site = u.replace(/\/_api.*/, '');
    env.usageCalls.push(site);
    if (env.failUsage(site)) return json({ error: { message: 'Access denied' } }, 403);
    return json({ Usage: { Storage: String(env.usage[site] || 1000), StoragePercentageUsed: 0.5 } });
  });
  const store = new GrowthStore(client, new AdminApi(client), ORIGIN + '/sites/admin/');
  return { env, store, client };
}

test('store: missing file loads as empty; track writes one file with the right path/body', async () => {
  const { env, store } = growthEnv();
  assert.deepStrictEqual((await store.load()).sites, {});
  const r = await store.track([{ url: ORIGIN + '/sites/a/', title: 'A' }, { url: ORIGIN + '/sites/b', title: 'B' }, { url: ORIGIN + '/sites/a', title: 'dup' }]);
  assert.strictEqual(r.added, 2);
  assert.deepStrictEqual(Object.keys(JSON.parse(env.file).sites), [ORIGIN + '/sites/a', ORIGIN + '/sites/b']);
});

test('store: tracking is capped at 500 sites', async () => {
  const { store } = growthEnv();
  const many = Array.from({ length: 520 }, (_, i) => ({ url: ORIGIN + '/sites/s' + i, title: 'S' + i }));
  const r = await store.track(many);
  assert.strictEqual(r.added, 500);
  assert.strictEqual(r.capped, true);
});

test('store: capture takes one usage request per site, merges, and skips fresh snapshots', async () => {
  const { env, store } = growthEnv({ usage: { [ORIGIN + '/sites/a']: 5 * GBYTES, [ORIGIN + '/sites/b']: 7 } });
  await store.track([{ url: ORIGIN + '/sites/a', title: 'A' }, { url: ORIGIN + '/sites/b', title: 'B' }]);
  let progress = 0;
  const r = await store.capture({ onProgress: () => progress++ });
  assert.strictEqual(r.captured, 2);
  assert.strictEqual(env.usageCalls.length, 2);
  assert.strictEqual(progress, 2);
  const doc = JSON.parse(env.file);
  assert.strictEqual(doc.sites[ORIGIN + '/sites/a'].s[0][1], 5 * GBYTES);
  assert.strictEqual(doc.sites[ORIGIN + '/sites/a'].s[0][2], 500, 'per-mille of quota');
  assert.ok(doc.lastCapture > 0);
  const again = await store.capture();
  assert.strictEqual(again.skipped, 2);
  assert.strictEqual(env.usageCalls.length, 2, 'no requests for fresh snapshots');
  const forced = await store.capture({ force: true });
  assert.strictEqual(forced.captured, 2);
  assert.strictEqual(JSON.parse(env.file).sites[ORIGIN + '/sites/a'].s.length, 1, 'forced re-capture within 6h replaces the point');
});

test('store: failures are reported per site and do not stop the others', async () => {
  const { env, store } = growthEnv({ failUsage: (s) => s.endsWith('/bad') });
  await store.track([{ url: ORIGIN + '/sites/ok', title: 'ok' }, { url: ORIGIN + '/sites/bad', title: 'bad' }]);
  const r = await store.capture();
  assert.strictEqual(r.captured, 1);
  assert.strictEqual(r.failed.length, 1);
  assert.ok(/Access denied/.test(r.failed[0].message));
  assert.strictEqual(JSON.parse(env.file).sites[ORIGIN + '/sites/ok'].s.length, 1);
});

test('store: circuit breaker stops after 10 consecutive failures instead of hammering', async () => {
  const { env, store } = growthEnv({ failUsage: () => true });
  await store.track(Array.from({ length: 60 }, (_, i) => ({ url: ORIGIN + '/sites/s' + i, title: 'S' })));
  const r = await store.capture();
  assert.strictEqual(r.aborted, true);
  assert.ok(env.usageCalls.length <= 12, 'requests made: ' + env.usageCalls.length);
});

test('store: cancel stops early but keeps what was captured', async () => {
  const { env, store } = growthEnv();
  await store.track(Array.from({ length: 30 }, (_, i) => ({ url: ORIGIN + '/sites/s' + i, title: 'S' })));
  let n = 0;
  const r = await store.capture({ shouldCancel: () => n >= 6, onProgress: () => n++ });
  assert.strictEqual(r.cancelled, true);
  assert.ok(r.captured >= 6 && r.captured < 30);
  const saved = Object.keys(JSON.parse(env.file).sites).filter((u) => JSON.parse(env.file).sites[u].s.length === 1).length;
  assert.strictEqual(saved, r.captured);
});

test('store: writes merge with another admin\'s concurrent changes (re-read before write)', async () => {
  const { env, store } = growthEnv();
  await store.track([{ url: ORIGIN + '/sites/a', title: 'A' }]);
  // another admin tracks site c in between
  const other = JSON.parse(env.file);
  other.sites[ORIGIN + '/sites/c'] = { title: 'C', s: [[1, 1, 0]] };
  env.file = JSON.stringify(other);
  await store.track([{ url: ORIGIN + '/sites/b', title: 'B' }]);
  assert.deepStrictEqual(Object.keys(JSON.parse(env.file).sites).sort(), [ORIGIN + '/sites/a', ORIGIN + '/sites/b', ORIGIN + '/sites/c']);
});

test('store: a 404 on save explains the Site Assets requirement; invalid JSON is reported', async () => {
  const bad = make(async (m, url) => (url.includes('Files/add') ? json({ error: { message: 'not found' } }, 404) : json({ error: { message: 'File Not Found.' } }, 404)));
  const s1 = new GrowthStore(bad, new AdminApi(bad), ORIGIN + '/sites/admin');
  await assert.rejects(() => s1.track([{ url: ORIGIN + '/sites/a', title: 'A' }]), /Site Assets library was not found/);
  const junk = make(async () => new Response('not json', { status: 200 }));
  await assert.rejects(() => new GrowthStore(junk, new AdminApi(junk), ORIGIN + '/sites/admin').load(), /not valid JSON/);
});

test('client: raw bodies are sent as-is (file content), not double-encoded', async () => {
  let body;
  const c = make(async (m, u, cfg, o) => { body = o.body; return json({}); });
  await c.post(ORIGIN + '/_api/x', '{"a":1}', { raw: true });
  assert.strictEqual(body, '{"a":1}');
});

// ---------------------------------------------------------------- tenant CSV import
const { parseCsv, detectDelimiter } = src('services/csv');
const SI = src('services/StorageImport');
const TG = src('services/TenantGrowth');
const { TenantStore, INDEX_FILE } = src('services/TenantStore');

test('csv: quotes, embedded commas/newlines, CRLF, BOM, blank lines', () => {
  const bom = String.fromCharCode(0xfeff);
  const rows = parseCsv(bom + 'Url,Title,Note\r\nhttps://a/sites/x,"Finance, Global","say ""hi""\nline2"\r\n\r\nhttps://a/sites/y,Plain,\r\n');
  assert.deepStrictEqual(rows, [['Url', 'Title', 'Note'], ['https://a/sites/x', 'Finance, Global', 'say "hi"\nline2'], ['https://a/sites/y', 'Plain', '']]);
  assert.deepStrictEqual(parseCsv('a;b\n1;2'), [['a', 'b'], ['1', '2']], 'semicolon detected');
  assert.strictEqual(detectDelimiter('a\tb\tc'), '\t');
  assert.deepStrictEqual(parseCsv('a,b\n1,2'), [['a', 'b'], ['1', '2']], 'no trailing newline');
  assert.deepStrictEqual(parseCsv(''), []);
});

test('import: parseNumber handles locales and junk', () => {
  const n = SI.parseNumber;
  assert.strictEqual(n('1234'), 1234);
  assert.strictEqual(n('1,234,567'), 1234567);
  assert.strictEqual(n('1,234.56'), 1234.56);
  assert.strictEqual(n('1.234,56'), 1234.56);
  assert.strictEqual(n('12,5'), 12.5);
  assert.strictEqual(n(' 42 GB '), 42);
  assert.strictEqual(n('1.2E+9'), 1.2e9);
  assert.ok(isNaN(n('')) && isNaN(n('n/a')) && isNaN(n(undefined)));
});

test('import: column and unit detection for typical exports', () => {
  const spo = [['Url', 'Title', 'StorageUsageCurrent', 'StorageQuota', 'LastContentModifiedDate'], ['https://t.sharepoint.com/sites/a', 'A', '1500', '1048576', ''], ['https://t.sharepoint.com/sites/b', 'B', '300', '1048576', '']];
  const m1 = SI.detectMapping(spo[0], spo);
  assert.deepStrictEqual([m1.url, m1.storage, m1.storageUnit, m1.quota, m1.quotaUnit, m1.title], ['Url', 'StorageUsageCurrent', 'mb', 'StorageQuota', 'mb', 'Title']);
  const flow = [['SiteUrl', 'StorageUsed', 'StorageQuota', 'TimeDeleted'], ['https://t.sharepoint.com/sites/a', '5368709120', '1099511627776', ''], ['https://t.sharepoint.com/sites/b', '104857600', '1099511627776', '']];
  const m2 = SI.detectMapping(flow[0], flow);
  assert.deepStrictEqual([m2.url, m2.storage, m2.storageUnit, m2.quotaUnit, m2.deleted], ['SiteUrl', 'StorageUsed', 'bytes', 'bytes', 'TimeDeleted']);
  const gb = [['Site URL', 'Storage used (GB)'], ['https://t.sharepoint.com/sites/a', '12.5']];
  assert.strictEqual(SI.detectMapping(gb[0], gb).storageUnit, 'gb');
  assert.strictEqual(SI.detectMapping(['Name', 'Foo'], [['Name', 'Foo']]), undefined, 'no url/storage columns -> no guess');
  assert.strictEqual(SI.unitFromHeader('StorageUsedMB'), 'mb');
  assert.strictEqual(SI.unitFromHeader('SizeInBytes'), 'bytes');
  assert.strictEqual(SI.unitFromHeader('StorageUsageCurrent'), undefined);
});

test('import: buildSnapshot converts units, skips deleted / OneDrive / bad rows, de-duplicates, stores relative paths', () => {
  const O = 'https://t.sharepoint.com';
  const rows = [
    ['SiteUrl', 'Title', 'Used', 'Quota', 'Deleted'],
    [O + '/sites/a/', 'Alpha', '2048', '1048576', ''],
    [O + '/sites/b', 'Beta', '1,024', '', ''],
    [O + '/sites/gone', 'Gone', '5', '', '2025-01-01'],
    ['https://t-my.sharepoint.com/personal/bob', 'Bob', '9', '', ''],
    [O + '/sites/bad', 'Bad', 'n/a', '', ''],
    ['not a url', 'x', '1', '', ''],
    [O + '/SITES/A', 'Alpha again', '4096', '1048576', ''],
    [O, 'Root', '10', '', '']
  ];
  const map = { url: 'SiteUrl', storage: 'Used', storageUnit: 'mb', quota: 'Quota', quotaUnit: 'mb', title: 'Title', deleted: 'Deleted', includeDeleted: false, excludeOneDrive: true };
  const r = SI.buildSnapshot(rows, map, O, 1000);
  assert.deepStrictEqual(r.data.u, ['/sites/a', '/sites/b', '/']);
  assert.strictEqual(r.data.b[0], 4096 * 1048576, 'duplicate (case-insensitive) keeps the later row');
  assert.strictEqual(r.data.b[1], 1024 * 1048576, 'thousands separator');
  assert.strictEqual(r.data.q[0], 4, 'per-mille of quota (4096/1048576 = 0.39% -> 4)');
  assert.strictEqual(r.data.q[1], 0, 'no quota');
  assert.deepStrictEqual({ ...r.stats }, { rows: 8, kept: 3, noUrl: 1, badNumber: 1, deleted: 1, archived: 0, oneDrive: 1, duplicates: 1 });
  assert.strictEqual(r.titles['/sites/b'], 'Beta');
  assert.strictEqual(SI.fromPath('/sites/a', O), O + '/sites/a');
  assert.strictEqual(SI.fromPath('/', O), O);
  const withOd = SI.buildSnapshot(rows, Object.assign({}, map, { excludeOneDrive: false }), O, 1);
  assert.strictEqual(withOd.stats.oneDrive, 0);
  assert.ok(withOd.data.u.indexOf('https://t-my.sharepoint.com/personal/bob') >= 0, 'other origins stay absolute');
});

const snap = (t, sites) => ({ t, u: Object.keys(sites), b: Object.values(sites).map((g) => Math.round(g * GBYTES)), q: Object.keys(sites).map(() => 0) });
const D = 86400;

test('tenant: selectSnapshots keeps latest + nearest to 7/1/14… days, at most 8', () => {
  const metas = Array.from({ length: 40 }, (_, i) => ({ t: NOW - (39 - i) * D, file: 'f' + i, sites: 1, bytes: 1 }));
  const sel = TG.selectSnapshots(metas, 8);
  assert.strictEqual(sel.length, 8);
  assert.strictEqual(sel[sel.length - 1].file, 'f39');
  assert.ok(sel.some((m) => NOW - m.t === 7 * D), 'a snapshot exactly 7 days back');
  assert.ok(sel.every((m, i) => i === 0 || m.t > sel[i - 1].t), 'ascending, no duplicates');
  assert.strictEqual(TG.selectSnapshots(metas.slice(0, 1)).length, 1);
  assert.deepStrictEqual(TG.selectSnapshots([]), []);
});

test('tenant: analyzeTenant finds the fast grower across snapshots, matching paths case-insensitively', () => {
  const O = 'https://t.sharepoint.com';
  const days = [7, 5, 3, 2, 1, 0];
  const snaps = days.map((d, i) => snap(NOW - d * D, {
    '/sites/Fast': 20 + i * 6,          // 20 -> 50 GB in 5 steps
    '/sites/calm': 100 + i * 0.1,
    '/sites/shrink': 80 - i * 3,
    ...(i === days.length - 1 ? { '/sites/new': 7 } : {})
  }));
  snaps[snaps.length - 1].u[0] = '/sites/fast'; // case differs in the latest snapshot
  const s = TG.analyzeTenant(snaps, GE.DEFAULT_SETTINGS, { '/sites/fast': 'Fast Site' }, O, NOW);
  assert.strictEqual(s.sites, 4);
  assert.strictEqual(s.basedOnSnapshots, 6);
  assert.strictEqual(s.anomalies.length, 1);
  assert.strictEqual(s.anomalies[0].url, O + '/sites/fast');
  assert.strictEqual(s.anomalies[0].title, 'Fast Site');
  assert.ok(s.anomalies[0].spark.length === 6, 'full series found despite the case change');
  assert.strictEqual(s.growers[0].url, O + '/sites/fast');
  assert.strictEqual(s.largest[0].url, O + '/sites/calm');
  assert.ok(s.totalBytes > 0 && s.growth7Bytes > 0);
  assert.strictEqual(TG.analyzeTenant([], GE.DEFAULT_SETTINGS, {}, O, NOW).sites, 0);
});

test('tenant: analysis of 17,000 sites x 8 snapshots is fast and bounded', () => {
  const sites = {};
  for (let i = 0; i < 17000; i++) sites['/sites/s' + i] = 10 + (i % 50);
  const snaps = [7, 6, 5, 4, 3, 2, 1, 0].map((d, i) => {
    const o = {};
    Object.keys(sites).forEach((k, j) => (o[k] = sites[k] + (j % 997 === 0 ? i * 3 : i * 0.01)));
    return snap(NOW - d * D, o);
  });
  const t0 = Date.now();
  const s = TG.analyzeTenant(snaps, GE.DEFAULT_SETTINGS, {}, 'https://t.sharepoint.com', NOW);
  const ms = Date.now() - t0;
  assert.strictEqual(s.sites, 17000);
  assert.ok(s.anomalies.length > 0 && s.anomalies.length <= TG.MAX_ANOMALIES);
  assert.ok(s.largest.length === TG.TOP_N);
  assert.ok(ms < 8000, 'took ' + ms + 'ms');
  const size = JSON.stringify(s).length;
  assert.ok(size < 400000, 'summary stays small: ' + size);
});

// In-memory SharePoint document library for the TenantStore.
function tenantEnv(opts = {}) {
  const env = { files: Object.assign({}, opts.files), calls: [], deleted: [], writes: [] };
  const client = make(async (m, url, cfg, o) => {
    const u = decodeURIComponent(url);
    env.calls.push(m + ' ' + u.replace(ORIGIN, ''));
    const fileM = /GetFileByServerRelativeUrl\('([^']+)'\)(\/\$value)?/.exec(u);
    if (fileM && m === 'GET') {
      const p = fileM[1];
      if (env.files[p] === undefined) return json({ error: { message: 'File Not Found.' } }, 404);
      if (fileM[2]) return new Response(env.files[p], { status: 200 });
      return json({ Name: p.split('/').pop(), ServerRelativeUrl: p, TimeLastModified: '2026-10-02T03:00:00Z', Length: String(env.files[p].length) });
    }
    if (fileM && o.headers && o.headers['X-HTTP-Method'] === 'DELETE') {
      env.deleted.push(fileM[1]);
      delete env.files[fileM[1]];
      return new Response(null, { status: 204 });
    }
    const add = /GetFolderByServerRelativeUrl\('([^']+)'\)\/Files\/add\(url='([^']+)',overwrite=true\)/.exec(u);
    if (add && m === 'POST') {
      env.files[add[1] + '/' + add[2]] = o.body;
      env.writes.push(add[2]);
      return json({ Name: add[2] });
    }
    if (/GetFolderByServerRelativeUrl\('([^']+)'\)\/Files\?/.test(u)) {
      return json({ value: [{ Name: 'old.csv', ServerRelativeUrl: '/sites/admin/Shared Documents/old.csv', TimeLastModified: '2026-09-01T00:00:00Z', Length: '10' }, { Name: 'notes.txt', ServerRelativeUrl: '/x/notes.txt' }, { Name: 'New.CSV', ServerRelativeUrl: '/sites/admin/Shared Documents/New.CSV', TimeLastModified: '2026-10-02T00:00:00Z', Length: '2048' }] });
    }
    if (/\/Folders\?/.test(u)) return json({ value: [{ Name: 'Forms', ServerRelativeUrl: '/a/Forms' }, { Name: 'Storage', ServerRelativeUrl: '/a/Storage' }] });
    if (/web\/lists\?/.test(u)) return json({ value: [{ Title: 'Documents', Hidden: false, RootFolder: { ServerRelativeUrl: '/sites/admin/Shared Documents' } }, { Title: 'Hidden', Hidden: true, RootFolder: { ServerRelativeUrl: '/x' } }] });
    return json({ error: { message: 'unexpected ' + u } }, 500);
  });
  return { env, store: new TenantStore(client, ORIGIN + '/sites/admin'), client };
}

const csvFor = (gbBySite) => 'SiteUrl,Title,StorageUsed\r\n' + Object.keys(gbBySite).map((k) => `${ORIGIN}${k},"${k.slice(7)}",${Math.round(gbBySite[k] * GBYTES)}`).join('\r\n') + '\r\n';
const MAPPING = { url: 'SiteUrl', storage: 'StorageUsed', storageUnit: 'bytes', quotaUnit: 'mb', title: 'Title', excludeOneDrive: true };

test('tenant store: first import writes one snapshot + index; baseline has no anomalies', async () => {
  const { env, store } = tenantEnv();
  const r = await store.importCsv(csvFor({ '/sites/a': 10, '/sites/b': 20 }), MAPPING, { t: NOW, fileName: 'sites.csv', modified: NOW }, GE.DEFAULT_SETTINGS);
  assert.strictEqual(r.stats.kept, 2);
  assert.strictEqual(r.summary.anomalies.length, 0);
  assert.deepStrictEqual(env.writes.sort(), [INDEX_FILE, `admin-center-snap-${NOW}.json`].sort());
  const idx = JSON.parse(env.files['/sites/admin/SiteAssets/' + INDEX_FILE]);
  assert.strictEqual(idx.snapshots.length, 1);
  assert.strictEqual(idx.summary.sites, 2);
  assert.strictEqual(idx.mapping.storage, 'StorageUsed', 'mapping remembered for next time');
  assert.deepStrictEqual(Object.keys(JSON.parse(env.files[`/sites/admin/SiteAssets/admin-center-snap-${NOW}.json`])).sort(), ['b', 'q', 's', 'titles', 'u', 'v']);
});

test('tenant store: later imports detect fast growth; the index summary is what the dashboard reads', async () => {
  const { env, store } = tenantEnv();
  for (let d = 7; d >= 0; d--) {
    await store.importCsv(csvFor({ '/sites/hot': 20 + (7 - d) * 5, '/sites/calm': 100 }), MAPPING, { t: NOW - d * D, fileName: 'sites.csv' }, GE.DEFAULT_SETTINGS);
  }
  const idx = await store.loadIndex(true);
  assert.strictEqual(idx.snapshots.length, 8);
  assert.strictEqual(idx.summary.anomalies.length, 1);
  assert.strictEqual(idx.summary.anomalies[0].title, 'hot');
  assert.ok(/Grew/.test(idx.summary.anomalies[0].reasons[0]));
});

test('tenant store: a snapshot within 6 h replaces the earlier one (and its file is deleted)', async () => {
  const { env, store } = tenantEnv();
  await store.importCsv(csvFor({ '/sites/a': 10 }), MAPPING, { t: NOW, fileName: 'a.csv' }, GE.DEFAULT_SETTINGS);
  const r = await store.importCsv(csvFor({ '/sites/a': 11 }), MAPPING, { t: NOW + 3600, fileName: 'a.csv' }, GE.DEFAULT_SETTINGS);
  assert.strictEqual(r.replaced, true);
  const idx = JSON.parse(env.files['/sites/admin/SiteAssets/' + INDEX_FILE]);
  assert.strictEqual(idx.snapshots.length, 1);
  assert.deepStrictEqual(env.deleted, [`/sites/admin/SiteAssets/admin-center-snap-${NOW}.json`]);
});

test('tenant store: retention keeps the newest 60 snapshots and deletes older files', async () => {
  const { env, store } = tenantEnv();
  const snapshots = Array.from({ length: 60 }, (_, i) => ({ t: NOW - (60 - i) * D, file: `admin-center-snap-${NOW - (60 - i) * D}.json`, sites: 1, bytes: 1 }));
  snapshots.forEach((s) => (env.files['/sites/admin/SiteAssets/' + s.file] = JSON.stringify({ v: 1, u: ['/sites/a'], b: [1], q: [0] })));
  env.files['/sites/admin/SiteAssets/' + INDEX_FILE] = JSON.stringify({ v: 1, snapshots });
  await store.importCsv(csvFor({ '/sites/a': 10 }), MAPPING, { t: NOW, fileName: 'a.csv' }, GE.DEFAULT_SETTINGS);
  const idx = JSON.parse(env.files['/sites/admin/SiteAssets/' + INDEX_FILE]);
  assert.strictEqual(idx.snapshots.length, 60);
  assert.deepStrictEqual(env.deleted, [`/sites/admin/SiteAssets/${snapshots[0].file}`]);
});

test('tenant store: an import costs about a dozen requests however many sites (17,000 here)', async () => {
  const { env, store } = tenantEnv();
  const big = {};
  for (let i = 0; i < 17000; i++) big['/sites/s' + i] = 10 + (i % 40);
  for (let d = 10; d >= 1; d--) {
    const sn = snap(NOW - d * D, big);
    env.files[`/sites/admin/SiteAssets/admin-center-snap-${sn.t}.json`] = JSON.stringify({ v: 1, u: sn.u, b: sn.b, q: sn.q });
  }
  env.files['/sites/admin/SiteAssets/' + INDEX_FILE] = JSON.stringify({ v: 1, snapshots: Array.from({ length: 10 }, (_, k) => ({ t: NOW - (10 - k) * D, file: `admin-center-snap-${NOW - (10 - k) * D}.json`, sites: 17000, bytes: 1 })) });
  env.calls.length = 0;
  const csv = csvFor(Object.assign({}, big, { '/sites/s5': 400 }));
  const r = await store.importCsv(csv, MAPPING, { t: NOW, fileName: 'sites.csv' }, GE.DEFAULT_SETTINGS);
  assert.strictEqual(r.stats.kept, 17000);
  assert.ok(env.calls.length <= 14, 'requests: ' + env.calls.length + '\n' + env.calls.join('\n'));
  assert.ok(r.summary.anomalies.some((a) => a.url.endsWith('/sites/s5')), 'the 400 GB jump is flagged');
  const idxSize = env.files['/sites/admin/SiteAssets/' + INDEX_FILE].length;
  assert.ok(idxSize < 400000, 'index (read by the dashboard) stays small: ' + idxSize);
});

test('tenant store: recalculate re-analyses stored snapshots with new thresholds; empty store is handled', async () => {
  const { store } = tenantEnv();
  assert.strictEqual(await store.recalculate(GE.DEFAULT_SETTINGS), undefined);
  for (let d = 7; d >= 0; d--) await store.importCsv(csvFor({ '/sites/hot': 20 + (7 - d) * 2, '/sites/calm': 100 }), MAPPING, { t: NOW - d * D, fileName: 'a.csv' }, GE.DEFAULT_SETTINGS);
  assert.strictEqual((await store.loadIndex(true)).summary.anomalies.length, 1, '+14 GB (+70%) in a week is flagged by default');
  const strict = await store.recalculate(Object.assign({}, GE.DEFAULT_SETTINGS, { pctPerWeek: 200, hugeGbPerWeek: 500 }));
  assert.strictEqual(strict.anomalies.length, 0);
  assert.strictEqual((await store.loadIndex(true)).summary.anomalies.length, 0, 'recalculated summary is saved');
  const back = await store.recalculate(GE.DEFAULT_SETTINGS);
  assert.strictEqual(back.anomalies.length, 1);
  assert.strictEqual(back.anomalies[0].title, 'hot', 'titles survive a recalculation');
});

test('tenant store: lists libraries, folders (not Forms) and CSV files newest first; reads text', async () => {
  const { env, store } = tenantEnv({ files: { '/sites/admin/Shared Documents/New.CSV': 'a,b' } });
  assert.deepStrictEqual((await store.listLibraries()).map((l) => l.title), ['Documents']);
  assert.deepStrictEqual((await store.listFolders('/a')).map((f) => f.name), ['Storage']);
  const files = await store.listCsvFiles('/sites/admin/Shared Documents');
  assert.deepStrictEqual(files.map((f) => f.name), ['New.CSV', 'old.csv']);
  assert.strictEqual(await store.readCsv('/sites/admin/Shared Documents/New.CSV'), 'a,b');
  const info = await store.getFileInfo('/sites/admin/Shared Documents/New.CSV');
  assert.strictEqual(info.name, 'New.CSV');
});

test('tenant store: file and folder listings are never served from cache (a new CSV must be found)', async () => {
  let n = 0;
  const c = make(async (m, url) => {
    n++;
    return /Files\?/.test(decodeURIComponent(url)) ? json({ value: Array.from({ length: n }, (_, i) => ({ Name: `f${i}.csv`, ServerRelativeUrl: `/a/f${i}.csv`, TimeLastModified: '2026-01-0' + (i + 1) + 'T00:00:00Z', Length: '1' })) }) : json({ value: [] });
  });
  const store = new TenantStore(c, ORIGIN + '/sites/admin');
  assert.strictEqual((await store.listCsvFiles('/a')).length, 1);
  assert.strictEqual((await store.listCsvFiles('/a')).length, 2, 'second listing reflects the new file');
});

test('tenant store: a CSV with the wrong columns explains itself instead of saving nothing silently', async () => {
  const { store } = tenantEnv();
  await assert.rejects(() => store.importCsv('a,b\n1,2\n', { url: 'a', storage: 'b', storageUnit: 'mb', quotaUnit: 'mb', excludeOneDrive: true }, { t: NOW, fileName: 'x.csv' }, GE.DEFAULT_SETTINGS), /No usable site rows/);
});

// ---------------------------------------------------------------- v1.3: activity fix, bulk remove, org-wide, site states
const AA = src('services/AdminApi');

test('activity: ChangeQuery no longer sends Feature, and drops any property SharePoint rejects, then retries', async () => {
  const sent = [];
  const c = make(async (m, url, cfg, o) => {
    const body = JSON.parse(o.body);
    sent.push(Object.keys(body.query));
    if ('Navigation' in body.query) {
      return json({ error: { code: '-1, Microsoft.SharePoint.Client.InvalidClientQueryException', message: { lang: 'en-US', value: "The property 'Navigation' does not exist on type 'SP.ChangeQuery'. Make sure to only use property names that are defined by the type." } } }, 400);
    }
    return json({ d: { results: [{ __metadata: { type: 'SP.ChangeGroup' }, ChangeType: 13, GroupId: 4, Time: '2026-10-01T00:00:00Z' }] } });
  });
  const rows = await new AdminApi(c).getChanges(ORIGIN, 'SITE', 7);
  assert.strictEqual(rows.length, 1);
  assert.strictEqual(sent.length, 2, 'one retry');
  assert.ok(sent.every((k) => k.indexOf('Feature') < 0), 'Feature is never sent');
  assert.ok(sent[1].indexOf('Navigation') < 0 && sent[1].indexOf('ChangeTokenStart') >= 0, 'only the rejected property is dropped');
});

test('activity: other errors are not retried, and the retry loop is bounded', async () => {
  let n = 0;
  const denied = make(async () => { n++; return json({ error: { message: { value: 'Access denied.' } } }, 403); });
  await assert.rejects(() => new AdminApi(denied).getChanges(ORIGIN, 'S', 7), /Access denied/);
  assert.strictEqual(n, 1);
  let m = 0;
  const props = ['Add', 'Update', 'DeleteObject', 'Restore', 'Rename', 'Site'];
  const stubborn = make(async () => { const p = props[m++ % props.length]; return json({ error: { message: { value: `The property '${p}' does not exist on type 'SP.ChangeQuery'.` } } }, 400); });
  await assert.rejects(() => new AdminApi(stubborn).getChanges(ORIGIN, 'S', 7), /does not exist/);
  assert.ok(m <= 5, 'gave up after ' + m + ' attempts');
  assert.strictEqual(AA.unsupportedChangeQueryProperty({ status: 400, message: "400: The property 'Feature' does not exist on type 'SP.ChangeQuery'." }), 'Feature');
  assert.strictEqual(AA.unsupportedChangeQueryProperty({ status: 500, message: "The property 'Feature' does not exist on type 'SP.ChangeQuery'." }), undefined);
});

test('people: matchUsers finds users by e-mail, claims login or account name; reports the rest', () => {
  const users = [
    { id: 1, title: 'Ann', email: 'Ann@Contoso.com', loginName: 'i:0#.f|membership|ann@contoso.com', principalType: 1, isSiteAdmin: false, kind: 'Member' },
    { id: 2, title: 'Guest', email: '', loginName: 'i:0#.f|membership|bob_fabrikam.com#ext#@contoso.onmicrosoft.com', principalType: 1, isSiteAdmin: false, kind: 'Guest' },
    { id: 3, title: 'Everyone', email: '', loginName: 'c:0(.s|true', principalType: 4, isSiteAdmin: false, kind: 'OrgWide' }
  ];
  const r = AA.matchUsers(['ann@contoso.com', 'ANN@contoso.com', 'bob_fabrikam.com#ext#@contoso.onmicrosoft.com', 'c:0(.s|true', 'nobody@x.com'], users);
  assert.deepStrictEqual(r.matched.map((u) => u.id), [1, 2, 3]);
  assert.deepStrictEqual(r.unmatched, ['nobody@x.com']);
  assert.ok(AA.isOrgWideLogin('c:0-.f|rolemanager|spo-grid-all-users/abc') && AA.isOrgWideLogin('c:0(.s|true'));
  assert.ok(!AA.isOrgWideLogin('i:0#.f|membership|a@b.com'));
});

test('people: removing a direct permission deletes the role assignment of that principal', async () => {
  let seen;
  const c = make(async (m, url, cfg, o) => { seen = { url: decodeURIComponent(url), o }; return new Response(null, { status: 204 }); });
  await new AdminApi(c).removeRoleAssignment(ORIGIN + '/sites/a', 17);
  assert.ok(seen.url.endsWith('/sites/a/_api/web/roleassignments/getbyprincipalid(17)'));
  assert.strictEqual(seen.o.headers['X-HTTP-Method'], 'DELETE');
});

test('import: archive values and deleted values are recognised', () => {
  ['FullyArchived', 'RecentlyArchived', 'Reactivating', 'True', 'Archived', '2026-01-02'].forEach((v) => assert.ok(SI.isArchivedValue(v), v));
  ['', 'NotArchived', 'Not Archived', 'none', 'false', '0', 'No', 'null'].forEach((v) => assert.ok(!SI.isArchivedValue(v), v));
  assert.ok(SI.isDeletedValue('2026-01-01') && !SI.isDeletedValue('') && !SI.isDeletedValue('null'));
  const m = SI.detectMapping(['SiteUrl', 'StorageUsed', 'ArchiveStatus', 'TimeDeleted'], [['SiteUrl', 'StorageUsed', 'ArchiveStatus', 'TimeDeleted']]);
  assert.strictEqual(m.archived, 'ArchiveStatus');
  assert.strictEqual(m.includeDeleted, true);
});

test('import: sites are kept as active / archived / deleted; deleted can still be skipped', () => {
  const O = 'https://t.sharepoint.com';
  const rows = [
    ['SiteUrl', 'StorageUsed', 'ArchiveStatus', 'TimeDeleted'],
    [O + '/sites/a', '10', 'NotArchived', ''],
    [O + '/sites/b', '20', 'FullyArchived', ''],
    [O + '/sites/c', '30', '', '2026-09-01'],
    [O + '/sites/d', '40', 'RecentlyArchived', '2026-09-02']
  ];
  const map = { url: 'SiteUrl', storage: 'StorageUsed', storageUnit: 'mb', quotaUnit: 'mb', archived: 'ArchiveStatus', deleted: 'TimeDeleted', includeDeleted: true, excludeOneDrive: true };
  const r = SI.buildSnapshot(rows, map, O, 1);
  assert.deepStrictEqual(r.data.s, [0, 1, 2, 2], 'deleted wins over archived');
  assert.strictEqual(r.stats.kept, 4);
  assert.strictEqual(r.stats.archived, 1);
  assert.strictEqual(r.stats.deleted, 2);
  const skip = SI.buildSnapshot(rows, Object.assign({}, map, { includeDeleted: false }), O, 1);
  assert.deepStrictEqual(skip.data.u, ['/sites/a', '/sites/b']);
  assert.strictEqual(skip.stats.deleted, 2);
});

test('tenant: segments total each state; deleted sites never raise growth alerts', () => {
  const O = 'https://t.sharepoint.com';
  const mk = (t, gbA, gbB, gbC) => ({ t, u: ['/sites/a', '/sites/b', '/sites/c'], b: [gbA, gbB, gbC].map((g) => g * GBYTES), q: [0, 0, 0], s: [0, 1, 2] });
  const snaps = [mk(NOW - 7 * D, 10, 50, 30), mk(NOW, 40, 51, 90)];
  const s = TG.analyzeTenant(snaps, GE.DEFAULT_SETTINGS, {}, O, NOW);
  assert.deepStrictEqual([s.segments.active.sites, s.segments.archived.sites, s.segments.deleted.sites], [1, 1, 1]);
  assert.strictEqual(s.segments.deleted.bytes, 90 * GBYTES);
  assert.strictEqual(s.segments.active.bytes, 40 * GBYTES);
  assert.deepStrictEqual(s.anomalies.map((a) => a.url), [O + '/sites/a'], 'the deleted site grew 60 GB but is not alerted');
  assert.ok(s.largest.every((r) => r.state !== 2));
  assert.strictEqual(s.anomalies[0].state, 0);
  const old = TG.analyzeTenant([{ t: NOW, u: ['/x'], b: [1], q: [0] }], GE.DEFAULT_SETTINGS, {}, O, NOW);
  assert.strictEqual(old.segments.active.sites, 1, 'snapshots without states count as active');
});

// ---------------------------------------------------------------- v1.4: full inventory, tenant dashboard
const CAP = src('services/capacity');

function docIdSearch(total, opts = {}) {
  const calls = [];
  const c = make(async (m, url) => {
    const u = decodeURIComponent(url);
    calls.push(u);
    const qt = /querytext='([^']*)'/.exec(u)[1];
    const after = /IndexDocId>(\d+)/.exec(qt);
    const start = after ? Number(after[1]) : Number((/startrow=(\d+)/.exec(u) || [0, 0])[1]);
    const rows = [];
    for (let id = start + 1; id <= Math.min(total, start + 500); id++) {
      const cells = [{ Key: 'Title', Value: 'S' + id }, { Key: 'Path', Value: ORIGIN + '/sites/s' + id }, { Key: 'WebTemplate', Value: 'GROUP' }];
      if (!opts.noDocId) cells.push({ Key: 'DocId', Value: String(id) });
      rows.push({ Cells: cells });
    }
    return json({ PrimaryQueryResult: { RelevantResults: { TotalRows: total, Table: { Rows: rows } } } });
  });
  return { c, calls };
}

test('inventory: 17,500 sites are all fetched with DocId paging (35 requests, no 5,000 cap)', async () => {
  const { c, calls } = docIdSearch(17500);
  let progress = 0;
  const r = await new SearchApi(c).listSites(ORIGIN, undefined, (n) => (progress = n));
  assert.strictEqual(r.sites.length, 17500);
  assert.strictEqual(r.truncated, false);
  assert.strictEqual(calls.length, 36, 'last short page ends it');
  assert.ok(calls[0].includes("sortlist='[DocId]:ascending'") && calls[1].includes('IndexDocId>500'), calls[1]);
  assert.strictEqual(progress, 17500);
  assert.strictEqual(new Set(r.sites.map((s) => s.url)).size, 17500, 'no duplicates');
});

test('inventory: without DocId it falls back to startrow paging and still gets everything', async () => {
  const { c, calls } = docIdSearch(1200, { noDocId: true });
  const r = await new SearchApi(c).listSites(ORIGIN);
  assert.strictEqual(r.sites.length, 1200);
  assert.ok(calls.some((u) => /startrow=1000/.test(u)), 'used startrow');
});

test('import: template, Teams and last-activity columns are detected and stored compactly', () => {
  const O = 'https://t.sharepoint.com';
  const rows = [
    ['SiteUrl', 'StorageUsed', 'TemplateName', 'IsTeamsConnected', 'LastActivityOn'],
    [O + '/sites/a', '10', 'GROUP#0', 'True', '2026-09-30T10:00:00Z'],
    [O + '/sites/b', '20', 'SITEPAGEPUBLISHING#0', 'False', ''],
    [O + '/sites/c', '30', 'group#0', '1', '/Date(1767225600000)/']
  ];
  const m = SI.detectMapping(rows[0], rows.slice(1));
  assert.deepStrictEqual([m.template, m.teams, m.lastActivity], ['TemplateName', 'IsTeamsConnected', 'LastActivityOn']);
  const r = SI.buildSnapshot(rows, m, O, 1);
  assert.deepStrictEqual(r.data.tpl, ['GROUP#0', 'SITEPAGEPUBLISHING#0']);
  assert.deepStrictEqual(r.data.tp, [0, 1, 0], 'case-insensitive template dictionary');
  assert.deepStrictEqual(r.data.tm, [1, 0, 1]);
  assert.strictEqual(r.data.la[1], 0);
  assert.strictEqual(r.data.la[2], Math.floor(1767225600000 / 86400000));
  const plain = SI.buildSnapshot([['SiteUrl', 'StorageUsed'], [O + '/sites/a', '1']], { url: 'SiteUrl', storage: 'StorageUsed', storageUnit: 'mb', quotaUnit: 'mb', excludeOneDrive: true }, O, 1);
  assert.strictEqual(plain.data.tp, undefined, 'no extra arrays when columns are not mapped');
});

test('tenant dashboard: inventory overview counts, bands, Teams, top 50 (deleted kept apart)', () => {
  const O = 'https://t.sharepoint.com';
  const n = 120;
  const today = Math.floor((NOW * 1000) / 86400000);
  const snap = { t: NOW, u: [], b: [], q: [], s: [], tp: [], tpl: ['GROUP#0', 'SITEPAGEPUBLISHING#0'], tm: [], la: [] };
  for (let i = 0; i < n; i++) {
    snap.u.push('/sites/s' + i);
    snap.b.push((i + 1) * GBYTES);
    snap.q.push(0);
    snap.s.push(i === 119 ? 2 : i % 10 === 0 ? 1 : 0);
    snap.tp.push(i % 3 === 0 ? 1 : 0);
    snap.tm.push(i % 2);
    snap.la.push(i < 5 ? 0 : today - i * 4);
  }
  const typeOf = (t) => (t.indexOf('GROUP') === 0 ? 'Team site (M365 group)' : 'Communication site');
  const o = TG.summarizeInventory(snap, { '/sites/s118': 'Big one' }, O, NOW, typeOf);
  assert.strictEqual(o.sites, 120);
  assert.deepStrictEqual([o.activeSites, o.archivedSites, o.deletedSites], [107, 12, 1]);
  assert.strictEqual(o.deletedBytes, 120 * GBYTES);
  assert.strictEqual(o.top.length, 50);
  assert.strictEqual(o.top[0].url, O + '/sites/s118', 'deleted s119 is not in the top list');
  assert.strictEqual(o.top[0].title, 'Big one');
  assert.strictEqual(o.sizeBands.reduce((s, b) => s + b.count, 0), 119);
  assert.strictEqual(o.sizeBands[0].count, 0);
  assert.strictEqual(o.sizeBands[1].count, 9, '1-10 GB: sites of 1..9 GB');
  assert.strictEqual(o.teamsSites, 59);
  assert.ok(o.byType.length === 2 && o.byType.reduce((s, b) => s + b.count, 0) === 119);
  assert.strictEqual(o.activity[o.activity.length - 1].count, 5, 'unknown activity');
  const bare = TG.summarizeInventory({ t: NOW, u: ['/x'], b: [1], q: [0] }, {}, O, NOW, typeOf);
  assert.strictEqual(bare.teamsSites, undefined);
  assert.strictEqual(bare.byType, undefined);
  assert.strictEqual(bare.activity, undefined);
});

test('tenant dashboard: capacity maths (archived billed separately; deleted optional)', () => {
  const TB = 1099511627776;
  const o = { activeBytes: 400 * TB, archivedBytes: 50 * TB, deletedBytes: 10 * TB };
  const c = CAP.capacityOf(o, 600, true);
  assert.strictEqual(c.usedBytes, 410 * TB);
  assert.strictEqual(c.leftBytes, 190 * TB);
  assert.ok(Math.abs(c.fraction - 410 / 600) < 1e-9);
  assert.strictEqual(CAP.capacityOf(o, 600, false).usedBytes, 400 * TB);
  assert.strictEqual(CAP.capacityOf(o, undefined, true).leftBytes, undefined);
  assert.strictEqual(CAP.formatTB(434.8 * TB), '435 TB');
  assert.strictEqual(CAP.formatTB(18.88 * TB), '18.9 TB');
  assert.strictEqual(CAP.formatTB(0.5 * TB), '0.50 TB');
});

test('tenant store: import stores the overview; capacity is saved and survives later imports', async () => {
  const { env, store } = tenantEnv();
  const csv = 'SiteUrl,Title,StorageUsed,TemplateName,IsTeamsConnected\r\n' + [1, 2, 3].map((i) => `${ORIGIN}/sites/x${i},"X ${i}",${i * GBYTES},GROUP#0,${i === 1 ? 'True' : 'False'}`).join('\r\n');
  const m = SI.detectMapping(parseCsv(csv)[0], parseCsv(csv).slice(1));
  await store.importCsv(csv, m, { t: NOW - 2 * D, fileName: 'a.csv' }, GE.DEFAULT_SETTINGS);
  let idx = await store.loadIndex(true);
  assert.strictEqual(idx.overview.sites, 3);
  assert.strictEqual(idx.overview.teamsSites, 1);
  assert.strictEqual(idx.overview.top[0].title, 'X 3');
  await store.saveCapacity(603.03, false);
  await store.importCsv(csv, m, { t: NOW, fileName: 'b.csv' }, GE.DEFAULT_SETTINGS);
  idx = await store.loadIndex(true);
  assert.strictEqual(idx.capacityTB, 603.03, 'capacity kept across imports');
  assert.strictEqual(idx.countDeleted, false);
  const snapFull = await store.loadLatestSnapshot();
  assert.strictEqual(snapFull.u.length, 3);
  assert.strictEqual(snapFull.titles['/sites/x2'], 'X 2', 'titles travel with the snapshot');
  assert.deepStrictEqual(snapFull.tm, [1, 0, 0]);
  const re = await store.recalculate(GE.DEFAULT_SETTINGS);
  assert.ok(re);
  idx = await store.loadIndex(true);
  assert.strictEqual(idx.overview.top[0].title, 'X 3', 'recalculate rebuilds the overview with titles');
});

// ---------------------------------------------------------------- runner
(async () => {
  let failed = 0;
  for (const [name, fn] of tests) {
    const t0 = Date.now();
    try {
      await fn();
      console.log(`  ok   ${name} (${Date.now() - t0}ms)`);
    } catch (e) {
      failed++;
      console.log(`  FAIL ${name}\n       ${e && e.stack ? e.stack.split('\n').slice(0, 4).join('\n       ') : e}`);
    }
  }
  console.log(`\n${tests.length - failed}/${tests.length} passed`);
  process.exit(failed ? 1 : 0);
})();
