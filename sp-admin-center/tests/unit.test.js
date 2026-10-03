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

test('listSites pages 500 at a time, stops on a short page, flags truncation, marks other domains', async () => {
  const page = (n, prefix) => ({ PrimaryQueryResult: { RelevantResults: { TotalRows: n, Table: { Rows: Array.from({ length: n }, (_, i) => ({ Cells: [{ Key: 'Title', Value: 'S' + i }, { Key: 'Path', Value: prefix + '/sites/s' + i }, { Key: 'WebTemplate', Value: 'STS' }, { Key: 'GroupId', Value: '00000000-0000-0000-0000-000000000000' }] })) } } } });
  let calls = 0;
  const c1 = make(async (m, url) => {
    calls++;
    return json(page(calls === 1 ? 500 : 20, ORIGIN));
  });
  const r1 = await new SearchApi(c1).listSites(ORIGIN, 10);
  assert.strictEqual(calls, 2);
  assert.strictEqual(r1.sites.length, 520);
  assert.strictEqual(r1.truncated, false);
  assert.strictEqual(r1.sites[0].groupConnected, false);
  let n = 0;
  const c2 = make(async () => {
    n++;
    return json(page(500, 'https://contoso-my.sharepoint.com/personal'));
  });
  const r2 = await new SearchApi(c2).listSites(ORIGIN, 3);
  assert.strictEqual(n, 3, 'hard page cap');
  assert.strictEqual(r2.truncated, true);
  assert.strictEqual(r2.sites[0].manageable, false);
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
