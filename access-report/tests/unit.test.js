// Unit tests for the web part's services, run on the real TypeScript sources under Node.
require('./register');
const assert = require('assert');
const path = require('path');
const src = (p) => require(path.join(__dirname, '..', 'src', 'webparts', 'accessReport', p));

const { SPClient, CancelledError, HttpError } = src('core/SPClient');
const P = src('services/permissions');
const C = src('services/reportCsv');
const { AccessScanner, grantsFor, levelOf, roleNames } = src('services/AccessScan');
const { SiteLister, inScope, isOneDriveSite } = src('services/Sites');
const { Store, folderProblem, aliasValue } = src('services/Store');
const { PeoplePicker, fromPicker } = src('services/PeoplePicker');
const R = src('services/RunJob');
const { compareReports, changeRows } = src('services/compare');
const { crc32, utf8, zip } = src('flow/zip');
const { FakeTenant, ORIGIN, FULL, READ, EDIT, DESIGN, LIMITED } = require('./fakeTenant');
const scenarios = require('./scenarios');
const { SPHttpClient } = require('./mocks/sp-http');

console.error = () => undefined; // SPClient logs every failed request; many tests fail requests on purpose
const local = {};
global.window = { localStorage: { getItem: (k) => (k in local ? local[k] : null), setItem: (k, v) => (local[k] = String(v)), removeItem: (k) => delete local[k] } };

const tests = [];
const test = (name, fn) => tests.push([name, fn]);
const json = (body, status = 200, headers = {}) => new Response(JSON.stringify(body), { status, headers: Object.assign({ 'content-type': 'application/json' }, headers) });
const make = (handler, gap = 0) => new SPClient(new SPHttpClient(handler), ORIGIN, gap);
const clientOf = (tenant) => make(tenant.spHttpHandler());

// ---------------------------------------------------------------- SPClient
test('SPClient: GET returns json; identical GETs are cached unless cacheTtlMs is 0', async () => {
  let calls = 0;
  const c = make(async () => json({ n: ++calls }));
  assert.deepStrictEqual(await c.get(ORIGIN + '/_api/a'), { n: 1 });
  assert.deepStrictEqual(await c.get(ORIGIN + '/_api/a'), { n: 1 });
  assert.deepStrictEqual(await c.get(ORIGIN + '/_api/a', { cacheTtlMs: 0 }), { n: 2 });
});

test('SPClient: refuses any other origin', async () => {
  const c = make(async () => json({}));
  await assert.rejects(() => c.get('https://evil.example.com/_api/web'), /different origin/);
  await assert.rejects(() => c.post('https://contoso.sharepoint.com.evil.example/_api/web', {}), /different origin/);
});

test('SPClient: a 429 is retried after Retry-After and counted', async () => {
  let calls = 0;
  const c = make(async () => (++calls === 1 ? json({}, 429, { 'Retry-After': '0' }) : json({ ok: true })));
  assert.deepStrictEqual(await c.get(ORIGIN + '/_api/a'), { ok: true });
  assert.strictEqual(c.stats.throttled, 1);
});

test('SPClient: errors carry the status and the SharePoint message; raw bodies are sent as text', async () => {
  const sent = [];
  const c = make(async (m, url, cfg, opts) => {
    sent.push(opts);
    return m === 'GET' ? json({ 'odata.error': { message: { value: 'Nope' } } }, 403) : json({});
  });
  await assert.rejects(
    () => c.get(ORIGIN + '/_api/a'),
    (e) => e instanceof HttpError && e.status === 403 && /Nope/.test(e.message)
  );
  await c.post(ORIGIN + '/_api/b', 'a,b\r\n', { raw: true });
  assert.strictEqual(sent[1].body, 'a,b\r\n');
  await c.post(ORIGIN + '/_api/c', { x: 1 }, { method: 'MERGE' });
  assert.strictEqual(sent[2].body, '{"x":1}');
  assert.strictEqual(sent[2].headers['X-HTTP-Method'], 'MERGE');
  assert.strictEqual(sent[2].headers['IF-MATCH'], '*');
});

test('SPClient: request starts are spaced by the gap and never more than two run at once', async () => {
  let active = 0;
  let peak = 0;
  const starts = [];
  const c = make(async () => {
    starts.push(Date.now());
    peak = Math.max(peak, ++active);
    await new Promise((r) => setTimeout(r, 40));
    active--;
    return json({});
  }, 30);
  await Promise.all([1, 2, 3, 4, 5].map((i) => c.get(`${ORIGIN}/_api/${i}`)));
  assert.ok(peak <= 2, `peak ${peak}`);
  for (let i = 1; i < starts.length; i++) assert.ok(starts[i] - starts[i - 1] >= 25, `gap ${starts[i] - starts[i - 1]}`);
});

// ---------------------------------------------------------------- permissions
test('permissions: the built-in levels map to the level their name says', () => {
  const lv = (m) => P.permissionLevel(Number(m.Low), Number(m.High));
  assert.strictEqual(lv(FULL), 'Full control');
  assert.strictEqual(lv(DESIGN), 'Design / manage');
  assert.strictEqual(lv(EDIT), 'Edit / contribute');
  assert.strictEqual(lv({ Low: '1011028719', High: '432' }), 'Edit / contribute', 'Contribute');
  assert.strictEqual(lv(READ), 'Read');
  assert.strictEqual(lv({ Low: '138612801', High: '176' }), 'Read', 'View Only');
  assert.strictEqual(lv(LIMITED), 'No access', 'Limited Access');
  assert.strictEqual(lv({ Low: '0', High: '0' }), 'No access');
  assert.strictEqual(lv({ Low: '4294967295', High: '2147483646' }), 'Design / manage', 'one bit short of full control');
  assert.strictEqual(P.permissionLevel(33554432, 0), 'Design / manage', 'ManagePermissions alone');
});

test('permissions: logins, kinds and group claims', () => {
  assert.strictEqual(P.toLogin(' Ann@Contoso.com '), 'i:0#.f|membership|ann@contoso.com');
  assert.strictEqual(P.toLogin('c:0t.c|tenant|abc'), 'c:0t.c|tenant|abc');
  assert.ok(P.isPersonLogin('i:0#.f|membership|a@b.c') && !P.isPersonLogin('c:0t.c|tenant|x'));
  const m = P.m365GroupLogin('ABCDEF00-0000-0000-0000-000000000000');
  assert.strictEqual(m, 'c:0o.c|federateddirectoryclaimprovider|abcdef00-0000-0000-0000-000000000000');
  assert.strictEqual(P.normaliseLogin(m + '_o'), m, 'the owners claim is stored as the group');
  assert.deepStrictEqual(P.groupClaims(m), [m, m + '_o']);
  assert.deepStrictEqual(P.groupClaims('c:0t.c|tenant|abc'), ['c:0t.c|tenant|abc']);
  assert.strictEqual(P.normaliseLogin('c:0t.c|tenant|0000000_o'), 'c:0t.c|tenant|0000000_o', 'only Microsoft 365 claims lose _o');
  assert.strictEqual(P.subjectKind('i:0#.f|membership|x_y.com#ext#@c.onmicrosoft.com'), 'Guest');
  assert.strictEqual(P.subjectKind('i:0#.f|membership|a@b.c'), 'Person');
  assert.strictEqual(P.subjectKind(m), 'Microsoft 365 group');
  assert.strictEqual(P.subjectKind(P.securityGroupLogin('1-2')), 'Security group');
  assert.strictEqual(P.subjectKind(P.everyoneExceptExternalLogin('T-1')), 'Everyone claim');
  assert.strictEqual(P.subjectKind(P.EVERYONE_LOGIN), 'Everyone claim');
});

test('permissions: adding by hand validates what was typed', () => {
  assert.deepStrictEqual(P.manualSubject('person', ' Ann@Contoso.com ', ''), { login: 'i:0#.f|membership|ann@contoso.com', label: 'Ann@Contoso.com' });
  assert.strictEqual(typeof P.manualSubject('person', 'ann', ''), 'string');
  assert.strictEqual(typeof P.manualSubject('security', 'not-a-guid', ''), 'string');
  assert.strictEqual(typeof P.manualSubject('claim', 'nopipe', ''), 'string');
  assert.strictEqual(typeof P.manualSubject('person', '  ', ''), 'string');
  assert.deepStrictEqual(P.manualSubject('m365', '11111111-1111-1111-1111-11111111111A', 'Team'), { login: 'c:0o.c|federateddirectoryclaimprovider|11111111-1111-1111-1111-11111111111a', label: 'Team' });
  assert.strictEqual(P.manualSubject('security', '11111111-1111-1111-1111-111111111111', '').login, 'c:0t.c|tenant|11111111-1111-1111-1111-111111111111');
});

// ---------------------------------------------------------------- report files
test('report CSV: every cell quoted, formulas defused, BOM and CRLF; reading it back gives the values', () => {
  const rows = [['=1+1', 'a "b"', 'x,y', "O'Brien", '', '-2', '@x', '+3', 'line\nbreak']];
  const text = C.toReportCsv(['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'], rows);
  assert.ok(text.charCodeAt(0) === 0xfeff);
  assert.strictEqual(text.substring(1), '"A","B","C","D","E","F","G","H","I"\r\n"\'=1+1","a ""b""","x,y","O\'Brien","","\'-2","\'@x","\'+3","line\nbreak"\r\n');
  assert.deepStrictEqual(C.parseReport(text), rows);
  assert.deepStrictEqual(C.parseCsv('a,b\n"c,d",e\r\n\r\n'), [['a', 'b'], ['c,d', 'e']]);
  assert.deepStrictEqual(C.parseReport('"h"\r\n"\'plain"\r\n'), [["'plain"]], 'an apostrophe that is not a guard stays');
});

test('report file names carry the run time in UTC and sort by it', () => {
  const d = new Date(Date.UTC(2026, 10, 5, 3, 7));
  assert.strictEqual(C.fileStamp(d), '2026-11-05-0307');
  assert.strictEqual(C.reportFileName(C.fileStamp(d)), 'access-report-2026-11-05-0307.csv');
  assert.deepStrictEqual(C.stampOf('access-report-2026-11-05-0307.csv'), { stamp: '2026-11-05-0307', unread: false, at: d });
  assert.strictEqual(C.stampOf('Access-Report-2026-11-05-0307-unread.csv').unread, true);
  assert.strictEqual(C.stampOf('notes.csv'), undefined);
  assert.strictEqual(C.stampOf('access-report-2026-11-05-0307.csv.bak'), undefined);
});

// ---------------------------------------------------------------- scanner
test('scanner: grants, role names and masks in every OData flavour', () => {
  const a = [
    { PrincipalId: 5, Member: { Title: 'Members', PrincipalType: 8 }, RoleDefinitionBindings: { results: [{ Name: 'Edit', RoleTypeKind: 6 }, { Name: 'Limited Access', RoleTypeKind: 1 }] } },
    { PrincipalId: 9, Member: { Title: 'G', PrincipalType: 4 }, RoleDefinitionBindings: [{ Name: 'Read', RoleTypeKind: 2 }] },
    { PrincipalId: 9, RoleDefinitionBindings: [{ Name: 'limited access (web only)' }] },
    { PrincipalId: 7, Member: { Title: 'Other', PrincipalType: 8 }, RoleDefinitionBindings: [{ Name: 'Full Control', RoleTypeKind: 5 }] }
  ];
  assert.deepStrictEqual(grantsFor(a, 9, [5]), [
    { level: 'Edit', how: 'SharePoint group: Members' },
    { level: 'Read', how: 'Direct' }
  ]);
  assert.deepStrictEqual(grantsFor(a, 1, []), []);
  assert.deepStrictEqual(roleNames(a[0]), ['Edit']);
  assert.strictEqual(levelOf({ Low: '138612833', High: '176' }), 'Read');
  assert.strictEqual(levelOf({ d: { GetUserEffectivePermissions: { Low: '4294967295', High: '2147483647' } } }), 'Full control');
  assert.strictEqual(levelOf({ value: { Low: 1011030767, High: 432 } }), 'Edit / contribute');
  assert.strictEqual(levelOf(undefined), 'No access');
});

test('scanner: a cancelled request stops the site instead of being recorded as unread', async () => {
  const c = make(async () => {
    throw new CancelledError();
  });
  await assert.rejects(() => new AccessScanner(c, ORIGIN + '/sites/it').scanSite({ url: ORIGIN + '/sites/a', title: 'A' }, [{ label: 'Ann', login: scenarios.ANN }], true), (e) => e instanceof CancelledError);
});

test('scanner: one request per person and site, one per group where it is unknown; libraries listed once', async () => {
  const t = scenarios.rich();
  const scanner = new AccessScanner(clientOf(t), t.home);
  const subjects = [
    { label: 'Ann', login: scenarios.ANN },
    { label: 'Auditors', login: scenarios.AUDITORS }
  ];
  const r = await scanner.scanSite({ url: ORIGIN + '/sites/legal/', title: 'Legal' }, subjects, false);
  assert.strictEqual(t.log.length, 2);
  assert.deepStrictEqual(r.rows, [['Ann', ORIGIN + '/sites/legal', 'Legal', 'Site', 'Legal', ORIGIN + '/sites/legal', 'Design / manage', 'Effective permission', '']]);
  t.log.length = 0;
  await scanner.scanSite({ url: ORIGIN + '/sites/it', title: 'IT' }, subjects, true);
  assert.strictEqual(t.log.filter((l) => /_api\/web\/lists\?/.test(l.url)).length, 1);
});

// ---------------------------------------------------------------- sites
test('sites: scope rules', () => {
  const o = ORIGIN;
  assert.ok(inScope(o + '/sites/a', 'A', 'GROUP', o, ''));
  assert.ok(inScope(o, 'Root', 'SITEPAGEPUBLISHING', o, ''));
  assert.ok(!inScope(o + 'x/sites/a', 'A', '', o, ''), 'a host that merely starts the same');
  assert.ok(!inScope('https://contoso-my.sharepoint.com/personal/a', 'A', '', o, ''));
  assert.ok(!inScope(o + '/personal/a', 'A', '', o, ''));
  assert.ok(!inScope(o + '/sites/a', 'A', 'SPSPERS#10', o, ''));
  assert.ok(!inScope('', 'A', '', o, ''));
  assert.ok(inScope(o + '/sites/a', 'Finance', '', o, ' fin '));
  assert.ok(inScope(o + '/sites/FIN', 'x', '', o, 'fin'));
  assert.ok(!inScope(o + '/sites/a', 'A', '', o, 'fin'));
  assert.ok(isOneDriveSite(o + '/personal/x') && !isOneDriveSite(o + '/sites/personal'));
});

test('sites: plain query walked by startrow in DocId order; a refused search is an error', async () => {
  const t = scenarios.large(1203);
  const r = await new SiteLister(clientOf(t)).list(t.home, '');
  assert.strictEqual(r.sites.length, 1203);
  assert.strictEqual(new Set(r.sites.map((s) => s.url)).size, 1203);
  const urls = t.log.map((l) => decodeURIComponent(l.url));
  assert.strictEqual(urls.length, 3);
  assert.ok(urls.every((u) => /querytext='contentclass:STS_Site'&/.test(u) && /sortlist='\[DocId\]:ascending'/.test(u) && !/IndexDocId/.test(u)), urls[0]);
  assert.deepStrictEqual(urls.map((u) => /startrow=(\d+)/.exec(u)[1]), ['0', '500', '1000']);
  t.search = 'reject';
  await assert.rejects(() => new SiteLister(clientOf(t)).list(t.home, ''), (e) => e instanceof HttpError && e.status === 400);
});

test('sites: duplicates are listed once, titles fall back to the address, progress is reported', async () => {
  const t = new FakeTenant({ sites: [{ url: ORIGIN + '/sites/a', title: 'A' }, { url: ORIGIN + '/sites/A/', title: 'A again' }, { url: ORIGIN + '/sites/b' }] }).provisioned();
  const seen = [];
  const r = await new SiteLister(clientOf(t)).list(t.home, '', (n) => seen.push(n));
  assert.deepStrictEqual(r.sites, [{ url: ORIGIN + '/sites/a', title: 'A' }, { url: ORIGIN + '/sites/b', title: ORIGIN + '/sites/b' }]);
  assert.deepStrictEqual(seen, [2]);
});

// ---------------------------------------------------------------- store
test('store: set-up creates the lists, columns, library and settings item once', async () => {
  const t = new FakeTenant();
  const s = new Store(clientOf(t), t.home + '/');
  assert.deepStrictEqual(await s.missing(), ['Access Report Subjects', 'Access Report Settings', 'Access Reports']);
  const steps = [];
  await s.provision((x) => steps.push(x));
  assert.strictEqual(steps.length, 3);
  assert.deepStrictEqual(await s.missing(), []);
  assert.deepStrictEqual(Object.keys(t.lists).sort(), ['/sites/it/AccessReports', '/sites/it/Lists/AccessReportSettings', '/sites/it/Lists/AccessReportSubjects']);
  assert.strictEqual(t.lists['/sites/it/Lists/AccessReportSubjects'].Title, 'Access Report Subjects');
  assert.deepStrictEqual(t.lists['/sites/it/Lists/AccessReportSubjects'].fields, ['Title', 'ARLogin', 'ARActive']);
  assert.deepStrictEqual(t.lists['/sites/it/Lists/AccessReportSettings'].fields, ['Title', 'ARIncludeLibraries', 'ARSiteFilter', 'ARReportFolder', 'ARLastRun']);
  assert.deepStrictEqual(await s.settings(), { id: 1, includeLibraries: true, siteFilter: '', reportFolder: '/sites/it/AccessReports', lastRun: '' });
  const before = t.log.filter((l) => l.method === 'POST').length;
  await s.provision();
  assert.strictEqual(t.log.filter((l) => l.method === 'POST').length, before, 'a second set-up changes nothing');
  assert.strictEqual(t.lists['/sites/it/Lists/AccessReportSettings'].items.length, 1);
  assert.ok(await s.folderExists('/sites/it/AccessReports'));
  assert.ok(!(await s.folderExists('/sites/it/Nope')));
});

test('store: a half-finished set-up is completed (missing column added)', async () => {
  const t = new FakeTenant().provisioned();
  t.lists['/sites/it/Lists/AccessReportSettings'].fields = ['Title', 'ARIncludeLibraries'];
  const s = new Store(clientOf(t), t.home);
  await s.provision();
  assert.deepStrictEqual(t.lists['/sites/it/Lists/AccessReportSettings'].fields, ['Title', 'ARIncludeLibraries', 'ARSiteFilter', 'ARReportFolder', 'ARLastRun']);
});

test('store: subjects and settings round-trip; root site paths have no double slash', async () => {
  const t = new FakeTenant({ home: ORIGIN }).provisioned();
  const s = new Store(clientOf(t), ORIGIN);
  assert.strictEqual(s.webPath, '');
  assert.strictEqual(s.defaultFolder, '/AccessReports');
  await s.addSubject({ label: '  Team  ', login: 'c:0o.c|federateddirectoryclaimprovider|abc_o' });
  await s.addSubject({ label: 'Ann', login: 'Ann@Contoso.com' });
  let list = await s.subjects();
  assert.deepStrictEqual(list, [
    { id: 1, label: 'Team', login: 'c:0o.c|federateddirectoryclaimprovider|abc', active: true },
    { id: 2, label: 'Ann', login: 'i:0#.f|membership|ann@contoso.com', active: true }
  ]);
  await s.setSubjectActive(1, false);
  await s.removeSubject(2);
  list = await s.subjects();
  assert.deepStrictEqual(list, [{ id: 1, label: 'Team', login: 'c:0o.c|federateddirectoryclaimprovider|abc', active: false }]);
  const settings = await s.settings();
  await s.saveSettings(Object.assign({}, settings, { includeLibraries: false, siteFilter: ' fin ', reportFolder: '/Shared Documents/Reports/' }));
  assert.deepStrictEqual(await s.settings(), { id: 1, includeLibraries: false, siteFilter: 'fin', reportFolder: '/Shared Documents/Reports', lastRun: '' });
});

test('store: report files are saved with a BOM, listed newest first, and read back', async () => {
  const t = new FakeTenant().provisioned();
  const s = new Store(clientOf(t), t.home);
  const f = '/sites/it/AccessReports';
  assert.strictEqual(await s.saveFile(f, 'access-report-2026-10-01-0300.csv', '"Who"\r\n'), f + '/access-report-2026-10-01-0300.csv');
  await s.saveFile(f, 'access-report-2026-11-01-0300.csv', C.BOM + '"Who"\r\n"Ann"\r\n');
  await s.saveFile(f, 'access-report-2026-11-01-0300-unread.csv', 'x');
  await s.saveFile(f, "it's notes.txt", 'x');
  assert.strictEqual(t.files[f + '/access-report-2026-10-01-0300.csv'], C.BOM + '"Who"\r\n', 'BOM added once');
  assert.strictEqual(t.files[f + '/access-report-2026-11-01-0300.csv'], C.BOM + '"Who"\r\n"Ann"\r\n');
  const list = await s.reports(f);
  assert.deepStrictEqual(list.map((r) => [r.name, r.unread]), [
    ['access-report-2026-11-01-0300.csv', false],
    ['access-report-2026-11-01-0300-unread.csv', true],
    ['access-report-2026-10-01-0300.csv', false]
  ]);
  assert.strictEqual(list[0].modifiedBy, 'Flow Owner');
  assert.deepStrictEqual(C.parseReport(await s.readFile(list[0].serverRelativeUrl)), [['Ann']]);
  await assert.rejects(() => s.saveFile('/sites/it/Missing', 'a.csv', 'x'), (e) => e instanceof HttpError && e.status === 404);
});

test('store: folder paths the file API cannot address are refused up front', () => {
  assert.strictEqual(folderProblem('/sites/it/AccessReports'), undefined);
  assert.strictEqual(folderProblem(''), undefined);
  assert.ok(/Start with/.test(folderProblem('AccessReports')));
  ['/a#b', '/a%20b', '/a?b', '/a*b', '/a"b', '/a\\b', '/a:b'].forEach((p) => assert.ok(folderProblem(p), p));
  assert.ok(/must be in this site/.test(folderProblem('/sites/other/Reports', '/sites/it')));
  assert.ok(/must be in this site/.test(folderProblem('/sites/italy/Reports', '/sites/it')));
  assert.strictEqual(folderProblem('/sites/IT/Reports', '/sites/it'), undefined);
  assert.strictEqual(folderProblem('/Shared Documents', ''), undefined);
  assert.strictEqual(aliasValue("/sites/it/O'Neil & Co"), "'%2Fsites%2Fit%2FO''Neil%20%26%20Co'");
});

// ---------------------------------------------------------------- people picker
test('people picker: users and groups become subjects; SharePoint groups and unresolved entries do not', async () => {
  const t = new FakeTenant().provisioned();
  t.picker = [
    { Key: 'i:0#.f|membership|ann@contoso.com', DisplayText: 'Ann Example', EntityType: 'User', EntityData: { Email: 'ann@contoso.com' } },
    { Key: 'c:0o.c|federateddirectoryclaimprovider|abc_o', DisplayText: 'Marketing Owners', EntityType: 'FormsRole', EntityData: {} },
    { Key: 'c:0o.c|federateddirectoryclaimprovider|abc', DisplayText: 'Marketing', EntityType: 'FormsRole', EntityData: { Email: 'marketing@contoso.com' } },
    { Key: 'c:0t.c|tenant|def', DisplayText: 'Auditors <auditors@contoso.com>', EntityType: 'SecGroup', EntityData: { Email: 'auditors@contoso.com' } },
    { Key: 'IT Members', DisplayText: 'IT Members', EntityType: 'SPGroup' },
    { Key: 'someone@nowhere.example', DisplayText: 'someone@nowhere.example', EntityType: '', IsResolved: false }
  ];
  const found = await new PeoplePicker(clientOf(t), t.home).search(' ann ');
  assert.deepStrictEqual(found, [
    { login: 'i:0#.f|membership|ann@contoso.com', label: 'Ann Example <ann@contoso.com>', kind: 'Person' },
    { login: 'c:0o.c|federateddirectoryclaimprovider|abc', label: 'Marketing <marketing@contoso.com>', kind: 'Microsoft 365 group' },
    { login: 'c:0t.c|tenant|def', label: 'Auditors <auditors@contoso.com>', kind: 'Security group' }
  ]);
  assert.deepStrictEqual(fromPicker([{ Key: 'c:0o.c|federateddirectoryclaimprovider|abc_o', DisplayText: 'Marketing Owners', EntityType: 'FormsRole' }]), [
    { login: 'c:0o.c|federateddirectoryclaimprovider|abc', label: 'Marketing', kind: 'Microsoft 365 group' }
  ]);
  assert.deepStrictEqual(await new PeoplePicker(clientOf(t), t.home).search('   '), []);
  assert.deepStrictEqual(fromPicker([]), []);
});

// ---------------------------------------------------------------- run job
const smallSites = (n, fail) => {
  const sites = [];
  for (let i = 0; i < n; i++) sites.push(Object.assign({ url: `${ORIGIN}/sites/s${i}`, title: `S${i}`, users: { [scenarios.ANN]: READ } }, fail && fail(i)));
  return sites;
};
const jobFor = (t, speed = 'gentle') => {
  const sites = t.sites.map((s) => ({ url: s.url, title: s.title }));
  return { sites, job: R.newJob([{ label: 'Ann', login: scenarios.ANN }], false, '', '/sites/it/AccessReports', speed, sites.length) };
};

test('run job: stop, save, reload and resume gives the same file as an uninterrupted run', async () => {
  const whole = new FakeTenant({ sites: smallSites(12) }).provisioned();
  const a = jobFor(whole);
  await R.runJob(a.job, a.sites, new AccessScanner(clientOf(whole), whole.home), { stop: false }, () => undefined, 0);
  assert.ok(a.job.finished);

  const t = new FakeTenant({ sites: smallSites(12) }).provisioned();
  const b = jobFor(t);
  R.saveJob(b.job, b.sites);
  const token = { stop: false };
  let n = 0;
  await R.runJob(b.job, b.sites, new AccessScanner(clientOf(t), t.home), token, () => (++n === 5 ? (token.stop = true) : 0), 0);
  assert.strictEqual(b.job.finished, false);
  assert.strictEqual(R.processed(b.job), 5);
  const back = R.loadJob();
  assert.deepStrictEqual(back.sites, b.sites);
  assert.deepStrictEqual(R.remaining(back.job), [5, 6, 7, 8, 9, 10, 11]);
  await R.runJob(back.job, back.sites, new AccessScanner(clientOf(t), t.home), { stop: false }, () => undefined, 0);
  assert.ok(back.job.finished);
  assert.strictEqual(R.reportCsv(back.job), R.reportCsv(a.job));
  assert.strictEqual(t.log.length, 12, 'no site was scanned twice');
  R.clearJob();
  assert.strictEqual(R.loadJob(), undefined);
});

test('run job: two workers finish out of order, the file is still in site order', async () => {
  const t = new FakeTenant({ sites: smallSites(9) }).provisioned();
  const { sites, job } = jobFor(t, 'normal');
  const handler = t.spHttpHandler();
  const c = make(async (...args) => {
    await new Promise((r) => setTimeout(r, /s[0246]\//.test(args[1]) ? 15 : 1));
    return handler(...args);
  });
  await R.runJob(job, sites, new AccessScanner(c, t.home), { stop: false }, () => undefined, 0);
  assert.notDeepStrictEqual(job.rows.map((r) => r[0]), [0, 1, 2, 3, 4, 5, 6, 7, 8], 'finished out of order');
  assert.deepStrictEqual(C.parseReport(R.reportCsv(job)).map((r) => r[2]), ['S0', 'S1', 'S2', 'S3', 'S4', 'S5', 'S6', 'S7', 'S8']);
});

test('run job: eight sites in a row failing for another reason than "no access" pause the run', async () => {
  const t = new FakeTenant({ sites: smallSites(20, (i) => (i >= 3 ? { broken: { status: 500, body: 'down', raw: 'down' } } : i === 1 ? { deny: 403 } : {})) }).provisioned();
  const { sites, job } = jobFor(t);
  const token = { stop: false };
  await R.runJob(job, sites, new AccessScanner(clientOf(t), t.home), token, () => undefined, 0);
  assert.ok(/Paused after 8 sites/.test(token.reason), token.reason);
  assert.strictEqual(job.finished, false);
  assert.strictEqual(R.processed(job), 11);
  assert.deepStrictEqual(R.unreadSites(job), [1, 3, 4, 5, 6, 7, 8, 9, 10]);
  // Twenty denied sites in a row are about the sites, not the connection: no pause.
  const denied = new FakeTenant({ sites: smallSites(20, () => ({ deny: 403 })) }).provisioned();
  const d = jobFor(denied);
  const tk = { stop: false };
  await R.runJob(d.job, d.sites, new AccessScanner(clientOf(denied), denied.home), tk, () => undefined, 0);
  assert.ok(d.job.finished && !tk.reason);
});

test('run job: retrying unread sites replaces what they produced and keeps the rest', async () => {
  const t = new FakeTenant({ sites: smallSites(4, (i) => (i === 2 ? { deny: 403 } : {})) }).provisioned();
  const { sites, job } = jobFor(t);
  const scanner = new AccessScanner(clientOf(t), t.home);
  await R.runJob(job, sites, scanner, { stop: false }, () => undefined, 0);
  assert.deepStrictEqual(R.unreadSites(job), [2]);
  assert.strictEqual(job.rows.length, 3);
  delete t.sites[2].deny;
  job.savedAs = '/x';
  R.requeueUnread(job);
  assert.deepStrictEqual(R.remaining(job), [2]);
  assert.strictEqual(job.savedAs, undefined);
  await R.runJob(job, sites, scanner, { stop: false }, () => undefined, 0);
  assert.ok(job.finished);
  assert.deepStrictEqual(R.unreadSites(job), []);
  assert.deepStrictEqual(C.parseReport(R.reportCsv(job)).map((r) => r[2]), ['S0', 'S1', 'S2', 'S3']);
  assert.deepStrictEqual(C.parseCsv(R.unreadCsv(job)), [C.UNREAD_HEADERS]);
});

test('run job: a saved job of another version or with a mismatched site list is ignored', () => {
  const j = R.newJob([], false, '', '/f', 'gentle', 2);
  R.saveJob(j, [{ url: 'a', title: 'A' }]);
  assert.strictEqual(R.loadJob(), undefined, 'site count differs');
  R.saveJob(j, [{ url: 'a', title: 'A' }, { url: 'b', title: 'B' }]);
  assert.ok(R.loadJob());
  local['accessReport.job.v2'] = JSON.stringify(Object.assign({}, j, { v: 1 }));
  assert.strictEqual(R.loadJob(), undefined);
  local['accessReport.job.v2'] = '{broken';
  assert.strictEqual(R.loadJob(), undefined);
  R.clearJob();
});

// ---------------------------------------------------------------- compare
test('compare: new, removed and changed access between two reports', () => {
  const row = (who, site, object, name, level, how) => [who, 'https://x/' + site, site, object, name, object === 'Site' ? 'https://x/' + site : 'https://x/' + site + '/' + name, level, how, ''];
  const earlier = [row('Ann', 'a', 'Site', 'a', 'Read', 'Effective permission'), row('Ann', 'b', 'Site', 'b', 'Read', 'Effective permission'), row('G', 'a', 'Site', 'a', 'Edit', 'Direct'), row('G', 'a', 'Site', 'a', 'Read', 'Direct')];
  const current = [row('ANN', 'a', 'Site', 'a', 'Full control', 'Effective permission'), row('Ann', 'c', 'Library', 'Docs', 'Read', 'Inherits from site'), row('G', 'a', 'Site', 'a', 'Read', 'Direct'), row('G', 'a', 'Site', 'a', 'Edit', 'Direct')];
  const changes = compareReports(earlier, current);
  assert.deepStrictEqual(changes.map((c) => [c.kind, c.row[0], c.row[2], c.was]), [
    ['Added', 'Ann', 'c', undefined],
    ['Changed', 'ANN', 'a', 'Read'],
    ['Removed', 'Ann', 'b', undefined]
  ]);
  assert.deepStrictEqual(changeRows(changes)[1].slice(0, 2).concat(changeRows(changes)[1].slice(7, 9)), ['Changed', 'ANN', 'Full control', 'Read']);
  assert.deepStrictEqual(compareReports(earlier, earlier), []);
  assert.strictEqual(compareReports([], current).length, 4);
});

// ---------------------------------------------------------------- zip
test('zip: CRC-32 and UTF-8 are right', () => {
  assert.strictEqual(crc32(utf8('123456789')), 0xcbf43926);
  assert.deepStrictEqual(Array.from(utf8('é–日')), [0xc3, 0xa9, 0xe2, 0x80, 0x93, 0xe6, 0x97, 0xa5]);
  const z = zip([{ name: 'a/b.txt', text: 'hello' }]);
  assert.deepStrictEqual(Array.from(z.slice(0, 4)), [0x50, 0x4b, 0x03, 0x04]);
  assert.deepStrictEqual(Array.from(z.slice(z.length - 22, z.length - 18)), [0x50, 0x4b, 0x05, 0x06]);
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
      console.log(`  FAIL ${name}\n       ${e && e.stack ? e.stack.split('\n').slice(0, 8).join('\n       ') : e}`);
    }
  }
  console.log(`\n${tests.length - failed}/${tests.length} passed${process.env.TEST_TARGET ? ` (${process.env.TEST_TARGET})` : ''}`);
  process.exit(failed ? 1 : 0);
})();
