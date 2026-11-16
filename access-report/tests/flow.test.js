// The Power Automate flow, run for real in the workflow-language simulator against a fake
// tenant, and compared with what the web part produces from the same tenant.
require('./register');
const assert = require('assert');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { execFileSync } = require('child_process');
const src = (p) => require(path.join(__dirname, '..', 'src', 'webparts', 'accessReport', p));

const { runFlow, evalStandalone, namesIn } = require('./wdl');
const { FakeTenant, ORIGIN } = require('./fakeTenant');
const scenarios = require('./scenarios');
const { SPHttpClient } = require('./mocks/sp-http');

const { SPClient } = src('core/SPClient');
const { buildDefinition, levelExpression, cell, PLACEHOLDER_URL } = src('flow/definition');
const pkg = src('flow/package');
const { permissionLevel } = src('services/permissions');
const { reportCell, parseCsv, BOM, REPORT_HEADERS, UNREAD_HEADERS } = src('services/reportCsv');
const { AccessScanner } = src('services/AccessScan');
const { SiteLister } = src('services/Sites');
const { Store } = src('services/Store');
const RunJob = src('services/RunJob');

const store = {};
console.error = () => undefined; // SPClient logs every failed request; the scenarios fail on purpose
global.window = { localStorage: { getItem: (k) => (k in store ? store[k] : null), setItem: (k, v) => (store[k] = String(v)), removeItem: (k) => delete store[k] } };

const tests = [];
const test = (name, fn) => tests.push([name, fn]);
const NOW = '2026-12-01T03:04:05Z';
const STAMP = '2026-12-01-0304';

/** What the web part's "Run now" writes for this tenant. */
async function viaWebPart(tenant) {
  const client = new SPClient(new SPHttpClient(tenant.spHttpHandler()), ORIGIN, 0);
  const st = new Store(client, tenant.home);
  const settings = await st.settings();
  const subjects = (await st.subjects()).filter((s) => s.active && s.login);
  const { sites } = await new SiteLister(client).list(tenant.home, settings.siteFilter);
  const job = RunJob.newJob(subjects, settings.includeLibraries, settings.siteFilter, settings.reportFolder, 'normal', sites.length);
  await RunJob.runJob(job, sites, new AccessScanner(client, tenant.home), { stop: false }, () => undefined, 0);
  assert.strictEqual(job.finished, true);
  return { report: RunJob.reportCsv(job), unread: RunJob.unreadCsv(job), sites, job, folder: settings.reportFolder };
}

/** What the flow writes for this tenant. */
async function viaFlow(tenant, homeUrl) {
  const definition = buildDefinition({ homeUrl: homeUrl || tenant.home });
  const run = await runFlow(definition, { connector: tenant.connector(), now: () => NOW });
  return { run, definition };
}

function sameUnread(flowText, webText) {
  const strip = (text) => parseCsv(text).map((r) => r.slice(0, 5));
  assert.deepStrictEqual(strip(flowText), strip(webText));
  assert.ok(parseCsv(flowText).every((r) => r.length === 6), 'every unread row has six cells');
}

async function parity(make, { homeUrl, expectUnread = true } = {}) {
  const web = await viaWebPart(make());
  const tenant = make();
  const { run } = await viaFlow(tenant, homeUrl);
  assert.deepStrictEqual(run.errors, [], 'no expression or action errors in the flow');
  assert.strictEqual(run.status, 'Succeeded');
  const report = tenant.files[`${web.folder}/access-report-${STAMP}.csv`];
  assert.ok(report !== undefined, `the flow saved its report in ${web.folder} (files: ${Object.keys(tenant.files).join(', ')})`);
  assert.strictEqual(report, web.report, 'the flow and the web part write the same file');
  const unread = tenant.files[`${web.folder}/access-report-${STAMP}-unread.csv`];
  if (expectUnread) {
    assert.ok(unread !== undefined, 'the flow saved the unread list');
    sameUnread(unread, web.unread);
  } else {
    assert.strictEqual(unread, undefined, 'nothing unread, so no unread file');
    assert.strictEqual(parseCsv(web.unread).length, 1);
  }
  return { web, tenant, run, report, unread };
}

// ---------------------------------------------------------------- building blocks
test('level expression equals permissionLevel() on the built-in masks and 20,000 random ones', () => {
  const expr = levelExpression('item()?[\'low\']', 'item()?[\'high\']');
  assert.ok(expr.length < 8192, `level expression is ${expr.length} characters`);
  const check = (low, high) => {
    const got = evalStandalone(expr, { low: String(low), high: String(high) });
    const want = permissionLevel(low, high);
    assert.strictEqual(got === '\u001f' ? 'No access' : got, want, `Low ${low} High ${high}`);
  };
  [0, 1, 2, 14, 15, 262144, 262145, 33554432, 33554433, 134287360, 138612833, 138612801, 1011028719, 1011030767, 1012866047, 999999999, 1000000000, 1000000001, 2147483647, 2147483648, 3999999999, 4000000000, 4294967294, 4294967295].forEach((low) =>
    [0, 48, 176, 432, 2147483646, 2147483647].forEach((high) => check(low, high))
  );
  let seed = 12345;
  const rnd = () => (seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648;
  for (let i = 0; i < 20000; i++) {
    check(Math.floor(rnd() * 4294967296), Math.floor(rnd() * 2147483648));
  }
  assert.strictEqual(evalStandalone(expr, { low: '', high: '' }), '', 'a row without a mask keeps its text');
});

test('cell expression writes exactly what reportCell() writes', () => {
  const expr = cell('item()');
  ['', 'plain', 'a,b', 'say "hi"', '=SUM(1)', '+1', '-1', '@x', '\tTab', '\rCr', "O'Brien", 'Résumé – 日本語', ' =lead', '"', "'=x"].forEach((v) => assert.strictEqual(evalStandalone(expr, v), reportCell(v), JSON.stringify(v)));
});

// ---------------------------------------------------------------- structure
test('definition is well-formed: names, loops, references, expression sizes', () => {
  const def = buildDefinition({ homeUrl: PLACEHOLDER_URL });
  const names = namesIn(def.actions);
  assert.strictEqual(new Set(names).size, names.length, 'action names are unique across the flow');
  names.forEach((n) => assert.ok(/^[A-Za-z][A-Za-z0-9_]{0,79}$/.test(n), `action name ${n}`));
  let longest = 0;
  const walk = (actions, loops, depth) => {
    Object.keys(actions).forEach((name) => {
      const a = actions[name];
      assert.ok(a.runAfter && typeof a.runAfter === 'object', `${name} has runAfter`);
      Object.keys(a.runAfter).forEach((p) => {
        assert.ok(p in actions, `${name} runs after ${p} in the same scope`);
        a.runAfter[p].forEach((s) => assert.ok(['Succeeded', 'Failed', 'Skipped', 'TimedOut'].indexOf(s) >= 0));
      });
      if (a.type === 'InitializeVariable') assert.strictEqual(depth, 0, `${name} is at the top level`);
      if (a.type === 'Foreach') assert.strictEqual(a.runtimeConfiguration.concurrency.repetitions, 1, `${name} runs one at a time`);
      if (a.type === 'OpenApiConnection') assert.strictEqual(a.inputs.authentication, "@parameters('$authentication')");
      const text = JSON.stringify(Object.assign({}, a, { actions: undefined, else: undefined }));
      const inner = a.type === 'Foreach' ? loops.concat(name) : loops;
      const strings = [];
      JSON.stringify(Object.assign({}, a, { actions: undefined, else: undefined }), (k, v) => (typeof v === 'string' ? (strings.push(v), v) : v));
      strings.forEach((s) => {
        const exprs = s.charAt(0) === '@' && s.charAt(1) !== '{' ? [s.slice(1)] : (s.match(/@\{(?:[^}']|'[^']*')*\}/g) || []).map((x) => x.slice(2, -1));
        exprs.forEach((e) => {
          longest = Math.max(longest, e.length);
          assert.ok(e.length < 8192, `${name}: an expression of ${e.length} characters`);
          (e.match(/(?:body|outputs)\('([^']+)'\)/g) || []).forEach((ref) => assert.ok(names.indexOf(/'([^']+)'/.exec(ref)[1]) >= 0, `${name} refers to ${ref}`));
          (e.match(/items\('([^']+)'\)/g) || []).forEach((ref) => assert.ok(loops.indexOf(/'([^']+)'/.exec(ref)[1]) >= 0, `${name} uses ${ref} outside that loop`));
        });
      });
      assert.ok(text.length > 0);
      if (a.actions) walk(a.actions, inner, depth + 1);
      if (a.else && a.else.actions) walk(a.else.actions, inner, depth + 1);
    });
  };
  walk(def.actions, [], 0);
  assert.ok(longest > 2000 && longest < 8192, `longest expression: ${longest}`);
  assert.deepStrictEqual(Object.keys(def.triggers), ['On_the_first_of_every_month']);
  assert.deepStrictEqual(def.triggers.On_the_first_of_every_month.recurrence, { frequency: 'Month', interval: 1, startTime: '2026-01-01T03:00:00Z' });
  // Only the standard SharePoint connector is used.
  const apis = new Set();
  JSON.stringify(def, (k, v) => (k === 'apiId' ? (apis.add(v), v) : v));
  assert.deepStrictEqual([...apis], ['/providers/Microsoft.PowerApps/apis/shared_sharepointonline']);
  console.log(`       (${names.length} actions, longest expression ${longest} characters)`);
});

// ---------------------------------------------------------------- parity
test('rich tenant: same report and same unread list as the web part', async () => {
  const { report, unread, tenant, run, web } = await parity(() => scenarios.rich());
  const rows = parseCsv(report);
  assert.deepStrictEqual(rows[0], REPORT_HEADERS);
  assert.deepStrictEqual(parseCsv(unread)[0], UNREAD_HEADERS);
  assert.ok(report.charAt(0) === BOM && /\r\n$/.test(report));
  const has = (cells) => assert.ok(rows.some((r) => cells.every((c, i) => c === undefined || r[i] === c)), `row ${JSON.stringify(cells)} in\n${rows.map((r) => r.join(' | ')).join('\n')}`);
  // People: effective permission on the site, inherited libraries, libraries with their own permissions.
  has(['Ann Example <ann@contoso.com>', ORIGIN + '/sites/it', 'IT', 'Site', 'IT', ORIGIN + '/sites/it', 'Full control', 'Effective permission', '']);
  has(['Ann Example <ann@contoso.com>', undefined, 'IT', 'Library', 'Documents', ORIGIN + '/sites/it/Shared Documents', 'Full control', 'Inherits from site', 'No']);
  has(['Ann Example <ann@contoso.com>', undefined, 'IT', 'Library', 'Secret "Plans", v2', undefined, 'Full control', 'Effective permission (own permissions)', 'Yes']);
  has(['Ann Example <ann@contoso.com>', undefined, "'=Marketing team".slice(0, 0) + 'Finance, =SUM(1)', 'Site', undefined, undefined, 'Read']);
  has(['Ann Example <ann@contoso.com>', ORIGIN + '/sites/legal', 'Legal', 'Site', undefined, undefined, 'Design / manage']);
  has(['Ann Example <ann@contoso.com>', ORIGIN, 'Home', 'Site', 'Home', ORIGIN, 'Edit / contribute']);
  // A guest with nothing on the site but something on one library.
  has(['O\'Brien, "Pat" (guest)', ORIGIN + '/sites/it', 'IT', 'Library', 'Secret "Plans", v2', undefined, 'Read', undefined, 'Yes']);
  has(['O\'Brien, "Pat" (guest)', ORIGIN + '/sites/fin', undefined, 'Library', 'Payroll', undefined, 'Edit / contribute', undefined, 'Yes']);
  assert.ok(!rows.some((r) => r[0].indexOf('Pat') >= 0 && r[1] === ORIGIN + '/sites/fin' && r[3] === 'Site'), 'Limited Access alone is not access');
  // Groups: each grant that reaches the group.
  has(['Finance Auditors', ORIGIN + '/sites/it', 'IT', 'Site', 'IT', undefined, 'Edit, Custom "Approve"', 'SharePoint group: IT Members', '']);
  has(['Finance Auditors', undefined, 'IT', 'Library', 'Documents', undefined, 'Edit, Custom "Approve"', 'Inherits from site - SharePoint group: IT Members', 'No']);
  has(['Finance Auditors', undefined, 'IT', 'Library', 'Secret "Plans", v2', undefined, 'Read, Contribute', 'Direct', 'Yes']);
  has(["'=Marketing team (owners)", ORIGIN + '/sites/it', 'IT', 'Site', 'IT', undefined, 'Full control', 'Site collection administrator', '']);
  has(["'=Marketing team (owners)", undefined, 'IT', 'Site', undefined, undefined, 'Full Control', 'SharePoint group: IT Owners']);
  has(["'=Marketing team (owners)", undefined, 'IT', 'Library', 'Secret "Plans", v2', undefined, 'Design', 'Direct', 'Yes']);
  has(["'=Marketing team", undefined, 'IT', 'Site', undefined, undefined, "'=Odd level", 'Direct']);
  has(['Everyone except external users', ORIGIN + '/sites/it', 'IT', 'Site', undefined, undefined, 'Read', 'Direct']);
  has(['Odd claim', ORIGIN + '/sites/it', 'IT', 'Site', undefined, undefined, 'Read', 'Direct']);
  has(['Finance Auditors', ORIGIN + '/sites/ops', 'Ops', 'Site', undefined, undefined, 'Full control', 'Site collection administrator']);
  // Sites that must not be in it.
  ['/personal/', '/sites/odd', 'contoso-my', 'fabrikam', 'lookalike'].forEach((s) => assert.ok(report.indexOf(s) < 0, `${s} is out of scope`));
  assert.ok(!rows.some((r) => /Paused|No login/.test(r[0])), 'inactive and empty subjects are skipped');
  assert.ok(!rows.some((r) => /Site Assets|Tasks|Hidden library/.test(r[4])), 'system libraries, lists and hidden libraries are skipped');
  // What could not be read.
  const u = parseCsv(unread);
  const unreadHas = (cells) => assert.ok(u.some((r) => cells.every((c, i) => c === undefined || r[i] === c)), `unread ${JSON.stringify(cells)} in\n${u.map((r) => r.join(' | ')).join('\n')}`);
  unreadHas(['(all)', ORIGIN + '/sites/hr', 'HR', 'Libraries', '403']);
  unreadHas(['Ann Example <ann@contoso.com>', ORIGIN + '/sites/hr', 'HR', 'Site', '403']);
  unreadHas(['Finance Auditors', ORIGIN + '/sites/hr', 'HR', 'Site', '403']);
  unreadHas(['(all)', ORIGIN + '/sites/legal', 'Legal', 'Libraries', '500']);
  unreadHas(['Finance Auditors', ORIGIN + '/sites/ops', 'Ops', 'Site permissions', '500']);
  unreadHas(['Ann Example <ann@contoso.com>', ORIGIN + '/sites/fin', undefined, 'Library: Locked "vault"', '404']);
  unreadHas(['Ghost', ORIGIN + '/sites/it', 'IT', 'Site', '404']);
  unreadHas(['Ann Example <ann@contoso.com>', ORIGIN + '/sites/broken', 'Broken', 'Site', '500']);
  unreadHas(['Ann Example <ann@contoso.com>', ORIGIN + '/sites/offline', 'Offline', 'Site', '']);
  const denied = u.filter((r) => r[1] === ORIGIN + '/sites/hr')[0];
  assert.ok(/Access denied/.test(denied[5]) && !/clientRequestId|[\r\n]/.test(denied[5]), `a readable one-line error: ${denied[5]}`);
  // The note the web part shows.
  const note = tenant.settingsItem().ARLastRun;
  assert.strictEqual(note, `Finished 2026-12-01 03:04 UTC: ${web.sites.length} sites, 7 people and groups, ${rows.length - 1} rows, ${u.length - 1} not read. File: access-report-${STAMP}.csv`);
  assert.strictEqual(run.results.Scan_completed.status, 'Succeeded');
  console.log(`       (${web.sites.length} sites x 7 subjects: ${run.executed} flow actions, ${tenant.log.length} requests; ${rows.length - 1} rows, ${u.length - 1} unread)`);
});

test('libraries switched off: same file, and no library request is made', async () => {
  const { tenant, report } = await parity(() => scenarios.rich({ includeLibraries: false }));
  assert.ok(!tenant.log.some((l) => /_api\/web\/lists/.test(l.url)), 'no list request');
  assert.ok(!parseCsv(report).some((r) => r[3] === 'Library'));
});

test('site filter: same file, only matching sites are visited', async () => {
  const { tenant, report } = await parity(() => scenarios.rich({ siteFilter: ' FINANCE ' }), { expectUnread: true });
  const visited = new Set(tenant.log.filter((l) => /getusereffectivepermissions/.test(l.url)).map((l) => l.url.split('/_api/')[0]));
  assert.deepStrictEqual([...visited], [ORIGIN + '/sites/fin']);
  assert.ok(parseCsv(report).length > 1);
});

test('home site is the root site, typed with a trailing slash and spaces; folder with an apostrophe', async () => {
  const folder = "/Shared Documents/Access 'Reports'";
  const make = () => {
    const t = scenarios.rich({ home: ORIGIN, folder });
    t.folders.add(folder.toLowerCase());
    return t;
  };
  const { tenant } = await parity(make, { homeUrl: `  ${ORIGIN}/  ` });
  assert.ok(Object.keys(tenant.files).every((f) => f.indexOf(folder + '/access-report-') === 0), Object.keys(tenant.files).join(', '));
});

test('1,203 sites: three search pages, same file, nothing unread', async () => {
  const { tenant, web, run } = await parity(() => scenarios.large(1203), { expectUnread: false });
  assert.strictEqual(web.sites.length, 1203);
  assert.strictEqual(tenant.log.filter((l) => /search\/query.*rowlimit=500/.test(l.url)).length, 3, "three full pages");
  assert.strictEqual(tenant.log.filter((l) => /search\/query.*rowlimit=1&/.test(l.url)).length, 3, "three one-row probes");
  assert.strictEqual(run.counts.Search_sites, 3);
  assert.strictEqual(run.variables.Sites.value.length, 1203);
  console.log(`       (1,203 sites x 2 subjects: ${run.executed} flow actions = ${(run.executed / 1203 / 2).toFixed(1)} per site and subject)`);
});

test('exactly 500 sites: one more (empty) page ends the listing', async () => {
  const { tenant, web } = await parity(() => scenarios.large(500), { expectUnread: false });
  assert.strictEqual(web.sites.length, 500);
  assert.strictEqual(tenant.log.filter((l) => /search\/query.*rowlimit=500/.test(l.url)).length, 2);
});

test('nobody to report on, and no sites: a header-only file', async () => {
  const none = await parity(() => new FakeTenant({ sites: scenarios.richSites() }).provisioned({ subjects: [] }), { expectUnread: true });
  assert.ok(parseCsv(none.report).length === 1);
  const empty = await parity(() => new FakeTenant({ sites: [] }).provisioned({ subjects: scenarios.SUBJECTS }), { expectUnread: false });
  assert.strictEqual(empty.report, BOM + REPORT_HEADERS.map((h) => `"${h}"`).join(',') + '\r\n');
});

// ---------------------------------------------------------------- when things go wrong
for (const [refused, expect] of [
  [[], 'Created:ascending'],
  [['Created:ascending'], 'LastModifiedTime:ascending'],
  [['Created:ascending', 'LastModifiedTime:ascending'], '']
]) {
  test(`search refuses DocId${refused.map((r) => ' and ' + r).join('')}: the flow and the web part both page by "${expect || 'no order'}" and write the same file`, async () => {
    const make = () => {
      const t = scenarios.large(1203);
      t.refusedSorts = refused;
      return t;
    };
    const { tenant, web, run } = await parity(make, { expectUnread: false });
    assert.strictEqual(web.sites.length, 1203);
    assert.strictEqual(run.variables.Sites.value.length, 1203);
    assert.deepStrictEqual(tenant.sortsUsed, [expect, expect, expect], 'every page used the chosen order');
  });
}

test('the flow and the web part try the same orderings in the same order', () => {
  assert.deepStrictEqual(src('flow/definition').SITE_ORDERS, src('services/Sites').SITE_ORDERS);
});

test('search refused: the run fails before scanning and the note still says Started', async () => {
  const tenant = scenarios.rich();
  tenant.search = 'reject';
  const { run } = await viaFlow(tenant);
  assert.strictEqual(run.status, 'Failed');
  assert.strictEqual(run.counts.Search_sites, 1, 'the loop does not repeat a failing search');
  assert.strictEqual(run.counts.Probe_site_order_3, 1);
  assert.strictEqual(run.counts.For_each_site, undefined);
  assert.deepStrictEqual(Object.keys(tenant.files), []);
  const note = tenant.settingsItem().ARLastRun;
  assert.ok(/^Failed 2026-12-01 03:04 UTC: SharePoint search refused every way of listing the sites\. Created:ascending: 400 We didn't understand your search terms\. Make sure they're using proper syntax\.; LastModifiedTime:ascending: 400 .*; no order: 400 /.test(note), note);
});

test('set-up not done: the run fails at the first step, nothing is written', async () => {
  const tenant = new FakeTenant({ sites: scenarios.richSites() });
  const { run } = await viaFlow(tenant);
  assert.strictEqual(run.status, 'Failed');
  assert.strictEqual(run.results.Get_settings.status, 'Failed');
  assert.strictEqual(run.counts.Mark_started, undefined);
  assert.deepStrictEqual(Object.keys(tenant.files), []);
});

test('report folder missing: the scan runs but the run fails at saving', async () => {
  const tenant = scenarios.rich({ folder: '/sites/it/Nowhere' });
  const { run } = await viaFlow(tenant);
  assert.strictEqual(run.results.Save_the_report.status, 'Failed');
  assert.strictEqual(run.status, 'Failed');
  assert.ok(/^Started/.test(tenant.settingsItem().ARLastRun));
});

// ---------------------------------------------------------------- packages
function unzipTo(bytes) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ar-pkg-'));
  const file = path.join(dir, 'p.zip');
  fs.writeFileSync(file, Buffer.from(bytes));
  const list = execFileSync('python3', ['-c', 'import zipfile,sys,json\nz=zipfile.ZipFile(sys.argv[1])\nassert z.testzip() is None\nprint(json.dumps({n:z.read(n).decode("utf-8") for n in z.namelist()}))', file], { maxBuffer: 64 * 1024 * 1024 }).toString();
  fs.rmSync(dir, { recursive: true, force: true });
  return JSON.parse(list);
}

function wellFormedXml(text) {
  execFileSync('python3', ['-c', 'import sys,xml.dom.minidom\nxml.dom.minidom.parseString(sys.stdin.buffer.read())'], { input: text });
}

test('legacy package: a valid zip with the files and ids Power Automate expects', () => {
  const files = unzipTo(pkg.legacyPackage('https://contoso.sharepoint.com/sites/it'));
  const names = Object.keys(files);
  const manifest = JSON.parse(files['manifest.json']);
  const flowId = Object.keys(manifest.resources).find((k) => manifest.resources[k].type === 'Microsoft.Flow/flows');
  assert.deepStrictEqual(names.sort(), ['manifest.json', 'Microsoft.Flow/flows/manifest.json', `Microsoft.Flow/flows/${flowId}/apisMap.json`, `Microsoft.Flow/flows/${flowId}/connectionsMap.json`, `Microsoft.Flow/flows/${flowId}/definition.json`].sort());
  assert.deepStrictEqual(JSON.parse(files['Microsoft.Flow/flows/manifest.json']).flowAssets.assetPaths, [flowId]);
  const apis = JSON.parse(files[`Microsoft.Flow/flows/${flowId}/apisMap.json`]);
  const conns = JSON.parse(files[`Microsoft.Flow/flows/${flowId}/connectionsMap.json`]);
  assert.strictEqual(manifest.resources[apis.shared_sharepointonline].type, 'Microsoft.PowerApps/apis');
  assert.strictEqual(manifest.resources[conns.shared_sharepointonline].type, 'Microsoft.PowerApps/apis/connections');
  assert.deepStrictEqual(manifest.resources[flowId].dependsOn.sort(), [apis.shared_sharepointonline, conns.shared_sharepointonline].sort());
  const def = JSON.parse(files[`Microsoft.Flow/flows/${flowId}/definition.json`]);
  assert.strictEqual(def.name, flowId);
  assert.deepStrictEqual(Object.keys(def.properties.connectionReferences), ['shared_sharepointonline']);
  assert.deepStrictEqual(def.properties.definition, JSON.parse(JSON.stringify(buildDefinition({ homeUrl: 'https://contoso.sharepoint.com/sites/it' }))));
  assert.strictEqual(def.properties.definition.actions.Home_site_URL.inputs, 'https://contoso.sharepoint.com/sites/it');
});

test('solution package: a valid zip, well-formed XML, matching ids and connection reference', () => {
  const files = unzipTo(pkg.solutionPackage('https://contoso.sharepoint.com/sites/it'));
  ['[Content_Types].xml', 'solution.xml', 'customizations.xml'].forEach((n) => wellFormedXml(files[n]));
  const jsonName = Object.keys(files).find((n) => /^Workflows\/.*\.json$/.test(n));
  const id = /RootComponent type="29" id="\{([0-9a-f-]+)\}"/.exec(files['solution.xml'])[1];
  assert.ok(files['customizations.xml'].indexOf(`WorkflowId="{${id}}"`) > 0);
  assert.ok(files['customizations.xml'].indexOf(`<JsonFileName>/${jsonName}</JsonFileName>`) > 0);
  assert.ok(jsonName.indexOf(id.toUpperCase()) > 0);
  const wf = JSON.parse(files[jsonName]);
  const ref = wf.properties.connectionReferences.shared_sharepointonline.connection.connectionReferenceLogicalName;
  assert.ok(files['customizations.xml'].indexOf(`connectionreferencelogicalname="${ref}"`) > 0);
  assert.ok(/^ar_/.test(ref) && /<CustomizationPrefix>ar<\/CustomizationPrefix>/.test(files['solution.xml']));
  assert.strictEqual(wf.properties.definition.actions.Home_site_URL.inputs, 'https://contoso.sharepoint.com/sites/it');
});

test('the packages in releases/ are the current build', () => {
  const dir = path.join(__dirname, '..', 'releases');
  [
    [pkg.LEGACY_FILE, pkg.legacyPackage(PLACEHOLDER_URL)],
    [pkg.SOLUTION_FILE, pkg.solutionPackage(PLACEHOLDER_URL)]
  ].forEach(([name, bytes]) => {
    assert.ok(fs.existsSync(path.join(dir, name)), `${name} exists (npm run build:flow)`);
    assert.ok(Buffer.from(bytes).equals(fs.readFileSync(path.join(dir, name))), `${name} is up to date (npm run build:flow)`);
  });
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
      console.log(`  FAIL ${name}\n       ${e && e.stack ? e.stack.split('\n').slice(0, 12).join('\n       ') : e}`);
    }
  }
  console.log(`\n${tests.length - failed}/${tests.length} passed`);
  process.exit(failed ? 1 : 0);
})();
