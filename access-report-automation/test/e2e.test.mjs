import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, readFileSync, existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';
import { parseCsv } from '../dist/csv.js';
import { main } from '../dist/cli.js';
import { ALL_PRINCIPALS, comparable, envFor, fixture, FOLDER, ID, ORACLE_KEYS, ORACLE_PRINCIPALS, oracleComparable, runTool, snapshotOf, startTenant, testCert } from './helpers.mjs';

const P = '2026-11';
const guestKey = 'user:' + ID.carol;

test('full run: report matches the permission model, files are posted, state is tidy', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  const { result, error, lines } = await runTool(tenant, { opts: { period: P } });
  assert.ifError(error);
  assert.equal(result.status, 'done');
  // 4 sites listed by Graph (Home, alpha, beta, delta, gamma = 5; personal left out; report site hidden).
  assert.equal(result.sitesTotal, 5);
  assert.equal(result.sitesFailed, 1, 'the locked site');
  assert.equal(result.principals, 8);

  const snap = snapshotOf(tenant, P);
  assert.ok(snap, 'snapshot saved');
  assert.deepEqual(comparable(snap.rows), oracleComparable(tenant, ORACLE_PRINCIPALS, ORACLE_KEYS));
  assert.deepEqual(snap.failedSites, [tenant.origin + '/sites/gamma']);

  const row = (p, url) => snap.rows.find((r) => r.p === p && r.url === tenant.origin + url);
  // How: nested security group -> SharePoint group.
  const aliceAlpha = row('user:' + ID.alice, '/sites/alpha');
  assert.equal(aliceAlpha.access, 'Full control');
  assert.match(aliceAlpha.how, /Group: Marketing \(owners\) -> SharePoint group: Alpha Owners/);
  assert.match(aliceAlpha.how, /Group: All staff -> SharePoint group: Alpha Visitors/);
  assert.equal(aliceAlpha.levels, 'Full Control, Read');
  // Direct grant on a library with its own permissions, through a security group.
  const aliceContracts = row('user:' + ID.alice, '/sites/alpha/Contracts');
  assert.equal(aliceContracts.unique, true);
  assert.equal(aliceContracts.how, 'Group: Finance');
  // Inheriting library.
  const aliceDocs = row('user:' + ID.alice, '/sites/alpha/Shared%20Documents');
  assert.equal(aliceDocs.how, 'Inherits from site');
  assert.equal(aliceDocs.unique, false);
  // Guest: only the one library, not the site (Limited Access is not access).
  assert.equal(row(guestKey, '/sites/alpha'), undefined);
  assert.equal(row(guestKey, '/sites/alpha/Contracts').how, 'Direct');
  assert.equal(row(guestKey, '/sites/beta'), undefined, 'guests are not in "everyone except external users"');
  assert.equal(row(guestKey, '').how, 'Everyone -> SharePoint group: Home Visitors');
  assert.equal(row(guestKey, '').access, 'View only');
  // Site collection administrator.
  const erin = row('user:' + ID.erin, '/sites/alpha/hr/payroll/Runs');
  assert.equal(erin.access, 'Full control');
  assert.equal(erin.how, 'Site collection administrator');
  // Subsites: own permissions, and inheriting ones.
  assert.equal(row('user:' + ID.bob, '/sites/alpha/hr').how, 'Direct');
  assert.equal(row('user:' + ID.bob, '/sites/alpha/hr').object, 'Subsite');
  assert.equal(row('user:' + ID.bob, '/sites/alpha/hr/payroll').how, 'Inherits from parent site');
  assert.equal(row('user:' + ID.bob, '/sites/alpha/news').access, 'Edit');
  // Group principals.
  assert.equal(row('group:' + ID.fin, '/sites/alpha').how, 'Group: All staff -> SharePoint group: Alpha Visitors');
  assert.equal(row('group:' + ID.mkt, '/sites/alpha').access, 'Full control');
  assert.equal(row('user:' + ID.bob, '/sites/beta').how.split('; ').includes('Everyone except external users'), true);
  // Nobody has anything on delta; dave is only a visitor there.
  assert.equal(snap.rows.filter((r) => r.siteUrl.endsWith('/sites/delta')).length, 0);
  // Lists and system / hidden libraries are out of scope by default.
  assert.equal(snap.rows.filter((r) => /Tasks|SiteAssets|_catalogs/.test(r.url)).length, 0);

  // Files.
  const xlsx = tenant.file(`${FOLDER}/Access report ${P}.xlsx`);
  const csv = tenant.file(`${FOLDER}/csv/Access report ${P}.csv`);
  const issues = tenant.file(`${FOLDER}/csv/Access report ${P} - not read.csv`);
  assert.ok(xlsx && csv && issues);
  assert.equal(tenant.file(`${FOLDER}/csv/Access report ${P} - changes.csv`), undefined, 'no changes file for a first report');
  assert.equal(tenant.file(`${FOLDER}/_state/checkpoint-${P}.json.gz`), undefined, 'progress file removed');
  const table = parseCsv(csv.toString('utf8'));
  assert.deepEqual(table[0], ['Who', 'Site URL', 'Site title', 'Object', 'Name', 'URL', 'Access', 'Permission levels', 'How', 'Own permissions', 'Note']);
  assert.equal(table.length - 1, snap.rows.length);
  assert.equal(csv[0], 0xef, 'byte-order mark for Excel');
  // A site title starting with "=" must not become a formula.
  assert.ok(table.some((r) => r[2] === "'Alpha =team") === false, 'title does not start with =, so it is untouched');
  assert.ok(table.some((r) => r[2] === 'Alpha =team'));
  assert.match(issues.toString('utf8'), /Site not read,.*\/sites\/gamma/);

  // The workbook opens in a real spreadsheet library and holds the same rows.
  const dir = mkdtempSync(join(tmpdir(), 'ar-x-'));
  writeFileSync(join(dir, 'r.xlsx'), xlsx);
  const py = `import openpyxl,json,sys
wb=openpyxl.load_workbook(sys.argv[1])
print(json.dumps({n:[[c for c in r] for r in wb[n].iter_rows(values_only=True)] for n in wb.sheetnames}))`;
  let book;
  try {
    book = JSON.parse(execFileSync('python3', ['-c', py, join(dir, 'r.xlsx')], { encoding: 'utf8', maxBuffer: 1 << 26 }));
  } catch (e) {
    t.diagnostic('openpyxl not available; workbook content not cross-checked: ' + String(e.message).split('\n')[0]);
  }
  if (book) {
    assert.deepEqual(Object.keys(book), ['Summary', 'Access', 'Changes', 'Not read']);
    assert.equal(book.Access.length - 1, snap.rows.length);
    assert.deepEqual(book.Access[0], table[0]);
    assert.equal(book.Summary[0][0], `Access report ${P}`);
    const sum = book.Summary.find((r) => r[0] && String(r[0]).startsWith('Alice Adams'));
    assert.equal(sum[1], 3, 'alice: Home, alpha, beta');
    assert.equal(sum[6], 'Full control');
    assert.equal(book['Not read'].length, 2);
  }

  // The only writes anywhere are sign-in and files in the report folder.
  const writes = tenant.requests.filter((r) => r.method !== 'GET');
  assert.ok(writes.length > 0);
  for (const w of writes) {
    const ok = /^\/login\/.*\/token$/.test(w.path) || (w.path.startsWith('/sites/gov/_api/web/') && decodeURIComponent(w.search).includes("'/sites/gov/Shared Documents"));
    assert.ok(ok, `unexpected write: ${w.method} ${w.path}`);
  }
  assert.equal(tenant.count((r) => r.method === 'GET' && r.body.length > 0), 0);

  // Nothing identifying in the default log.
  const text = lines.join('\n');
  assert.ok(!/contoso|127\.0\.0\.1|alpha|Alice/i.test(text), 'log is masked: ' + text);
});

test('second run in the same month does nothing; --force makes it again', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  assert.equal((await runTool(tenant, { opts: { period: P } })).result.status, 'done');
  const before = tenant.requests.length;
  const again = await runTool(tenant, { opts: { period: P } });
  assert.equal(again.result.status, 'skipped');
  assert.ok(tenant.requests.length - before <= 6, `a no-op run is cheap (${tenant.requests.length - before} requests)`);
  const forced = await runTool(tenant, { opts: { period: P, force: true } });
  assert.equal(forced.result.status, 'done');
  assert.equal(forced.result.comparedWith, undefined, 'a forced re-run is not compared with itself');
});

test('next month: changes since the previous report', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  assert.equal((await runTool(tenant, { opts: { period: '2026-10' } })).result.status, 'done');

  const m = tenant.model;
  const beta = m.sites.find((s) => s.path === '/sites/beta');
  beta.web.assignments = beta.web.assignments.filter((a) => a.to !== 'user:bob@contoso.com'); // bob: Design -> Read (via EEEU)
  m.sites.find((s) => s.path === '/sites/delta').web.assignments.push({ to: 'user:dave@contoso.com', roles: ['Contribute'] }); // dave: added
  const alpha = m.sites.find((s) => s.path === '/sites/alpha');
  alpha.web.lists.find((l) => l.title === 'Contracts').assignments = [{ to: 'group:' + ID.fin, roles: ['Contribute'] }]; // carol: removed
  alpha.web.lists.find((l) => l.title === 'Documents').title = 'Documents (renamed)'; // not a change
  m.sites.find((s) => s.path === '/sites/gamma').locked = false; // readable now, but unread last month: left out of changes
  tenant.index();

  const { result, error } = await runTool(tenant, { opts: { period: P } });
  assert.ifError(error);
  assert.equal(result.comparedWith, '2026-10');
  const changes = parseCsv(tenant.file(`${FOLDER}/csv/Access report ${P} - changes.csv`).toString('utf8')).slice(1);
  const find = (change, who, url) => changes.filter((c) => c[0] === change && c[1].startsWith(who) && c[6] === tenant.origin + url);
  assert.equal(find('Less access', "Bob O'Brien", '/sites/beta').length, 1);
  assert.deepEqual(find('Less access', "Bob O'Brien", '/sites/beta')[0].slice(7, 9), ['Design', 'Read']);
  assert.equal(find('Added', 'Dave Dunn', '/sites/delta').length, 1);
  assert.equal(find('Added', 'Dave Dunn', '/sites/delta/Shared%20Documents').length, 1);
  assert.equal(find('Removed', 'Carol', '/sites/alpha/Contracts').length, 1);
  assert.equal(changes.filter((c) => /gamma/.test(c[2])).length, 0, 'a site that was unreadable last month is not "added"');
  assert.equal(changes.filter((c) => /renamed/.test(c[5])).length, 0, 'a renamed library is not a change');
  assert.equal(result.added, 2);
  assert.equal(result.removed, 1);
  // Everything not touched is not listed.
  assert.equal(changes.filter((c) => c[1].startsWith('Alice')).length, 0);
});

test('time limit: stops cleanly, the next run resumes and the result equals an uninterrupted run', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  // A clock that jumps past the limit after the second site has been read.
  let sitesRead = 0;
  tenant.faults.push((r) => {
    if (/siteusers$/.test(r.path) && /PrincipalType eq 4/.test(decodeURIComponent(r.search))) sitesRead++;
    return undefined;
  });
  const start = Date.parse('2026-11-01T02:00:00Z');
  const now = () => (sitesRead >= 2 ? start + 10 * 60000 : start);
  const first = await runTool(tenant, { file: { scan: { maxMinutes: 5, concurrency: 1 } }, opts: { now } });
  assert.ifError(first.error);
  assert.equal(first.result.status, 'partial');
  assert.ok(first.result.sitesDone + first.result.sitesFailed < first.result.sitesTotal);
  assert.ok(tenant.file(`${FOLDER}/_state/checkpoint-${P}.json.gz`), 'progress saved');
  assert.equal(tenant.file(`${FOLDER}/Access report ${P}.xlsx`), undefined, 'no half report');

  tenant.faults.length = 0;
  const scannedBefore = tenant.count((r) => /\/sites\/alpha\/_api\/web\/roleassignments$/.test(r.path));
  const second = await runTool(tenant, { file: { scan: { maxMinutes: 5, concurrency: 1 } }, opts: { period: P } });
  assert.ifError(second.error);
  assert.equal(second.result.status, 'done');
  assert.ok(second.lines.some((l) => /Resuming: \d of 5/.test(l)));
  // Sites are sorted by address: Home and alpha were read in the first run and must not be read again.
  assert.equal(tenant.count((r) => /\/sites\/alpha\/_api\/web\/roleassignments$/.test(r.path)), scannedBefore);
  assert.deepEqual(comparable(snapshotOf(tenant, P).rows), oracleComparable(tenant, ORACLE_PRINCIPALS, ORACLE_KEYS));
  assert.equal(tenant.file(`${FOLDER}/_state/checkpoint-${P}.json.gz`), undefined);
});

test('saved progress from different settings is not reused', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  let n = 0;
  tenant.faults.push((r) => {
    if (/PrincipalType eq 4/.test(decodeURIComponent(r.search))) n++;
    return undefined;
  });
  const start = Date.parse('2026-11-01T02:00:00Z');
  const first = await runTool(tenant, { file: { scan: { maxMinutes: 5, concurrency: 1 } }, opts: { now: () => (n >= 2 ? start + 10 * 60000 : start) } });
  assert.equal(first.result.status, 'partial');
  tenant.faults.length = 0;
  const second = await runTool(tenant, { file: { principals: ['user:alice@contoso.com'] }, opts: { period: P } });
  assert.equal(second.result.status, 'done');
  assert.ok(second.lines.some((l) => /different settings/.test(l)));
  assert.ok(snapshotOf(tenant, P).rows.every((r) => r.p === 'user:' + ID.alice));
});

test('throttling (429 with Retry-After, 503) and a dropped connection are ridden out', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  let n = 0;
  tenant.faults.push((r) => {
    if (!r.path.includes('/_api/')) return undefined;
    n++;
    if (n === 3) return { status: 429, headers: { 'Retry-After': '1' }, once: false };
    if (n === 9) return { status: 503, once: false };
    if (n === 15) return { drop: true, once: false };
    if (n === 21) return { status: 500, once: false };
    return { once: false };
  });
  const t0 = Date.now();
  const { result, error, lines } = await runTool(tenant, { opts: { period: P } });
  assert.ifError(error);
  assert.equal(result.status, 'done');
  assert.equal(result.throttled, 2);
  assert.ok(Date.now() - t0 >= 1000, 'Retry-After was honoured');
  assert.ok(lines.some((l) => /Throttled/.test(l)));
  assert.deepEqual(comparable(snapshotOf(tenant, P).rows), oracleComparable(tenant, ORACLE_PRINCIPALS, ORACLE_KEYS));
});

test('a token that expires mid-run is renewed', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  let n = 0;
  tenant.faults.push((r) => {
    if (r.path.includes('/_api/') && ++n === 12) tenant.validTokens.clear();
    return { once: false };
  });
  const { result, error } = await runTool(tenant, { opts: { period: P } });
  assert.ifError(error);
  assert.equal(result.status, 'done');
  assert.ok(tenant.tokensIssued >= 3, 'signed in again');
  assert.deepEqual(comparable(snapshotOf(tenant, P).rows), oracleComparable(tenant, ORACLE_PRINCIPALS, ORACLE_KEYS));
});

test('an outage stops the run with progress saved; the next run finishes it', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  // Everything about beta, delta and gamma answers 500 (not a single-site kind of problem when it keeps happening).
  const model = fixture();
  for (let i = 0; i < 10; i++) model.sites.push({ path: `/sites/z${i}`, title: `Z${i}`, web: { path: `/sites/z${i}`, title: `Z${i}`, assignments: [{ to: 'eeeu', roles: ['Read'] }] } });
  const big = await startTenant(model);
  t.after(() => big.stop());
  big.faults.push((r) => (/\/sites\/z\d\/_api\//.test(r.path) ? { status: 500, once: false } : { once: false }));
  const first = await runTool(big, { file: { scan: { concurrency: 1 } }, opts: { period: P, maxRetries: 1 } });
  assert.ok(first.error, 'the run fails');
  assert.match(first.error.message, /Stopped after 8 errors in a row/);
  assert.ok(big.file(`${FOLDER}/_state/checkpoint-${P}.json.gz`));
  assert.equal(big.file(`${FOLDER}/Access report ${P}.xlsx`), undefined);
  big.faults.length = 0;
  const second = await runTool(big, { file: { scan: { concurrency: 1 } }, opts: { period: P } });
  assert.ifError(second.error);
  assert.equal(second.result.status, 'done');
  assert.equal(second.result.sitesFailed, 1, 'only the locked site');
  assert.deepEqual(comparable(snapshotOf(big, P).rows), oracleComparable(big, ORACLE_PRINCIPALS, ORACLE_KEYS));
});

test('a passing failure of one site is retried at the end of the run', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  let hits = 0;
  tenant.faults.push((r) => (/\/sites\/beta\/_api\/web\/roleassignments/.test(r.path) && ++hits <= 3 ? { status: 500, once: false } : { once: false }));
  const { result, error, lines } = await runTool(tenant, { opts: { period: P } });
  assert.ifError(error);
  assert.ok(lines.some((l) => /Trying 1 site again/.test(l)));
  assert.equal(result.sitesFailed, 1);
  assert.deepEqual(comparable(snapshotOf(tenant, P).rows), oracleComparable(tenant, ORACLE_PRINCIPALS, ORACLE_KEYS));
});

test('scope switches: lists, no subsites, no libraries', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  for (const scope of [
    { includeLists: true },
    { includeSubsites: false },
    { includeLibraries: false },
    { includeLibraries: false, includeSubsites: false }
  ]) {
    const { error } = await runTool(tenant, { file: { scan: scope }, opts: { period: P, force: true } });
    assert.ifError(error);
    assert.deepEqual(comparable(snapshotOf(tenant, P).rows), oracleComparable(tenant, ORACLE_PRINCIPALS, ORACLE_KEYS, scope), JSON.stringify(scope));
  }
  const withLists = await runTool(tenant, { file: { scan: { includeLists: true } }, opts: { period: P, force: true } });
  assert.ifError(withLists.error);
  assert.ok(snapshotOf(tenant, P).rows.some((r) => r.object === 'List' && r.title === 'Tasks'));
});

test('fast mode (thorough off) and SharePoint refusing the PrincipalType filter give the same answer', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  const thorough = await runTool(tenant, { opts: { period: P } });
  const thoroughCalls = tenant.count((r) => /getusereffectivepermissions/.test(r.path));
  tenant.requests.length = 0;
  tenant.rejectPrincipalTypeFilter = true;
  const fast = await runTool(tenant, { file: { scan: { thorough: false } }, opts: { period: P, force: true } });
  assert.ifError(thorough.error || fast.error);
  assert.deepEqual(comparable(snapshotOf(tenant, P).rows), oracleComparable(tenant, ORACLE_PRINCIPALS, ORACLE_KEYS));
  assert.ok(tenant.count((r) => /getusereffectivepermissions/.test(r.path)) < thoroughCalls, 'fast mode asks less');
});

test('role assignments without masks still classify built-in levels', async (t) => {
  const tenant = await startTenant(fixture(), { omitBasePermissions: true });
  t.after(() => tenant.stop());
  const { error } = await runTool(tenant, { file: { principals: ['group:Finance', 'group:' + ID.mkt, 'group:All staff'] }, opts: { period: P } });
  assert.ifError(error);
  assert.deepEqual(comparable(snapshotOf(tenant, P).rows), oracleComparable(tenant, ORACLE_PRINCIPALS.slice(5), ORACLE_KEYS.slice(5)));
});

test('group access SharePoint has not caught up with is reported, and marked unconfirmed', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  tenant.staleGroupUsers.add('alice@contoso.com');
  const { error } = await runTool(tenant, { file: { principals: ['user:alice@contoso.com'] }, opts: { period: P } });
  assert.ifError(error);
  const rows = snapshotOf(tenant, P).rows;
  // Same places as with a fully informed SharePoint.
  assert.deepEqual(comparable(rows), oracleComparable(tenant, ORACLE_PRINCIPALS.slice(0, 1), ORACLE_KEYS.slice(0, 1)));
  const alpha = rows.find((r) => r.url === tenant.origin + '/sites/alpha');
  assert.match(alpha.note, /not confirmed/);
  // Access through "everyone except external users" is not a directory group: confirmed as usual.
  assert.equal(rows.find((r) => r.url === tenant.origin + '/sites/beta').note, '');
});

test('an "everyone" grant that SharePoint does not apply to a guest is not reported', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  tenant.everyoneExcludesGuests = true;
  const { error } = await runTool(tenant, { file: { principals: ['user:carol@fabrikam.com'] }, opts: { period: P } });
  assert.ifError(error);
  const rows = snapshotOf(tenant, P).rows;
  assert.deepEqual(rows.map((r) => r.url), [tenant.origin + '/sites/alpha/Contracts']);
  assert.equal(rows[0].note, '');
});

test('an account SharePoint cannot evaluate does not break the run; group access is still reported', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  tenant.unknownUsers.add('alice@contoso.com');
  const { result, error } = await runTool(tenant, { opts: { period: P } });
  assert.ifError(error);
  assert.equal(result.sitesFailed, 1, 'only the locked site');
  assert.equal(result.warnings.length, 1);
  assert.match(result.warnings[0], /could not evaluate Alice Adams/);
  const rows = snapshotOf(tenant, P).rows;
  const alice = rows.filter((r) => r.p === 'user:' + ID.alice);
  // Routes through directory groups are reported (unconfirmed); "everyone" routes are not guessed.
  assert.ok(alice.length > 0 && alice.every((r) => /not confirmed/.test(r.note)));
  assert.ok(alice.some((r) => r.url === tenant.origin + '/sites/alpha' && r.access === 'Full control'));
  assert.equal(alice.filter((r) => r.siteUrl === tenant.origin + '/sites/beta').length, 0);
  // Everyone else is unaffected.
  const others = ORACLE_PRINCIPALS.slice(1);
  assert.deepEqual(comparable(rows.filter((r) => r.p !== 'user:' + ID.alice)), oracleComparable(tenant, others, ORACLE_KEYS.slice(1)));
});

test('a library deleted while its site is being read is skipped; odd characters in addresses are encoded', async (t) => {
  const model = fixture();
  model.sites.find((s) => s.path === '/sites/beta').web.lists.push({ id: 'bbbbbbbb-0000-0000-0000-00000000000f', title: 'Q#1 plans?', path: '/sites/beta/Q#1 plans?', assignments: [{ to: 'user:bob@contoso.com', roles: ['Read'] }] });
  const tenant = await startTenant(model);
  t.after(() => tenant.stop());
  tenant.faults.push((r) => (/aaaaaaaa-0000-0000-0000-00000000000b'\)\/roleassignments/.test(r.path) ? { status: 404, body: { 'odata.error': { message: { value: 'List does not exist.' } } }, once: false } : { once: false }));
  const { result, error } = await runTool(tenant, { opts: { period: P } });
  assert.ifError(error);
  assert.equal(result.sitesFailed, 1, 'only the locked site: alpha is still reported');
  const rows = snapshotOf(tenant, P).rows;
  assert.ok(rows.some((r) => r.url === tenant.origin + '/sites/alpha'));
  assert.equal(rows.filter((r) => /Contracts/.test(r.url)).length, 0);
  const odd = rows.find((r) => r.title === 'Q#1 plans?');
  assert.equal(odd.url, tenant.origin + '/sites/beta/Q%231%20plans%3F');
  assert.equal(new URL(odd.url).pathname, '/sites/beta/Q%231%20plans%3F');
});

test('without permission to read group memberships: levels are still right, with a warning', async (t) => {
  const tenant = await startTenant(fixture(), { graphRoles: ['Sites.Read.All', 'User.Read.All'] });
  t.after(() => tenant.stop());
  const users = ALL_PRINCIPALS.slice(0, 5);
  const { result, error } = await runTool(tenant, { file: { principals: users }, opts: { period: P } });
  assert.ifError(error);
  assert.equal(result.warnings.length, 1);
  assert.match(result.warnings[0], /GroupMember\.Read\.All/);
  const rows = snapshotOf(tenant, P).rows;
  assert.deepEqual(comparable(rows), oracleComparable(tenant, ORACLE_PRINCIPALS.slice(0, 5), ORACLE_KEYS.slice(0, 5)));
  assert.equal(rows.find((r) => r.p === 'user:' + ID.alice && r.url === tenant.origin + '/sites/alpha').how, 'Effective permission (the route was not identified)');
});

test('expandGroups off behaves the same without asking Graph for memberships', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  const { error } = await runTool(tenant, { file: { principals: ALL_PRINCIPALS.slice(0, 5), scan: { expandGroups: false } }, opts: { period: P } });
  assert.ifError(error);
  assert.equal(tenant.count((r) => /transitiveMemberOf|ownedObjects/.test(r.path)), 0);
  assert.deepEqual(comparable(snapshotOf(tenant, P).rows), oracleComparable(tenant, ORACLE_PRINCIPALS.slice(0, 5), ORACLE_KEYS.slice(0, 5)));
});

test('people and groups come from principals.csv in the report folder; unknown entries are reported, not fatal', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  // Nothing configured: a template is created and the run asks for it to be filled in.
  const empty = await runTool(tenant, { file: { principals: [] }, opts: { period: P } });
  assert.equal(empty.error.name, 'ConfigError');
  assert.match(empty.error.message, /principals\.csv was created/);
  assert.equal(tenant.file(`${FOLDER}/principals.csv`).toString('utf8'), '﻿Type,Identity,Notes\r\n');

  tenant.files.set(`${FOLDER}/principals.csv`, Buffer.from('﻿Type;Identity;Notes\r\nuser;alice@contoso.com;CFO\r\ngroup;Marketing;\r\nuser;ghost@contoso.com;left in May\r\n;bob@contoso.com;\r\n# comment;;\r\nteam;x;\r\n'));
  const { result, error } = await runTool(tenant, { file: { principals: [] }, opts: { period: P } });
  assert.ifError(error);
  assert.equal(result.principals, 3);
  assert.equal(result.unresolved, 1);
  const issues = tenant.file(`${FOLDER}/csv/Access report ${P} - not read.csv`).toString('utf8');
  assert.match(issues, /Not in the directory,ghost@contoso\.com/);
});

test('sites.only and include / exclude patterns', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  const only = await runTool(tenant, { file: { sites: { only: [tenant.origin + '/sites/beta/'] } }, opts: { period: P } });
  assert.ifError(only.error);
  assert.equal(only.result.sitesTotal, 1);
  assert.equal(tenant.count((r) => /getAllSites/.test(r.path)), 0);
  const filtered = await runTool(tenant, { file: { sites: { include: [tenant.origin + '/sites/*'], exclude: ['*/sites/g*', '*/delta'] } }, opts: { period: P, force: true } });
  assert.ifError(filtered.error);
  assert.equal(filtered.result.sitesTotal, 2);
  assert.equal(filtered.result.sitesFailed, 0);
});

test('a subsite that Graph also lists as a site is reported once', async (t) => {
  const model = fixture();
  model.sites.push({ path: '/sites/alpha/hr', title: 'HR (listed)', dup: true, web: { path: '/sites/alpha/hr-dup', title: 'x', assignments: [] } });
  const tenant = await startTenant(model);
  t.after(() => tenant.stop());
  // The duplicate entry answers as the real subsite does: point it at the same web.
  const real = model.sites.find((s) => s.path === '/sites/alpha').web.webs[0];
  const dup = model.sites.find((s) => s.dup);
  dup.web = real;
  dup.spGroups = model.sites.find((s) => s.path === '/sites/alpha').spGroups;
  dup.admins = ['user:erin@contoso.com'];
  tenant.index();
  const { error } = await runTool(tenant, { opts: { period: P } });
  assert.ifError(error);
  const rows = snapshotOf(tenant, P).rows.filter((r) => r.p === 'user:' + ID.bob && r.url === tenant.origin + '/sites/alpha/hr');
  assert.equal(rows.length, 1);
});

test('sign-in: federated works; a wrong certificate, a wrong tenant and missing permissions fail with the fix', async (t) => {
  const tenant = await startTenant(fixture(), { allowFederated: true });
  t.after(() => tenant.stop());
  const fed = await runTool(tenant, {
    env: { ACCESS_REPORT_AUTH: 'federated', ACCESS_REPORT_CERTIFICATE: '', ACTIONS_ID_TOKEN_REQUEST_URL: tenant.origin + '/oidc?api-version=2', ACTIONS_ID_TOKEN_REQUEST_TOKEN: tenant.oidcRequestToken },
    opts: { period: P }
  });
  assert.ifError(fed.error);
  assert.equal(fed.result.status, 'done');

  const odir = mkdtempSync(join(tmpdir(), 'ar-other-'));
  execFileSync('openssl', ['req', '-x509', '-newkey', 'rsa:2048', '-nodes', '-keyout', join(odir, 'k.pem'), '-out', join(odir, 'c.pem'), '-days', '2', '-subj', '/CN=other'], { stdio: 'ignore' });
  const other = readFileSync(join(odir, 'c.pem'), 'utf8') + readFileSync(join(odir, 'k.pem'), 'utf8');
  const wrongCert = await runTool(tenant, { env: { ACCESS_REPORT_CERTIFICATE: other }, opts: { period: '2026-12' } });
  assert.match(wrongCert.error.message, /AADSTS700027.*not registered/);
  assert.ok(!wrongCert.error.message.includes('PRIVATE KEY'));

  const wrongTenant = await runTool(tenant, { env: { ACCESS_REPORT_TENANT_ID: '99999999-9999-9999-9999-999999999999' }, opts: { period: '2026-12' } });
  assert.match(wrongTenant.error.message, /AADSTS90002.*tenant/i);

  const noSp = await startTenant(fixture(), { spRoles: ['Sites.Read.All'] });
  t.after(() => noSp.stop());
  const refused = await runTool(noSp, { opts: { period: P } });
  assert.match(refused.error.message, /Sites\.FullControl\.All/);

  const noGraphSites = await startTenant(fixture(), { graphRoles: ['User.Read.All', 'GroupMember.Read.All'] });
  t.after(() => noGraphSites.stop());
  const noList = await runTool(noGraphSites, { opts: { period: P } });
  assert.match(noList.error.message, /Sites\.Read\.All/);
});

test('the report folder: a missing library is an error, missing folders are created, a locked file is saved beside', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  const missing = await runTool(tenant, { env: { ACCESS_REPORT_FOLDER_URL: `${tenant.origin}/sites/gov/Reports/Access` }, opts: { period: P } });
  assert.match(missing.error.message, /document library \/sites\/gov\/Reports does not exist/);
  assert.equal(tenant.count((r) => /getAllSites/.test(r.path)), 0, 'found before any scanning');

  const deep = await runTool(tenant, { env: { ACCESS_REPORT_FOLDER_URL: `${tenant.origin}/sites/gov/Shared%20Documents/Governance/Access reports/Monthly` }, opts: { period: P } });
  assert.ifError(deep.error);
  assert.ok(tenant.file(`/sites/gov/Shared Documents/Governance/Access reports/Monthly/Access report ${P}.xlsx`));

  tenant.lockedFiles.add(`${FOLDER}/Access report ${P}.xlsx`.toLowerCase());
  const locked = await runTool(tenant, { opts: { period: P } });
  assert.ifError(locked.error);
  assert.match(locked.result.files[0], /Access report 2026-11 \(\d{8}-\d{4}\)\.xlsx$/);
  assert.ok(tenant.file(locked.result.files[0]));
});

test('trial run: --no-upload writes locally and touches nothing in SharePoint', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  const out = mkdtempSync(join(tmpdir(), 'ar-out-'));
  const { result, error } = await runTool(tenant, { opts: { period: P, upload: false, outDir: out, limit: 2 } });
  assert.ifError(error);
  assert.equal(result.sitesTotal, 2);
  assert.ok(existsSync(join(out, `Access report ${P}.xlsx`)));
  assert.ok(existsSync(join(out, 'csv', `Access report ${P}.csv`)));
  assert.equal(tenant.files.size, 0);
  assert.equal(tenant.count((r) => r.method === 'POST' && r.path.includes('/_api/')), 0);
  assert.match(readFileSync(join(out, 'csv', `Access report ${P}.csv`), 'utf8'), /Home/);
});

test('command line: exit codes, check command, step summary', async (t) => {
  const tenant = await startTenant();
  t.after(() => tenant.stop());
  const dir = mkdtempSync(join(tmpdir(), 'ar-cli-'));
  const cfgPath = join(dir, 'config.json');
  writeFileSync(cfgPath, JSON.stringify({ principals: ['user:alice@contoso.com', 'group:Finance'] }));
  const summary = join(dir, 'summary.md');
  const env = { ...envFor(tenant), GITHUB_STEP_SUMMARY: summary };
  const quiet = process.stderr.write;
  const captured = [];
  process.stderr.write = (s) => (captured.push(String(s)), true);
  t.after(() => (process.stderr.write = quiet));
  try {
    assert.equal(await main(['check', '--config', cfgPath], env), 0);
    assert.ok(captured.some((l) => /All checks passed/.test(l)), captured.join(''));
    assert.equal(tenant.file(`${FOLDER}/_state/check.txt`), undefined, 'check cleans up after itself');
    assert.equal(await main(['run', '--config', cfgPath, '--period', P], env), 0);
    const md = readFileSync(summary, 'utf8');
    assert.match(md, /Access report 2026-11: done/);
    assert.ok(!/contoso|127\.0\.0\.1/.test(md), 'summary has numbers only');
    assert.equal(await main(['run', '--config', cfgPath, '--period', P], env), 0, 'already done');
    assert.equal(await main(['run', '--config', cfgPath, '--period', 'November'], env), 2);
    assert.equal(await main(['run', '--config', join(dir, 'nope.json')], env), 2);
    assert.equal(await main(['run', '--bogus'], env), 2);
    assert.equal(await main(['run', '--config', cfgPath, '--period', '2026-12'], { ...env, ACCESS_REPORT_CLIENT_ID: '' }), 2);
    assert.equal(await main(['run', '--config', cfgPath, '--period', '2026-12'], { ...env, ACCESS_REPORT_CERTIFICATE: testCert().cert }), 1, 'certificate without its key');
    // check reports missing permissions.
    captured.length = 0;
    const weak = await startTenant(fixture(), { spRoles: [], graphRoles: ['User.Read.All'] });
    t.after(() => weak.stop());
    assert.equal(await main(['check', '--config', cfgPath], envFor(weak)), 1);
    const text = captured.join('');
    assert.match(text, /\[FAIL\].*Sites\.Read\.All/);
    assert.match(text, /\[FAIL\].*GroupMember\.Read\.All/);
    assert.match(text, /\[FAIL\].*Sites\.FullControl\.All/);
  } finally {
    process.stderr.write = quiet;
  }
});
