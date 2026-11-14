import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { createPublicKey, createVerify } from 'node:crypto';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';
import { CertificateCredential, explainSignInError, normalisePem, tokenRoles } from '../dist/auth.js';
import { parseArgs, stepSummary } from '../dist/cli.js';
import { buildConfig, parsePrincipalList, parsePrincipalsCsv, siteFilter, splitFolderUrl } from '../dist/config.js';
import { csvCell, parseCsv, toCsv } from '../dist/csv.js';
import { errorMessageFrom, Http, HttpError, parseRetryAfter } from '../dist/http.js';
import { redact } from '../dist/log.js';
import { adminExplanation, explainAssignments, groupClaims, permissionLevel, toLogin } from '../dist/permissions.js';
import { buildReport, diffSnapshots, packSnapshot, unpackSnapshot } from '../dist/report.js';
import { currentPeriod } from '../dist/runner.js';
import { alias, joinPath, serverRelative, SharePoint } from '../dist/sharepoint.js';
import { buildXlsx, columnName, crc32, sheetNames, xmlText } from '../dist/xlsx.js';
import { MockTenant } from './mock-tenant.mjs';
import { testCert } from './helpers.mjs';

const GOOD = { ACCESS_REPORT_TENANT_ID: '11111111-2222-3333-4444-555555555555', ACCESS_REPORT_CLIENT_ID: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee', ACCESS_REPORT_FOLDER_URL: 'https://contoso.sharepoint.com/sites/gov/Shared Documents/Access reports', ACCESS_REPORT_CERTIFICATE: 'x' };

test('permission masks to levels (the real masks of the built-in levels)', () => {
  const level = (low, high) => permissionLevel({ low, high });
  assert.equal(level(4294967295, 2147483647), 'Full control');
  assert.equal(level(1012866047, 432), 'Design');
  assert.equal(level(1011030767, 432), 'Edit');
  assert.equal(level(1011028719, 432), 'Contribute');
  assert.equal(level(138612833, 176), 'Read');
  assert.equal(level(138612801, 176), 'View only');
  assert.equal(level(134287360, 0), 'No access', 'Limited Access');
  assert.equal(level(0, 0), 'No access');
  // Read on the site + Contribute through a group = Contribute.
  assert.equal(level((138612833 | 1011028719) >>> 0, 176 | 432), 'Contribute');
});

test('claims and logins', () => {
  assert.equal(toLogin('Alice@Contoso.com'), 'i:0#.f|membership|alice@contoso.com');
  assert.equal(toLogin('i:0#.f|membership|x@y.z'), 'i:0#.f|membership|x@y.z');
  assert.deepEqual(groupClaims('security', 'ABCDEF00-0000-0000-0000-000000000000'), ['c:0t.c|tenant|abcdef00-0000-0000-0000-000000000000']);
  assert.deepEqual(groupClaims('m365', 'abc'), ['c:0o.c|federateddirectoryclaimprovider|abc', 'c:0o.c|federateddirectoryclaimprovider|abc_o']);
});

test('role assignments: routes, Limited Access, verbose OData shape', () => {
  const read = { Name: 'Read', RoleTypeKind: 2, BasePermissions: { Low: '138612833', High: '176' } };
  const edit = { Name: 'Edit', RoleTypeKind: 6, BasePermissions: { Low: '1011030767', High: '432' } };
  const limited = { Name: 'Limited Access', RoleTypeKind: 1, BasePermissions: { Low: '134287360', High: '0' } };
  const assignments = [
    { PrincipalId: 7, RoleDefinitionBindings: [read] },
    { PrincipalId: 20, RoleDefinitionBindings: { results: [edit] } },
    { PrincipalId: 7, RoleDefinitionBindings: [limited] },
    { PrincipalId: 99, RoleDefinitionBindings: [edit] }
  ];
  const e = explainAssignments(assignments, [{ id: 7, via: '', isSiteAdmin: false }], [{ id: 20, title: 'Members', via: 'Finance' }]);
  assert.deepEqual(e.levels, ['Read', 'Edit']);
  assert.deepEqual(e.how, ['Direct', 'Group: Finance -> SharePoint group: Members']);
  assert.equal(permissionLevel(e.bits), 'Edit');
  assert.equal(explainAssignments([{ PrincipalId: 7, RoleDefinitionBindings: [limited] }], [{ id: 7, via: '', isSiteAdmin: false }], []), undefined);
  assert.equal(explainAssignments(assignments, [], []), undefined);
  assert.deepEqual(adminExplanation([{ id: 1, via: 'Marketing (owners)', isSiteAdmin: true }]).how, ['Site collection administrator (group: Marketing (owners))']);
  assert.equal(adminExplanation([{ id: 1, via: '', isSiteAdmin: false }]), undefined);
});

test('CSV: quoting, formula guard, round trip, delimiters', () => {
  assert.equal(csvCell('=HYPERLINK("x")'), '"\'=HYPERLINK(""x"")"');
  assert.equal(csvCell('+1'), "'+1");
  assert.equal(csvCell('-x'), "'-x");
  assert.equal(csvCell('@x'), "'@x");
  assert.equal(csvCell('a,b'), '"a,b"');
  assert.equal(csvCell(null), '');
  const rows = [['a "q"', 'line\r\nbreak', 'ü, 你好', ''], ['', '', '', 'last']];
  const text = toCsv(['h1', 'h2', 'h3', 'h4'], rows);
  assert.equal(text.charCodeAt(0), 0xfeff);
  assert.deepEqual(parseCsv(text), [['h1', 'h2', 'h3', 'h4'], ...rows]);
  assert.deepEqual(parseCsv('a;b\n1;2\n\n'), [['a', 'b'], ['1', '2']]);
  assert.deepEqual(parseCsv('a\tb\r\n"x\ty"\t2'), [['a', 'b'], ['x\ty', '2']]);
  assert.deepEqual(parseCsv(''), []);
});

test('principal lists', () => {
  assert.deepEqual(parsePrincipalList('user:a@b.c; group: Finance team\n#x\nd@e.f\n12345678-1234-1234-1234-123456789012'), [
    { type: 'user', id: 'a@b.c' },
    { type: 'group', id: 'Finance team' },
    { type: 'user', id: 'd@e.f' },
    { type: 'group', id: '12345678-1234-1234-1234-123456789012' }
  ]);
  const r = parsePrincipalsCsv('Identity,Type\r\na@b.c,User\r\nFinance,GROUP\r\nx,robot\r\n');
  assert.deepEqual(r.principals, [{ type: 'user', id: 'a@b.c' }, { type: 'group', id: 'Finance' }]);
  assert.equal(r.problems.length, 1);
  assert.equal(parsePrincipalsCsv('foo,bar\r\n1,2').problems.length, 1);
  assert.deepEqual(parsePrincipalsCsv('').principals, []);
});

test('report folder addresses', () => {
  assert.deepEqual(splitFolderUrl('https://contoso.sharepoint.com/sites/gov/Shared%20Documents/Access%20reports/'), { siteUrl: 'https://contoso.sharepoint.com/sites/gov', libraryPath: '/sites/gov/Shared Documents', folder: 'Access reports' });
  assert.deepEqual(splitFolderUrl('https://contoso.sharepoint.com/teams/x/Docs'), { siteUrl: 'https://contoso.sharepoint.com/teams/x', libraryPath: '/teams/x/Docs', folder: '' });
  assert.deepEqual(splitFolderUrl('https://contoso.sharepoint.com/Shared Documents/a/b'), { siteUrl: 'https://contoso.sharepoint.com', libraryPath: '/Shared Documents', folder: 'a/b' });
  // Copied from the browser while looking at the folder.
  assert.deepEqual(splitFolderUrl('https://contoso.sharepoint.com/sites/gov/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2Fgov%2FShared%20Documents%2FAccess%20reports&viewid=abc'), {
    siteUrl: 'https://contoso.sharepoint.com/sites/gov',
    libraryPath: '/sites/gov/Shared Documents',
    folder: 'Access reports'
  });
  assert.deepEqual(splitFolderUrl('https://contoso.sharepoint.com/sites/gov/Shared%20Documents/Forms/AllItems.aspx'), { siteUrl: 'https://contoso.sharepoint.com/sites/gov', libraryPath: '/sites/gov/Shared Documents', folder: '' });
  assert.throws(() => splitFolderUrl('https://contoso.sharepoint.com/sites/gov'), /document library/);
  assert.throws(() => splitFolderUrl('https://contoso.sharepoint.com/sites/gov/SitePages/Home.aspx'), /not a page/);
  assert.throws(() => splitFolderUrl('http://contoso.sharepoint.com/sites/gov/Docs'), /https/);
  assert.throws(() => splitFolderUrl('nonsense'), /not a valid address/);
});

test('settings: defaults, environment wins, clear errors', async () => {
  const cfg = await buildConfig({ scan: { concurrency: 4, thorough: 'no' }, report: { name: 'Who can get in' }, principals: [{ type: 'group', id: 'X' }, 'a@b.c'] }, { ...GOOD, ACCESS_REPORT_MAX_MINUTES: '45' });
  assert.equal(cfg.scan.concurrency, 4);
  assert.equal(cfg.scan.thorough, false);
  assert.equal(cfg.scan.maxMinutes, 45);
  assert.equal(cfg.scan.includeLibraries, true);
  assert.equal(cfg.report.name, 'Who can get in');
  assert.equal(cfg.endpoints.graph, 'https://graph.microsoft.com');
  assert.deepEqual(cfg.principals, [{ type: 'group', id: 'X' }, { type: 'user', id: 'a@b.c' }]);
  const bad = async (file, env, re) => assert.rejects(() => buildConfig(file, { ...GOOD, ...env }), re);
  await bad({}, { ACCESS_REPORT_TENANT_ID: '' }, /tenant id is missing/);
  await bad({}, { ACCESS_REPORT_CLIENT_ID: 'nope' }, /client id/);
  await bad({}, { ACCESS_REPORT_CERTIFICATE: '' }, /certificate is missing/);
  await bad({}, { ACCESS_REPORT_FOLDER_URL: '' }, /Where should the report go/);
  await bad({ scna: {} }, {}, /Unknown setting "scna"/);
  await bad({ scan: { concurency: 2 } }, {}, /Unknown setting "scan.concurency"/);
  await bad({ scan: { concurrency: 50 } }, {}, /between 1 and 8/);
  await bad({ scan: { thorough: 'maybe' } }, {}, /true or false/);
  await bad({ report: { name: 'a/b' } }, {}, /report.name/);
  await bad({ principalsFile: '../x.csv' }, {}, /principalsFile/);
  await bad({ sites: { only: ['ftp://x'] } }, {}, /https/);
  await bad({ auth: { mode: 'secret' } }, {}, /certificate.*federated/);
  await bad({ endpoints: { graph: 'http://evil.example' } }, {}, /https/);
  await bad({ principals: [{ type: 'robot', id: 'x' }] }, {}, /principals\[0\]/);
  // Federated needs no certificate; certificate files are read when named.
  const fed = await buildConfig({ auth: { mode: 'federated' } }, { ...GOOD, ACCESS_REPORT_CERTIFICATE: '' });
  assert.equal(fed.auth.certificate, undefined);
  const viaFiles = await buildConfig({ auth: { certificatePath: 'c.pem', privateKeyPath: 'k.pem' } }, { ...GOOD, ACCESS_REPORT_CERTIFICATE: '' }, async (p) => `<${p}>`);
  assert.equal(viaFiles.auth.certificate, '<c.pem>\n<k.pem>');
  await assert.rejects(() => buildConfig({ auth: { certificatePath: '/nope' } }, { ...GOOD, ACCESS_REPORT_CERTIFICATE: '' }), /cannot read/);
  const split = await buildConfig({ report: { siteUrl: 'https://contoso.sharepoint.com/sites/a/sub', folder: '/Docs/Reports/Access/' } }, { ...GOOD, ACCESS_REPORT_FOLDER_URL: '' });
  assert.deepEqual([split.report.siteUrl, split.report.libraryPath, split.report.folder], ['https://contoso.sharepoint.com/sites/a/sub', '/sites/a/sub/Docs', 'Reports/Access']);
});

test('site patterns', () => {
  const keep = siteFilter(['https://c.sharepoint.com/sites/hr*', '*/teams/*'], ['*-archive']);
  assert.equal(keep('https://c.sharepoint.com/sites/HR-payroll/'), true);
  assert.equal(keep('https://c.sharepoint.com/sites/hr-archive'), false);
  assert.equal(keep('https://c.sharepoint.com/teams/x'), true);
  assert.equal(keep('https://c.sharepoint.com/sites/it'), false);
  assert.equal(siteFilter([], [])('anything'), true);
  assert.equal(siteFilter(['https://c.sharepoint.com/sites/a.b'], [])('https://c.sharepoint.com/sites/aXb'), false, 'a dot is a dot');
});

test('paths and OData literals', () => {
  assert.equal(serverRelative('https://c.sharepoint.com/sites/a%20b/'), '/sites/a b');
  assert.equal(serverRelative('https://c.sharepoint.com'), '/');
  assert.equal(joinPath('/sites/a/', '/Docs', 'x y'), '/sites/a/Docs/x y');
  assert.equal(joinPath('/', 'Docs'), '/Docs');
  assert.equal(alias('p', "/sites/o'brien/Docs & more#1"), "@p='%2Fsites%2Fo''brien%2FDocs%20%26%20more%231'");
});

test('certificate sign-in assertion is a valid, short-lived RS256 JWT', async () => {
  const c = testCert();
  const cred = new CertificateCredential('client-1', c.key + '\n' + c.cert);
  const jwt = await cred.assertion('https://login.example/t/oauth2/v2.0/token');
  const [h, p, s] = jwt.split('.');
  const header = JSON.parse(Buffer.from(h, 'base64url'));
  const payload = JSON.parse(Buffer.from(p, 'base64url'));
  assert.equal(header.alg, 'RS256');
  assert.equal(Buffer.from(header.x5t, 'base64url').toString('hex').toUpperCase(), cred.thumbprint);
  const fp = execFileSync('openssl', ['x509', '-noout', '-fingerprint', '-sha1', '-in', join(c.dir, 'c.pem')], { encoding: 'utf8' });
  assert.equal(fp.split('=')[1].trim().replace(/:/g, ''), cred.thumbprint, 'thumbprint as openssl and the Entra portal show it');
  assert.ok(createVerify('RSA-SHA256').update(`${h}.${p}`).verify(createPublicKey(c.cert), Buffer.from(s, 'base64url')));
  assert.equal(payload.aud, 'https://login.example/t/oauth2/v2.0/token');
  assert.equal(payload.iss, 'client-1');
  assert.equal(payload.sub, 'client-1');
  assert.ok(payload.exp - payload.nbf <= 600 && payload.exp > Date.now() / 1000);
  assert.notEqual((await cred.assertion('x')).split('.')[1], (await cred.assertion('x')).split('.')[1], 'a new jti every time');
  assert.ok(cred.daysLeft() > 390 && cred.daysLeft() <= 400);
});

test('certificate input forms and mistakes', () => {
  const c = testCert();
  const pem = c.cert + c.key;
  for (const form of [pem, pem.replace(/\n/g, '\\n'), Buffer.from(pem).toString('base64'), pem.replace(/\n/g, '\r\n'), '  ' + pem + '  ']) {
    assert.equal(new CertificateCredential('x', form).thumbprint.length, 40);
  }
  assert.throws(() => new CertificateCredential('x', c.cert), /No private key found/);
  assert.throws(() => new CertificateCredential('x', c.key), /No certificate found/);
  assert.throws(() => new CertificateCredential('x', 'garbage'), /No certificate found/);
  const dir = mkdtempSync(join(tmpdir(), 'ar-k-'));
  execFileSync('openssl', ['genrsa', '-out', join(dir, 'other.pem'), '2048'], { stdio: 'ignore' });
  assert.throws(() => new CertificateCredential('x', c.cert + readFileSync(join(dir, 'other.pem'), 'utf8')), /does not belong to the certificate/);
  execFileSync('openssl', ['pkcs8', '-topk8', '-in', join(c.dir, 'k.pem'), '-out', join(dir, 'enc.pem'), '-passout', 'pass:secret', '-v2', 'aes-256-cbc'], { stdio: 'ignore' });
  const enc = readFileSync(join(dir, 'enc.pem'), 'utf8');
  assert.throws(() => new CertificateCredential('x', c.cert + enc), /encrypted/);
  assert.equal(new CertificateCredential('x', c.cert + enc, 'secret').thumbprint.length, 40);
  // A .pfx converted with openssl (bag attributes and all) works as it comes out.
  try {
    execFileSync('openssl', ['pkcs12', '-export', '-in', join(c.dir, 'c.pem'), '-inkey', join(c.dir, 'k.pem'), '-out', join(dir, 'x.pfx'), '-passout', 'pass:', '-name', 'friendly'], { stdio: 'ignore' });
    execFileSync('openssl', ['pkcs12', '-in', join(dir, 'x.pfx'), '-nodes', '-passin', 'pass:', '-out', join(dir, 'fromp12.pem')], { stdio: 'ignore' });
    const converted = readFileSync(join(dir, 'fromp12.pem'), 'utf8');
    assert.match(converted, /Bag Attributes/);
    assert.equal(new CertificateCredential('x', converted).thumbprint, new CertificateCredential('x', pem).thumbprint);
  } catch (e) {
    if (e.code === 'ERR_ASSERTION') throw e;
  }
  assert.equal(normalisePem(''), '');
});

test('sign-in errors name the fix and never echo secrets', () => {
  const e = new HttpError(401, "AADSTS700027: The certificate with identifier used to sign the client assertion is not registered on application.", 'u', '');
  assert.match(explainSignInError(e, 'certificate'), /AADSTS700027.*Upload the \.cer/);
  assert.match(explainSignInError(new HttpError(401, 'AADSTS70021: No matching federated identity record', 'u'), 'federated'), /subject/);
  assert.match(explainSignInError(new Error('boom'), 'certificate'), /Check the client id/);
  assert.deepEqual(tokenRoles('a.' + Buffer.from('{"roles":["A","B"]}').toString('base64url') + '.c'), ['A', 'B']);
  assert.deepEqual(tokenRoles('garbage'), []);
});

test('HTTP: retry rules, Retry-After, concurrency ceiling, no redirects followed', async (t) => {
  assert.equal(parseRetryAfter('7'), 7000);
  assert.equal(parseRetryAfter(new Date(Date.now() + 5000).toUTCString()) > 3000, true);
  assert.equal(parseRetryAfter(null), undefined);
  assert.equal(parseRetryAfter('soon'), undefined);
  assert.equal(errorMessageFrom('{"odata.error":{"message":{"value":"Nope"}}}', 403), 'Nope');
  assert.equal(errorMessageFrom('{"error":{"code":"x","message":"Graph says no"}}', 403), 'Graph says no');
  assert.equal(errorMessageFrom('{"error":"invalid_client","error_description":"AADSTS1: bad\\r\\nTrace"}', 400), 'AADSTS1: bad');
  assert.equal(errorMessageFrom('', 502), 'HTTP 502');
  assert.equal(errorMessageFrom('<html>oops</html>', 500), '<html>oops</html>');

  let active = 0;
  let peak = 0;
  const hits = {};
  const server = createServer(async (req, res) => {
    hits[req.url] = (hits[req.url] || 0) + 1;
    active++;
    peak = Math.max(peak, active);
    await new Promise((r) => setTimeout(r, 15));
    active--;
    const n = hits[req.url];
    if (req.url === '/flaky' && n < 3) return res.writeHead(503, { 'Retry-After': '0' }).end();
    if (req.url === '/always429') return res.writeHead(429).end('{"error":{"message":"slow down"}}');
    if (req.url === '/404') return res.writeHead(404).end('{"error":{"message":"gone"}}');
    if (req.url === '/500') return res.writeHead(500).end('x');
    if (req.url === '/redirect') return res.writeHead(302, { Location: 'http://example.invalid/steal' }).end();
    if (req.url === '/badjson') return res.writeHead(200).end(n < 2 ? '{"trunc' : '{"ok":true}');
    if (req.url === '/post' && n < 2) return res.writeHead(503).end();
    res.writeHead(200, { 'Content-Type': 'application/json' }).end('{"ok":true}');
  });
  await new Promise((r) => server.listen(0, '127.0.0.1', r));
  t.after(() => server.close());
  const base = `http://127.0.0.1:${server.address().port}`;
  const http = new Http({ concurrency: 3, baseDelayMs: 2, maxRetries: 3 });
  assert.deepEqual((await http.request(base + '/flaky')).data, { ok: true });
  assert.equal(hits['/flaky'], 3);
  await assert.rejects(() => http.request(base + '/always429'), (e) => e.status === 429 && e.message === 'slow down');
  assert.equal(hits['/always429'], 4, '1 try + 3 retries');
  await assert.rejects(() => http.request(base + '/404'), (e) => e.status === 404);
  assert.equal(hits['/404'], 1, 'client errors are not repeated');
  await assert.rejects(() => http.request(base + '/500'), (e) => e.status === 500);
  assert.equal(hits['/500'], 3, '500 is tried three times, no more');
  await assert.rejects(() => http.request(base + '/redirect'), (e) => e.status === 302);
  assert.equal(hits['/redirect'], 1);
  assert.deepEqual((await http.request(base + '/badjson')).data, { ok: true });
  // 'throttled' retry mode: a refused request is repeated, a broken connection is not.
  assert.equal((await http.request(base + '/post', { method: 'POST', retry: 'throttled', as: 'none' })).status, 200);
  await assert.rejects(() => http.request('http://127.0.0.1:1/x', { retry: 'throttled' }), /Network error/);
  await assert.rejects(() => http.request(base + '/ok', { timeoutMs: 1, retry: false }), /Network error/);
  // One throttled request makes every other request wait too.
  const arrivals = [];
  const slow = createServer((req, res) => {
    arrivals.push([req.url, Date.now()]);
    if (req.url === '/a' && arrivals.filter((x) => x[0] === '/a').length === 1) return res.writeHead(429, { 'Retry-After': '1' }).end();
    res.writeHead(200).end('{}');
  });
  await new Promise((r) => slow.listen(0, '127.0.0.1', r));
  t.after(() => slow.close());
  const sbase = `http://127.0.0.1:${slow.address().port}`;
  const h2 = new Http({ concurrency: 3, baseDelayMs: 2 });
  const a = h2.request(sbase + '/a');
  await new Promise((r) => setTimeout(r, 80));
  await Promise.all([a, h2.request(sbase + '/b')]);
  const refusedAt = arrivals[0][1];
  assert.ok(arrivals.find((x) => x[0] === '/b')[1] - refusedAt >= 900, 'the other request waited for the pause');
  assert.equal(h2.stats.throttled, 1);

  // Let the server finish the request the client gave up on.
  await new Promise((r) => setTimeout(r, 60));
  peak = 0;
  await Promise.all(Array.from({ length: 12 }, (_, i) => http.request(base + '/ok' + i)));
  assert.equal(peak, 3, 'never more than `concurrency` requests in flight');
});

test('log masking', () => {
  const s = redact('User alice.o-brien+x@contoso.com at https://contoso.sharepoint.com/sites/Finance%20Team?x=1 (i:0#.f|membership|alice@contoso.com, c:0t.c|tenant|11111111-2222-3333-4444-555555555555) id 11111111-2222-3333-4444-555555555555 done');
  assert.equal(s, 'User <e-mail> at <address> (<login>, <login>) id <id> done');
});

test('Excel workbook: valid for Excel-compatible readers, awkward text survives', (t) => {
  assert.equal(columnName(0), 'A');
  assert.equal(columnName(25), 'Z');
  assert.equal(columnName(26), 'AA');
  assert.equal(columnName(701), 'ZZ');
  assert.equal(columnName(702), 'AAA');
  assert.equal(crc32(Buffer.from('123456789')), 0xcbf43926);
  assert.deepEqual(sheetNames(['A/B', 'a b', "'x'", '', 'Sheet', 'x'.repeat(40), 'x'.repeat(40)]), ['A B', 'a b (2)', 'x', 'Sheet', 'Sheet (2)', 'x'.repeat(31), 'x'.repeat(27) + ' (2)']);
  assert.equal(xmlText('a<b>&"c"\u0001\uD800z'), 'a&lt;b&gt;&amp;&quot;c&quot;z');
  assert.equal(xmlText('x'.repeat(40000)).length, 32767);

  const nasty = ['=1+1', '  leading and trailing  ', 'line1\nline2', '<tag attr="1">&amp;</tag>', 'emoji 😀 and 你好 and ümlaut', "it's", '\ttab', '0123', '1e5', 'TRUE'];
  const data = buildXlsx([
    { name: 'Data', widths: [20, 12, 12], header: true, rows: [['Text', 'Number', 'Empty'], ...nasty.map((s, i) => [s, i * 1.5, null])] },
    { name: 'Second: sheet?', widths: [10], rows: [['only']], bold: [0] },
    { name: 'Empty', widths: [10, 10], header: true, rows: [['a', 'b']] }
  ]);
  const dir = mkdtempSync(join(tmpdir(), 'ar-xlsx-'));
  const file = join(dir, 'book.xlsx');
  writeFileSync(file, data);

  // 1. The ZIP container is sound.
  let zipOk;
  try {
    zipOk = execFileSync('python3', ['-c', 'import zipfile,sys\nz=zipfile.ZipFile(sys.argv[1])\nprint(z.testzip(), len(z.namelist()))', file], { encoding: 'utf8' }).trim();
  } catch {
    t.diagnostic('python3 not available; ZIP not cross-checked');
  }
  if (zipOk) assert.equal(zipOk, 'None 8');

  // 2. openpyxl reads back exactly what was written.
  const py = `import openpyxl,json,sys
wb=openpyxl.load_workbook(sys.argv[1])
ws=wb['Data']
print(json.dumps({'names':wb.sheetnames,'rows':[[c for c in r] for r in ws.iter_rows(values_only=True)],'freeze':ws.freeze_panes,'filter':ws.auto_filter.ref,'bold':ws['A1'].font.b,'plain':ws['A2'].font.b,'types':[ws.cell(row=i,column=1).data_type for i in range(2,12)]}))`;
  let got;
  try {
    got = JSON.parse(execFileSync('python3', ['-c', py, file], { encoding: 'utf8' }));
  } catch (e) {
    t.diagnostic('openpyxl not available; workbook not cross-checked: ' + String(e.message).split('\n')[0]);
  }
  if (got) {
    assert.deepEqual(got.names, ['Data', 'Second  sheet', 'Empty']);
    assert.deepEqual(got.rows.slice(1).map((r) => r[0]), nasty);
    assert.deepEqual(got.rows.slice(1).map((r) => r[1]), nasty.map((_, i) => i * 1.5));
    assert.equal(got.freeze, 'A2');
    assert.equal(got.filter, 'A1:C11');
    assert.equal(got.bold, true);
    assert.equal(got.plain, false);
    assert.ok(got.types.every((x) => x === 'inlineStr' || x === 's'), 'text stays text: no formulas, no numbers guessed from text');
  }

  // 3. LibreOffice opens it and converts it (a stricter reader than openpyxl).
  try {
    execFileSync('soffice', ['--headless', '--convert-to', 'csv:Text - txt - csv (StarCalc):44,34,76,1', '--outdir', dir, file], { stdio: 'ignore', timeout: 120000, env: { ...process.env, HOME: dir } });
    const back = parseCsv(readFileSync(join(dir, 'book.csv'), 'utf8'));
    assert.equal(back.length, 11);
    assert.equal(back[1][0], '=1+1', 'still text after a round trip through LibreOffice');
    assert.equal(back[5][0], 'emoji 😀 and 你好 and ümlaut');
    assert.equal(back[8][0], '0123');
  } catch (e) {
    if (e.code === 'ERR_ASSERTION') throw e;
    t.diagnostic('LibreOffice not available; skipped that reader');
  }
  assert.throws(() => buildXlsx([]), /at least one sheet/);
});

test('changes between two reports', () => {
  const row = (p, url, access, extra = {}) => ({ p, siteUrl: 'https://c/sites/' + url.split('/')[0], siteTitle: 't', object: url.includes('/') ? 'Library' : 'Site', title: url, url: 'https://c/sites/' + url, access, levels: access, how: 'Direct', unique: false, note: '', ...extra });
  const prev = { v: 1, period: '2026-10', generatedAt: '', principals: [{ key: 'u1', label: 'U1', kind: 'user' }, { key: 'gone', label: 'Gone', kind: 'user' }], failedSites: ['https://c/sites/wasdown'], rows: [row('u1', 'a', 'Read'), row('u1', 'b', 'Edit'), row('u1', 'c', 'Read'), row('u1', 'down', 'Read'), row('u1', 'same', 'Read'), row('u1', 'how', 'Read'), row('gone', 'a', 'Read')] };
  const cur = { v: 1, period: '2026-11', generatedAt: '', principals: [{ key: 'u1', label: 'U1', kind: 'user' }, { key: 'new', label: 'New', kind: 'user' }], failedSites: ['https://c/sites/down'], rows: [row('u1', 'a', 'Edit'), row('u1', 'b', 'Read'), row('u1', 'd', 'Read'), row('u1', 'wasdown', 'Read'), row('u1', 'same', 'Read', { title: 'renamed', note: 'x' }), row('u1', 'how', 'Read', { how: 'SharePoint group: V' }), row('new', 'a', 'Read')] };
  const got = diffSnapshots(prev, cur).map((c) => `${c.change}:${c.row.p}:${c.row.url.split('/').pop()}`).sort();
  assert.deepEqual(got, ['Added:u1:d', 'Changed:u1:how', 'Less access:u1:b', 'More access:u1:a', 'Removed:u1:c']);
  assert.deepEqual(unpackSnapshot(packSnapshot(cur)), cur);
  assert.throws(() => unpackSnapshot(packSnapshot({ v: 2 })), /Not a snapshot/);

  const files = buildReport({ name: 'Access report', period: '2026-11', generatedAt: new Date('2026-11-01T02:00:00Z'), principals: cur.principals, rows: cur.rows, failedSites: [{ url: 'https://c/sites/down', title: 'Down', error: 'locked' }], unresolved: [], warnings: ['w1'], sitesTotal: 9, previous: prev, scope: 's', maxSheetRows: 3 });
  assert.equal(files.counts.added, 1);
  assert.equal(files.counts.removed, 1);
  assert.equal(files.counts.changed, 3);
  assert.match(files.changesCsv, /More access,U1/);
  assert.equal(parseCsv(files.accessCsv).length, 8, 'the CSV always has every row');
  assert.ok(files.workbook.length > 1000);
});

test('command line arguments and summaries', () => {
  assert.deepEqual(parseArgs(['run', '--config=c.json', '--period', '2026-11', '--force', '--no-upload', '--out', 'o', '--limit=5', '-v', '--json']), { command: 'run', config: 'c.json', period: '2026-11', out: 'o', limit: 5, force: true, upload: false, verbose: true, json: true });
  assert.throws(() => parseArgs(['run', '--limit', 'x']), /whole number/);
  assert.throws(() => parseArgs(['run', '--config']), /needs a value/);
  assert.throws(() => parseArgs(['run', 'check']), /Unexpected/);
  assert.equal(parseArgs(['--help']).command, 'help');
  assert.equal(currentPeriod(Date.parse('2026-12-31T23:59:59Z')), '2026-12');
  assert.equal(currentPeriod(Date.parse('2027-01-01T00:00:00Z')), '2027-01');
  const md = stepSummary({ status: 'partial', period: '2026-11', sitesTotal: 10, sitesDone: 4, sitesFailed: 1, principals: 2, unresolved: 1, rows: 0, added: 0, removed: 0, changed: 0, files: [], warnings: [], requests: 99, throttled: 2 });
  assert.match(md, /not finished yet/);
  assert.match(md, /4 of 10/);
  assert.match(md, /next scheduled run carries on/);
});

test('large files go up in pieces and arrive intact; a failed piece restarts the upload', async (t) => {
  const c = testCert();
  const tenant = await new MockTenant({ users: [], groups: [], sites: [{ path: '/sites/gov', title: 'G', web: { path: '/sites/gov', title: 'G', assignments: [] } }] }, { certPem: c.cert }).start();
  t.after(() => tenant.stop());
  tenant.addLibrary('/sites/gov/Docs');
  const { TokenProvider } = await import('../dist/auth.js');
  const http = new Http({ baseDelayMs: 2 });
  const tokens = new TokenProvider(http, tenant.loginBase, tenant.opts.tenantId, tenant.opts.clientId, new CertificateCredential(tenant.opts.clientId, c.cert + c.key));
  const sp = new SharePoint(http, tokens, { chunkThreshold: 1000, chunkSize: 400 });
  const site = tenant.origin + '/sites/gov';
  const data = Buffer.alloc(1801);
  for (let i = 0; i < data.length; i++) data[i] = (i * 31) % 251;

  await sp.upload(site, '/sites/gov/Docs', 'big file.bin', data);
  assert.ok(tenant.file('/sites/gov/Docs/big file.bin').equals(data));
  assert.equal(tenant.count((r) => /startupload/i.test(r.path)), 1);
  assert.equal(tenant.count((r) => /continueupload/i.test(r.path)), 3);
  assert.equal(tenant.count((r) => /finishupload/i.test(r.path)), 1);

  // Exactly one piece long (above the threshold only because the threshold is tiny here).
  const one = new SharePoint(http, tokens, { chunkThreshold: 10, chunkSize: 400 });
  await one.upload(site, '/sites/gov/Docs', 'one.bin', data.subarray(0, 400));
  assert.ok(tenant.file('/sites/gov/Docs/one.bin').equals(data.subarray(0, 400)));

  // A piece is lost on the way: the session is cancelled and the whole file is sent again.
  tenant.requests.length = 0;
  let dropped = false;
  tenant.faults.push((r) => {
    if (!dropped && /continueupload/i.test(r.path)) {
      dropped = true;
      return { status: 500, once: true };
    }
    return undefined;
  });
  await sp.upload(site, '/sites/gov/Docs', 'retry.bin', data);
  assert.ok(tenant.file('/sites/gov/Docs/retry.bin').equals(data));
  assert.equal(tenant.count((r) => /cancelupload/i.test(r.path)), 1);
  assert.equal(tenant.count((r) => /startupload/i.test(r.path)), 2);

  // Small files: one request; download, list and delete work; missing files are undefined, not errors.
  await sp.upload(site, '/sites/gov/Docs', "o'brien & co #1.txt", Buffer.from('hi'));
  assert.equal(Buffer.from(await sp.download(site, "/sites/gov/Docs/o'brien & co #1.txt")).toString(), 'hi');
  assert.equal(await sp.download(site, '/sites/gov/Docs/nope.txt'), undefined);
  assert.ok((await sp.listFiles(site, '/sites/gov/Docs')).some((f) => f.name === "o'brien & co #1.txt" && f.length === 2));
  await sp.deleteFile(site, "/sites/gov/Docs/o'brien & co #1.txt");
  await sp.deleteFile(site, "/sites/gov/Docs/o'brien & co #1.txt");
  assert.equal(await sp.download(site, "/sites/gov/Docs/o'brien & co #1.txt"), undefined);
  assert.deepEqual(await sp.listFiles(site, '/sites/gov/Nope'), []);
  assert.equal(await sp.ensureFolder(site, '/sites/gov/Docs', 'a/b c/d'), '/sites/gov/Docs/a/b c/d');
  assert.equal(await sp.folderExists(site, '/sites/gov/Docs/a/b c'), true);
});
