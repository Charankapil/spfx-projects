import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { gunzipSync } from 'node:zlib';
import { buildConfig } from '../dist/config.js';
import { createLogger } from '../dist/log.js';
import { createClients, run } from '../dist/runner.js';
import { MockTenant } from './mock-tenant.mjs';

let cached;
/** A throw-away self-signed certificate + key (PEM), made once per test process. */
export function testCert() {
  if (!cached) {
    const dir = mkdtempSync(join(tmpdir(), 'ar-cert-'));
    execFileSync('openssl', ['req', '-x509', '-newkey', 'rsa:2048', '-nodes', '-keyout', join(dir, 'k.pem'), '-out', join(dir, 'c.pem'), '-days', '400', '-subj', '/CN=access-report-test'], { stdio: 'ignore' });
    cached = { cert: readFileSync(join(dir, 'c.pem'), 'utf8'), key: readFileSync(join(dir, 'k.pem'), 'utf8'), dir };
  }
  return cached;
}

export const ID = {
  alice: '00000000-0000-0000-0000-00000000a11c',
  bob: '00000000-0000-0000-0000-000000000b0b',
  carol: '00000000-0000-0000-0000-00000000ca01',
  dave: '00000000-0000-0000-0000-00000000da7e',
  erin: '00000000-0000-0000-0000-0000000e4100',
  fin: '10000000-0000-0000-0000-00000000f100',
  staff: '10000000-0000-0000-0000-0000000a1100',
  mkt: '10000000-0000-0000-0000-000000003c70'
};

/** A small tenant that exercises every route to access. */
export function fixture() {
  return {
    users: [
      { id: ID.alice, upn: 'alice@contoso.com', displayName: 'Alice Adams' },
      { id: ID.bob, upn: 'bob@contoso.com', displayName: "Bob O'Brien" },
      { id: ID.carol, upn: 'carol_fabrikam.com#ext#@contoso.onmicrosoft.com', mail: 'carol@fabrikam.com', displayName: 'Carol (Fabrikam)', userType: 'Guest' },
      { id: ID.dave, upn: 'dave@contoso.com', displayName: 'Dave Dunn' },
      { id: ID.erin, upn: 'erin@contoso.com', displayName: 'Erin Eze' }
    ],
    groups: [
      { id: ID.fin, displayName: 'Finance', members: [ID.alice] },
      { id: ID.staff, displayName: 'All staff', members: [ID.fin, ID.bob] },
      { id: ID.mkt, displayName: 'Marketing', mail: 'marketing@contoso.com', unified: true, members: [ID.bob], owners: [ID.alice] }
    ],
    sites: [
      {
        path: '',
        title: 'Home',
        spGroups: [{ title: 'Home Visitors', members: ['everyone'] }],
        web: { path: '', title: 'Home', assignments: [{ to: 'sp:Home Visitors', roles: ['View Only'] }], lists: [{ id: 'aaaaaaaa-0000-0000-0000-000000000001', title: 'Site Pages', path: '/SitePages' }] }
      },
      {
        path: '/sites/alpha',
        title: 'Alpha =team',
        admins: ['user:erin@contoso.com'],
        spGroups: [
          { title: 'Alpha Owners', members: ['owners:' + ID.mkt] },
          { title: 'Alpha Members', members: ['group:' + ID.mkt] },
          { title: 'Alpha Visitors', members: ['group:' + ID.staff] }
        ],
        web: {
          path: '/sites/alpha',
          title: 'Alpha =team',
          assignments: [
            { to: 'sp:Alpha Owners', roles: ['Full Control'] },
            { to: 'sp:Alpha Members', roles: ['Edit'] },
            { to: 'sp:Alpha Visitors', roles: ['Read'] },
            { to: 'user:carol_fabrikam.com#ext#@contoso.onmicrosoft.com', roles: ['Limited Access'] }
          ],
          lists: [
            { id: 'aaaaaaaa-0000-0000-0000-00000000000a', title: 'Documents', path: '/sites/alpha/Shared Documents' },
            {
              id: 'aaaaaaaa-0000-0000-0000-00000000000b',
              title: 'Contracts',
              path: '/sites/alpha/Contracts',
              assignments: [
                { to: 'group:' + ID.fin, roles: ['Contribute'] },
                { to: 'user:carol_fabrikam.com#ext#@contoso.onmicrosoft.com', roles: ['Read'] }
              ]
            },
            { id: 'aaaaaaaa-0000-0000-0000-00000000000c', title: 'Hidden lib', path: '/sites/alpha/_catalogs/x', hidden: true },
            { id: 'aaaaaaaa-0000-0000-0000-00000000000d', title: 'Site Assets', path: '/sites/alpha/SiteAssets', system: true },
            { id: 'aaaaaaaa-0000-0000-0000-00000000000e', title: 'Tasks', path: '/sites/alpha/Lists/Tasks', baseType: 0 }
          ],
          webs: [
            {
              path: '/sites/alpha/hr',
              title: 'HR',
              assignments: [{ to: 'user:bob@contoso.com', roles: ['Full Control'] }],
              lists: [{ id: 'aaaaaaaa-0000-0000-0000-00000000001a', title: 'Policies', path: '/sites/alpha/hr/Policies' }],
              webs: [{ path: '/sites/alpha/hr/payroll', title: 'Payroll', lists: [{ id: 'aaaaaaaa-0000-0000-0000-00000000002a', title: 'Runs', path: '/sites/alpha/hr/payroll/Runs', assignments: [{ to: 'group:' + ID.fin, roles: ['Read'] }] }] }]
            },
            { path: '/sites/alpha/news', title: 'News' }
          ]
        }
      },
      {
        path: '/sites/beta',
        title: 'Beta',
        web: {
          path: '/sites/beta',
          title: 'Beta',
          assignments: [
            { to: 'eeeu', roles: ['Read'] },
            { to: 'user:bob@contoso.com', roles: ['Design'] }
          ],
          lists: [{ id: 'bbbbbbbb-0000-0000-0000-00000000000a', title: 'Documents', path: '/sites/beta/Shared Documents' }]
        }
      },
      { path: '/sites/delta', title: 'Delta', visitors: ['user:dave@contoso.com'], web: { path: '/sites/delta', title: 'Delta', assignments: [{ to: 'user:zed@contoso.com', roles: ['Full Control'] }], lists: [{ id: 'dddddddd-0000-0000-0000-00000000000a', title: 'Documents', path: '/sites/delta/Shared Documents' }] } },
      { path: '/sites/gamma', title: 'Gamma (locked)', locked: true, web: { path: '/sites/gamma', title: 'Gamma', assignments: [{ to: 'eeeu', roles: ['Read'] }] } },
      { path: '/personal/alice_contoso_com', title: 'Alice', personal: true, web: { path: '/personal/alice_contoso_com', title: 'Alice', assignments: [{ to: 'user:alice@contoso.com', roles: ['Full Control'] }] } }
    ]
  };
}

export const ALL_PRINCIPALS = ['user:alice@contoso.com', 'user:bob@contoso.com', 'user:carol@fabrikam.com', 'user:dave@contoso.com', 'user:erin@contoso.com', 'group:Finance', 'group:' + ID.mkt, 'group:All staff'];
/** The same people as the oracle names them. */
export const ORACLE_PRINCIPALS = ['user:alice@contoso.com', 'user:bob@contoso.com', 'user:carol_fabrikam.com#ext#@contoso.onmicrosoft.com', 'user:dave@contoso.com', 'user:erin@contoso.com', 'group:' + ID.fin, 'group:' + ID.mkt, 'group:' + ID.staff];
export const ORACLE_KEYS = ['user:' + ID.alice, 'user:' + ID.bob, 'user:' + ID.carol, 'user:' + ID.dave, 'user:' + ID.erin, 'group:' + ID.fin, 'group:' + ID.mkt, 'group:' + ID.staff];

export async function startTenant(model = fixture(), opts = {}) {
  const c = testCert();
  const t = await new MockTenant(model, { certPem: c.cert, ...opts }).start();
  t.addLibrary('/sites/gov/Shared Documents');
  // The report site itself.
  if (!model.sites.some((s) => s.path === '/sites/gov')) {
    model.sites.push({ path: '/sites/gov', title: 'Governance', hiddenFromGraph: opts.hideReportSite !== false, web: { path: '/sites/gov', title: 'Governance', assignments: [] } });
    t.index();
  }
  return t;
}

export function envFor(t, extra = {}) {
  const c = testCert();
  return {
    ACCESS_REPORT_TENANT_ID: t.opts.tenantId,
    ACCESS_REPORT_CLIENT_ID: t.opts.clientId,
    ACCESS_REPORT_FOLDER_URL: `${t.origin}/sites/gov/Shared Documents/Access reports`,
    ACCESS_REPORT_CERTIFICATE: c.cert + '\n' + c.key,
    ACCESS_REPORT_LOGIN_URL: t.loginBase,
    ACCESS_REPORT_GRAPH_URL: t.graphBase,
    ...extra
  };
}

/** Runs the tool against a mock tenant. Returns { result, lines, error }. */
export async function runTool(t, { file = {}, env = {}, opts = {}, verbose = false } = {}) {
  const lines = [];
  const log = createLogger(verbose, (l) => lines.push(l));
  const fullEnv = envFor(t, env);
  const cfg = await buildConfig({ principals: ALL_PRINCIPALS, ...file }, fullEnv);
  const clients = createClients(cfg, log, fullEnv, { baseDelayMs: 5, maxRetries: opts.maxRetries ?? 5 });
  delete opts.maxRetries;
  try {
    const result = await run(cfg, clients, log, opts);
    return { result, lines, clients, cfg };
  } catch (error) {
    return { error, lines, clients, cfg };
  }
}

export const FOLDER = '/sites/gov/Shared Documents/Access reports';

export function snapshotOf(t, period) {
  const data = t.file(`${FOLDER}/_state/snapshot-${period}.json.gz`);
  return data ? JSON.parse(gunzipSync(data).toString('utf8')) : undefined;
}

/** Rows reduced to what the oracle speaks about, sorted, for comparison. */
export function comparable(rows, keyToOracle) {
  return rows
    .map((r) => `${keyToOracle ? keyToOracle(r.p) : r.p} | ${r.object} | ${r.url} | ${r.access}`)
    .sort();
}

export function oracleComparable(t, oraclePrincipals, keys, scope) {
  const map = new Map(oraclePrincipals.map((p, i) => [p, keys[i]]));
  return t
    .oracle(oraclePrincipals, scope)
    .map((r) => `${map.get(r.who)} | ${r.object} | ${r.url} | ${r.access}`)
    .sort();
}
