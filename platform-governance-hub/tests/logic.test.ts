import { test } from 'node:test';
import * as assert from 'node:assert/strict';
import { connectorNameFromId, makeConnector, resolveTier } from '../src/webparts/platformHub/app/logic/connectors';
import { countBy, monthlyBuckets, topN, weeklyBuckets } from '../src/webparts/platformHub/app/logic/aggregate';
import { evaluateGovernance, isCoveredByDlp } from '../src/webparts/platformHub/app/logic/governance';
import { daysSince, fmtBytes } from '../src/webparts/platformHub/app/logic/format';
import { isPowerPlatformSku } from '../src/webparts/platformHub/app/logic/skus';
import { toCsv } from '../src/webparts/platformHub/app/export/tabular';
import { sheetName } from '../src/webparts/platformHub/app/export/excel';
import { IEnvironment, IDlpPolicy, IUserInfo } from '../src/webparts/platformHub/app/models';

const env = (id: string, extra: Partial<IEnvironment> = {}): IEnvironment => ({
  id, displayName: id, type: 'Production', region: 'Europe', isDefault: false, isManaged: false, hasDataverse: false, state: 'Enabled',
  capacity: { databaseMb: 0, fileMb: 0, logMb: 0 }, ...extra
});

test('connector ids and tiers', () => {
  assert.equal(connectorNameFromId('/providers/Microsoft.PowerApps/apis/shared_SharePointOnline'), 'shared_sharepointonline');
  assert.equal(resolveTier('shared_sql'), 'Premium');
  assert.equal(resolveTier('shared_office365'), 'Standard');
  assert.equal(resolveTier('shared_office365', 'Premium'), 'Premium');
  assert.equal(resolveTier('shared_contoso-20api-5f1e3c4b9a8d7e6f12'), 'Custom');
  assert.equal(makeConnector(undefined), undefined);
  assert.equal(makeConnector('/x/shared_teams')!.displayName, 'Microsoft Teams');
});

test('aggregation helpers', () => {
  const c = countBy(['a', 'b', 'a', undefined], (x) => x);
  assert.deepEqual(c[0], { label: 'a', value: 2 });
  assert.equal(c.filter((x) => x.label === 'Unknown')[0].value, 1);
  const t = topN([{ label: 'a', value: 5 }, { label: 'b', value: 4 }, { label: 'c', value: 3 }, { label: 'd', value: 2 }], 3);
  assert.deepEqual(t.map((x) => x.label), ['a', 'b', 'Other']);
  assert.equal(t[2].value, 5);
  const now = Date.parse('2026-11-04T12:00:00Z'); // Wednesday
  const w = weeklyBuckets([['2026-11-03T08:00:00Z', '2026-10-27T08:00:00Z', '2020-01-01T00:00:00Z', undefined]], 4, now);
  assert.equal(w.buckets.length, 4);
  assert.equal(w.values[0].reduce((s, x) => s + x, 0), 2, 'items outside the window are ignored');
  assert.equal(w.values[0][3], 1);
  const m = monthlyBuckets([['2026-11-01T10:00:00Z', '2026-10-15T00:00:00Z', '2026-10-20T00:00:00Z']], 3, now);
  assert.deepEqual(m.values[0], [0, 2, 1]);
});

test('format helpers', () => {
  assert.equal(fmtBytes(0), '0 B');
  assert.equal(fmtBytes(1536), '1.5 KB');
  assert.equal(daysSince(undefined), undefined);
  assert.equal(daysSince('2026-11-01T00:00:00Z', Date.parse('2026-11-04T00:00:00Z')), 3);
});

test('Power Platform SKU detection', () => {
  assert.ok(isPowerPlatformSku('POWERAPPS_PER_USER'));
  assert.ok(isPowerPlatformSku('FLOW_FREE'));
  assert.ok(!isPowerPlatformSku('SPE_E3', ['EXCHANGE_S_ENTERPRISE', 'POWERAPPS_O365_P2']));
});

test('DLP coverage for every scope', () => {
  const pol = (scope: IDlpPolicy['scope'], ids: string[]): IDlpPolicy => ({
    id: 'p', displayName: 'p', scope, environmentIds: ids, businessConnectors: [], nonBusinessConnectors: [], blockedConnectors: [], defaultGroup: 'Non-business'
  });
  assert.ok(isCoveredByDlp(env('a'), [pol('All environments', [])]));
  assert.ok(!isCoveredByDlp(env('a'), [pol('All except', ['a'])]));
  assert.ok(isCoveredByDlp(env('b'), [pol('All except', ['a'])]));
  assert.ok(isCoveredByDlp(env('a'), [pol('Only selected', ['a'])]));
  assert.ok(!isCoveredByDlp(env('b'), [pol('Only selected', ['a'])]));
  assert.ok(!isCoveredByDlp(env('b'), []));
});

test('governance rules', () => {
  const users: Record<string, IUserInfo> = { gone: { id: 'gone', status: 'Deleted' }, ok: { id: 'ok', status: 'Active', displayName: 'Ok' } };
  const base = { environmentId: 'prod', environmentName: 'prod', connectors: [], usesPremium: false, usesCustom: false };
  const findings = evaluateGovernance({
    environments: [env('prod'), env('dev', { type: 'Developer' }), env('def', { isDefault: true, type: 'Default' })],
    apps: [
      { ...base, id: 'a1', displayName: 'Orphan app', kind: 'Canvas', owner: { id: 'gone' } },
      { ...base, id: 'a2', displayName: 'Open app', kind: 'Canvas', owner: { id: 'ok' }, sharedWithEveryone: true, lastModifiedTime: new Date().toISOString() }
    ],
    flows: [
      { ...base, id: 'f1', displayName: 'Off orphan', state: 'Off', owner: { id: 'gone' }, triggerType: 'Scheduled', actionCount: 1 },
      { ...base, id: 'f2', displayName: 'Suspended', state: 'Suspended', owner: { id: 'ok' }, triggerType: 'Scheduled', actionCount: 1 }
    ],
    agents: [
      { id: 'b1', displayName: 'Public bot', environmentId: 'prod', environmentName: 'prod', owner: { id: 'ok' }, authMode: 'None', publishedTime: '2026-01-01', state: 'Active' },
      { id: 'b2', displayName: 'Draft bot', environmentId: 'prod', environmentName: 'prod', owner: { id: 'ok' }, authMode: 'None', state: 'Active' }
    ],
    users,
    dlp: [{ id: 'p', displayName: 'p', scope: 'Only selected', environmentIds: ['prod'], businessConnectors: [], nonBusinessConnectors: [], blockedConnectors: [], defaultGroup: 'x' }],
    groups: [{ id: 'g', displayName: 'Ownerless', visibility: 'Private', isTeam: true, ownerCount: 0 }],
    credentials: [
      { appId: '1', appDisplayName: 'Soon', type: 'Secret', displayName: 's', endDateTime: '', daysLeft: 5 },
      { appId: '2', appDisplayName: 'Ancient', type: 'Secret', displayName: 's', endDateTime: '', daysLeft: -400 }
    ]
  });
  const ids = (rule: string): string[] => findings.filter((f) => f.ruleId === rule).map((f) => f.resourceName);
  assert.deepEqual(ids('orphan-app'), ['Orphan app']);
  assert.deepEqual(ids('orphan-flow'), [], 'turned-off flows are not flagged as orphaned');
  assert.deepEqual(ids('app-everyone'), ['Open app']);
  assert.deepEqual(ids('flow-suspended'), ['Suspended']);
  assert.deepEqual(ids('agent-noauth'), ['Public bot'], 'unpublished agents are not flagged');
  assert.deepEqual(ids('env-nodlp').sort(), ['def', 'dev']);
  assert.deepEqual(ids('env-nosg'), ['prod']);
  assert.deepEqual(ids('env-empty'), ['dev']);
  assert.deepEqual(ids('group-ownerless'), ['Ownerless']);
  assert.deepEqual(ids('cred-expiring'), ['Soon']);
  assert.equal(findings[0].severity, 'High', 'sorted by severity');
});

test('CSV escaping and formula injection protection', () => {
  const csv = toCsv({ name: 't', columns: [{ header: 'A' }, { header: 'B' }], rows: [['=cmd()', 'a,"b"'], [3, true], [undefined, '+1']] });
  const lines = csv.replace('﻿', '').split('\r\n');
  assert.equal(lines[0], 'A,B');
  assert.equal(lines[1], `'=cmd(),"a,""b"""`);
  assert.equal(lines[2], '3,Yes');
  assert.equal(lines[3], `,'+1`);
});

test('Excel sheet names are valid and unique', () => {
  const used: Record<string, boolean> = {};
  assert.equal(sheetName('Apps: [all]?', used), 'Apps   all');
  const long = 'A very long sheet name that exceeds the limit';
  const a = sheetName(long, used);
  const b = sheetName(long, used);
  assert.equal(a.length, 31);
  assert.ok(b.length <= 31 && b !== a && b.endsWith('(2)'));
});
