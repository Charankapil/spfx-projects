import { test } from 'node:test';
import * as assert from 'node:assert/strict';
import { ApiClient, ApiError, IHttpResponseLike, IHttpTransport, mapWithConcurrency } from '../src/webparts/platformHub/app/http/ApiClient';
import { classifyTrigger, mapApp, mapDlpV2, mapEnvironment, mapFlow } from '../src/webparts/platformHub/app/services/PowerPlatformService';
import { buildMakers } from '../src/webparts/platformHub/app/providers/LiveDataProvider';
import { SampleDataProvider } from '../src/webparts/platformHub/app/providers/SampleDataProvider';

const res = (status: number, body: unknown, headers: Record<string, string> = {}): IHttpResponseLike => ({
  status, ok: status >= 200 && status < 300,
  headers: { get: (n: string) => headers[n] ?? headers[n.toLowerCase()] ?? null },
  json: async () => body, text: async () => JSON.stringify(body)
});

test('retries throttled requests using Retry-After', async () => {
  let calls = 0;
  const t: IHttpTransport = { fetch: async () => (++calls < 3 ? res(429, { error: { message: 'slow down' } }, { 'Retry-After': '0' }) : res(200, { value: [1] })) };
  const data = await new ApiClient(t).getJson<{ value: number[] }>('r', 'u');
  assert.equal(calls, 3);
  assert.deepEqual(data.value, [1]);
});

test('follows nextLink and @odata.nextLink and honours maxItems', async () => {
  const pages: Record<string, unknown> = {
    a: { value: [1, 2], nextLink: 'b' },
    b: { value: [3], '@odata.nextLink': 'c' },
    c: { value: [4, 5] }
  };
  const t: IHttpTransport = { fetch: async (_r, url) => res(200, pages[url]) };
  const api = new ApiClient(t);
  assert.deepEqual(await api.getAll('r', 'a'), [1, 2, 3, 4, 5]);
  assert.deepEqual(await api.getAll('r', 'a', { maxItems: 3 }), [1, 2, 3]);
});

test('classifies failures for the Data sources panel', async () => {
  const forbidden = new ApiClient({ fetch: async () => res(403, { error: { message: 'Insufficient privileges' } }) });
  await assert.rejects(forbidden.getJson('r', 'u'), (e: ApiError) => e.kind === 'forbidden' && /Insufficient/.test(e.message));
  const blocked = new ApiClient({ fetch: async () => { throw new TypeError('Failed to fetch'); } });
  await assert.rejects(blocked.getJson('r', 'u'), (e: ApiError) => e.kind === 'blocked');
  const consent = new ApiClient({ fetch: async () => { throw new Error('AADSTS65001: The user or administrator has not consented'); } });
  await assert.rejects(consent.getJson('r', 'u'), (e: ApiError) => e.kind === 'forbidden' && /API access/.test(e.message));
  const missing = new ApiClient({ fetch: async () => res(404, {}) });
  await assert.rejects(missing.getJson('r', 'u'), (e: ApiError) => e.kind === 'notFound');
});

test('mapWithConcurrency keeps order and limits parallelism', async () => {
  let active = 0; let peak = 0;
  const out = await mapWithConcurrency([1, 2, 3, 4, 5, 6], 2, async (n) => {
    active++; peak = Math.max(peak, active);
    await new Promise((r) => setTimeout(r, 5));
    active--;
    return n * 10;
  });
  assert.deepEqual(out, [10, 20, 30, 40, 50, 60]);
  assert.equal(peak, 2);
});

test('maps BAP environment payloads', () => {
  const e = mapEnvironment({
    name: 'abc', location: 'unitedstates',
    properties: {
      displayName: 'Finance', environmentSku: 'Production', isDefault: false, createdTime: '2025-01-01',
      createdBy: { displayName: 'Ana' },
      linkedEnvironmentMetadata: { instanceUrl: 'https://org.crm.dynamics.com/', version: '9.2', securityGroupId: 'sg' },
      governanceConfiguration: { protectionLevel: 'Standard' },
      states: { runtime: { id: 'Enabled' } },
      capacity: [{ capacityType: 'Database', actualConsumption: 2, capacityUnit: 'GB' }, { capacityType: 'File', actualConsumption: 100, capacityUnit: 'MB' }]
    }
  });
  assert.equal(e.region, 'United States');
  assert.equal(e.orgUrl, 'https://org.crm.dynamics.com');
  assert.ok(e.isManaged && e.hasDataverse);
  assert.equal(e.capacity.databaseMb, 2048);
  assert.equal(e.capacity.fileMb, 100);
});

test('maps apps and flows including connectors and triggers', () => {
  const env = mapEnvironment({ name: 'e1', properties: { displayName: 'E1' } });
  const app = mapApp({
    name: 'app1',
    properties: {
      displayName: 'Expenses', appType: 'ClassicCanvasApp', owner: { id: 'u1', displayName: 'Ana' }, sharedUsersCount: 3,
      connectionReferences: { k1: { id: '/providers/microsoft.powerapps/apis/shared_sql', displayName: 'SQL Server', apiTier: 'Premium' }, k2: { id: '/x/shared_office365' } }
    }
  }, env);
  assert.equal(app.kind, 'Canvas');
  assert.equal(app.connectors.length, 2);
  assert.ok(app.usesPremium);
  const flow = mapFlow({
    name: 'f1',
    properties: {
      displayName: 'Sync', state: 'Suspended', flowSuspensionReason: 'DlpPolicyViolation', creator: { objectId: 'u2' },
      definitionSummary: { triggers: [{ type: 'Recurrence' }], actions: [{ type: 'Http' }, { type: 'OpenApiConnection', api: { name: 'shared_teams' } }] },
      connectionReferences: [{ id: '/x/shared_teams' }]
    }
  }, env);
  assert.equal(flow.state, 'Suspended');
  assert.equal(flow.triggerType, 'Scheduled');
  assert.deepEqual(flow.connectors.map((c) => c.name).sort(), ['http', 'shared_teams']);
  assert.ok(flow.usesPremium, 'the HTTP action is premium');
  assert.equal(classifyTrigger({ type: 'Request', kind: 'Http' }), 'HTTP request');
  assert.equal(classifyTrigger({ type: 'OpenApiConnectionWebhook' }), 'Automated');
});

test('maps v2 DLP policies', () => {
  const p = mapDlpV2({
    name: 'p1', displayName: 'Baseline', environmentType: 'ExceptEnvironments', environments: [{ name: 'e1' }],
    connectorGroups: [{ classification: 'Confidential', connectors: [{ id: '/x/shared_sql', name: 'shared_sql' }] }, { classification: 'Blocked', connectors: [{ name: 'shared_http' }] }],
    defaultConnectorsClassification: 'General'
  });
  assert.equal(p.scope, 'All except');
  assert.deepEqual(p.environmentIds, ['e1']);
  assert.deepEqual(p.businessConnectors, ['shared_sql']);
  assert.deepEqual(p.blockedConnectors, ['shared_http']);
  assert.equal(p.defaultGroup, 'Non-business');
});

test('sample provider produces a consistent inventory', async () => {
  const p = new SampleDataProvider(Date.parse('2026-11-02T09:00:00Z'));
  const inv = await p.getInventory();
  const makers = await p.getMakers();
  assert.ok(inv.environments.length > 5 && inv.apps.length > 100 && inv.flows.length > 100 && inv.agents.length > 0);
  const envIds = new Set(inv.environments.map((e) => e.id));
  assert.ok(inv.apps.every((a) => envIds.has(a.environmentId)));
  const total = makers.makers.reduce((s, m) => s + m.apps + m.flows + m.agents, 0);
  assert.equal(total, inv.apps.length + inv.flows.length + inv.agents.length, 'every resource is attributed to one maker');
  assert.deepEqual(buildMakers(inv, makers.users).length, makers.makers.length);
});
