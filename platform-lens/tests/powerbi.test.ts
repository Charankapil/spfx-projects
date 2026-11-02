import { test } from 'node:test';
import * as assert from 'node:assert/strict';
import { ApiClient, IHttpResponseLike } from '../src/webparts/platformHub/app/http/ApiClient';
import { IPbiContent, mapWorkspace, PowerBiService } from '../src/webparts/platformHub/app/services/PowerBiService';
import { activityWindows, PbiActivityAggregator } from '../src/webparts/platformHub/app/logic/pbiActivity';
import { mergePowerBi } from '../src/webparts/platformHub/app/providers/LiveDataProvider';
import { evaluateGovernance } from '../src/webparts/platformHub/app/logic/governance';
import { pbiReportRows, unusedReports } from '../src/webparts/platformHub/app/ui/pbiDerive';
import { SampleDataProvider } from '../src/webparts/platformHub/app/providers/SampleDataProvider';

const res = (body: unknown, status: number = 200): IHttpResponseLike => ({
  status, ok: status < 300, headers: { get: () => null }, json: async () => body, text: async () => JSON.stringify(body)
});

test('maps admin/groups rows with expanded content', () => {
  const into: IPbiContent = { workspaces: [], reports: [], datasets: [], dashboards: [], dataflows: [] };
  mapWorkspace({
    id: 'w1', name: 'Sales', type: 'Workspace', state: 'Active', isOnDedicatedCapacity: true, capacityId: 'CAP1',
    users: [{ emailAddress: 'a@x.com', groupUserAccessRight: 'Admin', principalType: 'User' }, { displayName: 'Team', groupUserAccessRight: 'Viewer', principalType: 'Group' }],
    reports: [{ id: 'r1', name: 'Pipeline', datasetId: 'd1', reportType: 'PaginatedReport', endorsementDetails: { endorsement: 'Certified' } }],
    datasets: [{ id: 'd1', name: 'Sales model', configuredBy: 'a@x.com', isRefreshable: true, isOnPremGatewayRequired: true, targetStorageMode: 'Abf' }],
    dashboards: [{ id: 'db1', displayName: 'KPIs' }],
    dataflows: [{ objectId: 'df1', name: 'Staging' }]
  }, into);
  mapWorkspace({ id: 'p1', name: 'PersonalWorkspace Ana', type: 'PersonalGroup' }, into);
  assert.equal(into.workspaces[0].admins, 1);
  assert.equal(into.workspaces[0].reports, 1);
  assert.equal(into.workspaces[1].type, 'Personal');
  assert.equal(into.reports[0].reportType, 'Paginated');
  assert.equal(into.reports[0].endorsement, 'Certified');
  assert.equal(into.datasets[0].storageMode, 'Import');
  assert.ok(into.datasets[0].gatewayRequired);
  assert.equal(into.dataflows[0].id, 'df1');
  const merged = mergePowerBi(into, [{ id: 'cap1', name: 'F64', sku: 'F64', state: 'Active', region: 'EU', admins: [] }], [], false,
    [{ id: 'd1', status: 'Failed', refreshFailures: 3 }], false);
  assert.equal(merged.workspaces[0].capacityName, 'F64', 'capacity ids match case-insensitively');
  assert.equal(merged.datasets[0].lastRefreshStatus, 'Failed');
});

test('pages admin/groups with $top/$skip until a short page', async () => {
  const urls: string[] = [];
  const api = new ApiClient({
    fetch: async (_r, url) => {
      urls.push(url);
      const skip = Number(/\$skip=(\d+)/.exec(url)![1]);
      return res({ value: skip === 0 ? new Array(5000).fill(0).map((_, i) => ({ id: `w${i}`, name: 'w' })) : [{ id: 'last', name: 'l' }] });
    }
  });
  const c = await new PowerBiService(api).getContent(100000);
  assert.equal(c.workspaces.length, 5001);
  assert.equal(urls.length, 2);
  assert.ok(/\$expand=users,reports,dashboards,datasets,dataflows&\$top=5000&\$skip=5000/.test(urls[1]));
});

test('activity: one UTC day per call, follows continuationUri, aggregates', async () => {
  const urls: string[] = [];
  const api = new ApiClient({
    fetch: async (_r, url) => {
      urls.push(url);
      if (url === 'next') return res({ activityEventEntities: [{ CreationTime: '2026-11-01T10:00:00Z', UserId: 'B@x.com', Activity: 'ExportReport', ItemName: 'Pipeline' }], lastResultSet: true });
      return res({
        activityEventEntities: [
          { CreationTime: '2026-11-01T09:00:00Z', UserId: 'a@x.com', Activity: 'ViewReport', ReportId: 'r1', ReportName: 'Pipeline', ConsumptionMethod: 'Power BI Web' },
          { CreationTime: '2026-11-01T09:05:00Z', UserId: 'b@x.com', Activity: 'ViewReport', ReportId: 'r1', ReportName: 'Pipeline' }
        ],
        continuationUri: urls.length === 1 ? 'next' : undefined,
        lastResultSet: urls.length !== 1
      });
    }
  });
  const a = await new PowerBiService(api).getActivity(2, 1000);
  assert.ok(/activityevents\?startDateTime='\d{4}-\d{2}-\d{2}T00:00:00.000Z'&endDateTime='/.test(urls[0]));
  assert.equal(urls[1], 'next');
  assert.equal(a.totalEvents, 5);
  assert.equal(a.activeUsers, 2, 'user ids are case-insensitive');
  assert.equal(a.views, 4);
  assert.equal(a.exports, 1);
  assert.equal(a.reports[0].viewers, 2);
  assert.ok(!a.truncated);
  const capped = await new PowerBiService(api).getActivity(1, 1);
  assert.equal(capped.totalEvents, 1);
  assert.ok(capped.truncated);
});

test('activity: a permission failure stops immediately', async () => {
  let calls = 0;
  const api = new ApiClient({ fetch: async () => { calls++; return res({ error: { message: 'Forbidden' } }, 403); } });
  await assert.rejects(new PowerBiService(api).getActivity(30, 100), (e: { kind: string }) => e.kind === 'forbidden');
  assert.equal(calls, 1);
});

test('activity windows cover N UTC days ending now', () => {
  const w = activityWindows(3, Date.parse('2026-11-02T10:00:00Z'));
  assert.deepEqual(w.map((x) => x.day), ['2026-10-31', '2026-11-01', '2026-11-02']);
  assert.equal(w[0].end, '2026-10-31T23:59:59.999Z');
  assert.equal(w[2].end, '2026-11-02T09:55:00.000Z');
});

test('aggregator flags exports, shares and publish to web', () => {
  const agg = new PbiActivityAggregator();
  ['AnalyzeInExcel', 'ShareReport', 'PublishToWebReport', 'ViewDashboard'].forEach((Activity, i) =>
    agg.add({ CreationTime: `2026-11-01T0${i}:00:00Z`, UserId: 'u', Activity, DashboardId: i === 3 ? 'd' : undefined }));
  const r = agg.result(1, '', '', []);
  assert.equal(r.exports, 1);
  assert.equal(r.shares, 1);
  assert.equal(r.publishToWeb, 1);
  assert.equal(r.sensitive.length, 3);
  assert.equal(r.sensitive[0].activity, 'PublishToWebReport', 'newest first');
});

test('Power BI governance rules and unused reports', async () => {
  const p = new SampleDataProvider(Date.parse('2026-11-02T09:00:00Z'));
  const data = await p.getPowerBi();
  const act = await p.getPowerBiActivity(7);
  const f = evaluateGovernance({ environments: [], apps: [], flows: [], agents: [], users: {}, pbi: data });
  const orphan = data.workspaces.filter((w) => w.type !== 'Personal' && (w.admins === 0 || w.state === 'Orphaned')).length;
  assert.equal(f.filter((x) => x.ruleId === 'pbi-orphan').length, orphan);
  assert.ok(orphan > 0);
  assert.equal(f.filter((x) => x.ruleId === 'pbi-refresh-failed').length, data.datasets.filter((d) => d.lastRefreshStatus === 'Failed').length);
  assert.ok(f.every((x) => !/PersonalWorkspace/.test(x.resourceName)), 'personal workspaces are not flagged');
  const rows = pbiReportRows(data, act);
  const unused = unusedReports(rows, data);
  assert.ok(unused.length > 0 && unused.every((r) => r.views === 0));
  assert.equal(rows.reduce((s, r) => s + (r.views || 0), 0), act.reports.reduce((s, r) => s + r.views, 0));
  assert.equal(pbiReportRows(data, undefined)[0].views, undefined, 'usage is unknown, not zero, without activity');
});
