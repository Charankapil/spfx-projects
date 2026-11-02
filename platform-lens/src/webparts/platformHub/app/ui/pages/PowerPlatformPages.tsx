import * as React from 'react';
import {
  Checkbox, DefaultButton, Dropdown, IDropdownOption, MessageBar, MessageBarType, PrimaryButton, ProgressIndicator, Text, TextField
} from '@fluentui/react';
import { IDataState, IDetail, useApp, useData } from '../AppContext';
import { IAgent, IApp, IDlpPolicy, IFinding, IFlow, IMaker, ISku, IProgress } from '../../models';
import { IDataProvider } from '../../providers/IDataProvider';
import { Card, Empty, Gate, KpiRow, Notice, PageGrid, Pill, severityTone } from '../components/Layout';
import { BarList, ChartCard, ColumnChart, Donut, Meter } from '../components/Charts';
import { DataGrid } from '../components/DataGrid';
import { PageFrame } from '../components/PageFrame';
import {
  agentColumns, appColumns, connectorColumns, dlpColumns, envColumns, findingColumns, flowColumns, IEnvRow, makerColumns, ownerLabel, skuColumns
} from '../columns';
import {
  agentSummary, appSummary, connectorUsage, envNameMap, envRows, envSummary, findingSummary, findings, flowSummary, IPPData, makerSummary, skuRows, unusedCost
} from '../derive';
import { loadPP } from '../sections';
import { loadPrices, savePrices } from '../exportHelpers';
import { countBy, monthlyBuckets, topN, weeklyBuckets } from '../../logic/aggregate';
import { fmtDate, fmtDateTime, fmtMb, fmtNumber } from '../../logic/format';
import { isCoveredByDlp, links, RULES, ruleAction } from '../../logic/governance';
import { describeError } from '../../http/ApiClient';

/* ---------- shared loaders ---------- */

function usePP(): IDataState<IPPData> {
  const { envFilter } = useApp();
  return useData((p) => loadPP(p, envFilter), [envFilter.join('|')]);
}

const safeDlp = (p: IDataProvider): Promise<IDlpPolicy[] | undefined> => p.getDlpPolicies().then((d) => d, () => undefined);

function useDlp(): IDataState<IDlpPolicy[] | undefined> {
  return useData(safeDlp);
}

const monthLabel = (iso: string): string => ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][new Date(Date.parse(iso)).getUTCMonth()];
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const weekLabel = (iso: string): string => {
  const d = new Date(Date.parse(iso));
  return `${d.getDate()} ${MONTHS[d.getMonth()]}`;
};
const monthTitle = (iso: string): string => {
  const d = new Date(Date.parse(iso));
  return `${monthLabel(iso)} ${d.getUTCFullYear()}`;
};

function options(values: string[], allLabel: string): IDropdownOption[] {
  return [{ key: '', text: allLabel }].concat(values.map((v) => ({ key: v, text: v })));
}

const FilterDropdown: React.FC<{ label: string; value: string; values: string[]; onChange: (v: string) => void; width?: number }> = (p) => (
  <Dropdown
    ariaLabel={p.label}
    selectedKey={p.value}
    options={options(p.values, p.label)}
    onChange={(_, o) => p.onChange(o ? String(o.key) : '')}
    styles={{ root: { width: p.width || 170 } }}
  />
);

/* ---------- detail builders ---------- */

const AppSharing: React.FC<{ app: IApp }> = ({ app }) => {
  // Model-driven apps are shared through Dataverse security roles, not app permissions.
  const sharing = useData((p) => (app.kind === 'Canvas' ? p.getAppSharing(app) : Promise.resolve(undefined)), [app.id]);
  if (app.kind !== 'Canvas') return null;
  return (
    <Card title="Sharing">
      <Gate loading={sharing.loading} error={sharing.error} label="Sharing">
        {sharing.data && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
            {sharing.data.everyone && <Pill tone="critical" text="Shared with everyone in the organisation" />}
            <Text variant="small">{`${sharing.data.users} users · ${sharing.data.groups} groups`}</Text>
            {sharing.data.principals.slice(0, 25).map((pr, i) => (
              <Text key={i} variant="small">{`${pr.name} — ${pr.type}, ${pr.role}`}</Text>
            ))}
          </div>
        )}
      </Gate>
    </Card>
  );
};

const FlowRuns: React.FC<{ flow: IFlow }> = ({ flow }) => {
  const runs = useData((p) => p.getFlowRuns(flow), [flow.id]);
  return (
    <Card title="Recent runs">
      <Gate loading={runs.loading} error={runs.error} label="Run history">
        {runs.data && runs.data.length ? (
          <>
            <Text variant="small">
              {`${runs.data.length} runs · ${runs.data.filter((r) => r.status === 'Failed').length} failed · ${Math.round(
                (runs.data.filter((r) => r.status === 'Succeeded').length / runs.data.length) * 100
              )}% succeeded`}
            </Text>
            {runs.data.slice(0, 12).map((r, i) => (
              <div key={i} style={{ display: 'flex', gap: 8, fontSize: 12 }}>
                <Pill tone={r.status === 'Succeeded' ? 'good' : r.status === 'Failed' ? 'critical' : 'neutral'} text={r.status} />
                <span>{fmtDateTime(r.startTime)}</span>
              </div>
            ))}
          </>
        ) : (
          <Empty text="No runs in the run history window." />
        )}
      </Gate>
    </Card>
  );
};

function appDetail(a: IApp, pp: IPPData): IDetail {
  return {
    title: a.displayName,
    subtitle: `${a.kind} app in ${a.environmentName}`,
    fields: [
      { label: 'Owner', value: ownerLabel(a.owner, pp.makers.users) },
      { label: 'Owner status', value: a.owner.id && pp.makers.users[a.owner.id] ? pp.makers.users[a.owner.id].status : 'Unknown' },
      { label: 'Created', value: fmtDate(a.createdTime) },
      { label: 'Last modified', value: fmtDate(a.lastModifiedTime) },
      { label: 'Last published', value: fmtDate(a.lastPublishTime) },
      { label: 'Connectors', value: a.connectors.map((c) => `${c.displayName} (${c.tier})`).join(', ') },
      { label: 'Premium licence needed', value: a.usesPremium ? 'Yes' : 'No' },
      { label: 'App ID', value: a.id }
    ],
    links: [{ label: 'Open in Power Apps', href: a.kind === 'Canvas' ? links.app(a.environmentId, a.id) : links.modelApp(a.environmentId) }, { label: 'Open environment in admin center', href: links.environment(a.environmentId) }],
    extra: <AppSharing app={a} />
  };
}

function flowDetail(f: IFlow, pp: IPPData): IDetail {
  return {
    title: f.displayName,
    subtitle: `${f.triggerType} flow in ${f.environmentName}`,
    fields: [
      { label: 'State', value: f.state + (f.suspensionReason ? ` (${f.suspensionReason})` : '') },
      { label: 'Owner', value: ownerLabel(f.owner, pp.makers.users) },
      { label: 'Created', value: fmtDate(f.createdTime) },
      { label: 'Last modified', value: fmtDate(f.lastModifiedTime) },
      { label: 'Actions', value: fmtNumber(f.actionCount) },
      { label: 'Connectors', value: f.connectors.map((c) => `${c.displayName} (${c.tier})`).join(', ') },
      { label: 'Flow ID', value: f.id }
    ],
    links: [{ label: 'Open in Power Automate', href: links.flow(f.environmentId, f.id) }],
    extra: <FlowRuns flow={f} />
  };
}

function agentDetail(b: IAgent, pp: IPPData): IDetail {
  return {
    title: b.displayName,
    subtitle: `Copilot Studio agent in ${b.environmentName}`,
    fields: [
      { label: 'Owner', value: ownerLabel(b.owner, pp.makers.users) },
      { label: 'Authentication', value: b.authMode === 'None' ? 'No authentication (public)' : b.authMode },
      { label: 'Published', value: b.publishedTime ? fmtDate(b.publishedTime) : 'Never' },
      { label: 'State', value: b.state },
      { label: 'Created', value: fmtDate(b.createdTime) },
      { label: 'Modified', value: fmtDate(b.modifiedTime) },
      { label: 'Agent ID', value: b.id }
    ],
    links: [{ label: 'Open in Copilot Studio', href: links.agent(b.environmentId, b.id) }]
  };
}

/* ---------- Overview ---------- */

export const OverviewPage: React.FC = () => {
  const { navigate, theme } = useApp();
  const pp = usePP();
  const dlp = useDlp();
  const health = useData((p) => p.getServiceHealth());
  const list = React.useMemo(() => (pp.data ? findings(pp.data, dlp.data, undefined, undefined) : []), [pp.data, dlp.data]);
  return (
    <PageFrame id="overview">
      <Gate loading={pp.loading} error={pp.error} onRetry={pp.reload} label="Power Platform inventory">
        {pp.data && (() => {
          const { inv, makers } = pp.data;
          const weeks = weeklyBuckets([inv.apps.map((a) => a.createdTime), inv.flows.map((f) => f.createdTime), inv.agents.map((b) => b.createdTime)], 26);
          const high = list.filter((f) => f.severity === 'High').length;
          const services = health.data ? health.data.overviews.filter((o) => o.status !== 'serviceOperational') : [];
          const envCounts: Record<string, number> = {};
          inv.apps.forEach((a) => (envCounts[a.environmentName] = (envCounts[a.environmentName] || 0) + 1));
          inv.flows.forEach((f) => (envCounts[f.environmentName] = (envCounts[f.environmentName] || 0) + 1));
          const byEnv = topN(Object.keys(envCounts).map((k) => ({ label: k, value: envCounts[k] })).sort((a, b) => b.value - a.value), 8);
          return (
            <PageGrid>
              <KpiRow
                items={[
                  { label: 'Environments', value: fmtNumber(inv.environments.length), sub: `${inv.environments.filter((e) => e.type === 'Production').length} production`, onClick: () => navigate('environments') },
                  { label: 'Apps', value: fmtNumber(inv.apps.length), sub: `+${fmtNumber(weeks.values[0].slice(-4).reduce((s, x) => s + x, 0))} in 4 weeks`, onClick: () => navigate('apps') },
                  { label: 'Cloud flows', value: fmtNumber(inv.flows.length), sub: `+${fmtNumber(weeks.values[1].slice(-4).reduce((s, x) => s + x, 0))} in 4 weeks`, onClick: () => navigate('flows') },
                  { label: 'Agents', value: fmtNumber(inv.agents.length), onClick: () => navigate('agents') },
                  { label: 'Makers', value: fmtNumber(makers.makers.length), onClick: () => navigate('makers') },
                  { label: 'High risks', value: fmtNumber(high), tone: high ? 'critical' : 'good', sub: high ? 'Open governance actions' : 'None found', onClick: () => navigate('governance') }
                ]}
              />
              {inv.errors.length > 0 && (
                <Notice kind="warning">
                  {`${inv.errors.length} environment request(s) failed, so some counts may be incomplete. `}
                  See the Environments page for details.
                </Notice>
              )}
              <ChartCard title="Resources created per week" subtitle="Last 26 weeks, by creation date" span={8}>
                <ColumnChart
                  title="Resources created per week"
                  categories={weeks.buckets.map(weekLabel)}
                  categoryTitles={weeks.buckets.map((b) => `Week of ${fmtDate(b)}`)}
                  series={[{ name: 'Apps', values: weeks.values[0] }, { name: 'Flows', values: weeks.values[1] }, { name: 'Agents', values: weeks.values[2] }]}
                  stacked
                />
              </ChartCard>
              <ChartCard title="Environments by type" span={4}>
                <Donut title="Environments by type" segments={topN(countBy(inv.environments, (e) => e.type), 6)} centerLabel="environments" />
              </ChartCard>
              <Card title="Top risks" subtitle="From the governance rules" span={7} actions={<DefaultButton text="All actions" onClick={() => navigate('governance')} />}>
                {list.length ? (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                    {list.slice(0, 8).map((f) => (
                      <div key={f.id} style={{ display: 'grid', gridTemplateColumns: '70px 1fr', gap: 8, alignItems: 'baseline' }}>
                        <Pill tone={severityTone(f.severity)} text={f.severity} />
                        <Text variant="small">
                          <strong>{f.title}</strong>
                          {` — ${f.resourceName}${f.environmentName && f.environmentName !== f.resourceName && f.environmentName !== '—' ? ` (${f.environmentName})` : ''}`}
                        </Text>
                      </div>
                    ))}
                  </div>
                ) : (
                  <Empty text="No risks found by the governance rules." />
                )}
              </Card>
              <Card title="Service health" span={5}>
                <Gate loading={health.loading} error={health.error} label="Service health">
                  {health.data && (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                      {services.length ? (
                        services.map((s) => <Pill key={s.service} tone={/restored|resolved/i.test(s.status) ? 'neutral' : 'warning'} text={`${s.service}: ${s.status.replace(/^service/, '').replace(/([A-Z])/g, ' $1').trim()}`} />)
                      ) : (
                        <Pill tone="good" text="All services operational" />
                      )}
                      <Text variant="small" style={{ color: theme.semanticColors.bodySubtext, marginTop: 6 }}>Latest message center posts</Text>
                      {health.data.messages.slice(0, 4).map((m) => (
                        <Text key={m.id} variant="small">{`${m.isMajorChange ? '★ ' : ''}${m.title} (${fmtDate(m.lastModified)})`}</Text>
                      ))}
                    </div>
                  )}
                </Gate>
              </Card>
              <Card title="Apps and flows by environment" span={6}>
                <BarList title="Apps and flows by environment" rows={byEnv} />
              </Card>
              <Card title="Most used connectors" subtitle="Apps + flows using each connector" span={6}>
                <BarList title="Most used connectors" rows={topN(connectorUsage(inv.apps, inv.flows, undefined).map((c) => ({ label: c.displayName, value: c.apps + c.flows })), 8)} colorIndex={1} />
              </Card>
            </PageGrid>
          );
        })()}
      </Gate>
    </PageFrame>
  );
};

/* ---------- Environments ---------- */

export const EnvironmentsPage: React.FC = () => {
  const { openDetail } = useApp();
  const pp = usePP();
  const dlp = useDlp();
  const [type, setType] = React.useState('');
  const rows = React.useMemo(() => (pp.data ? envRows(pp.data.inv, dlp.data) : []), [pp.data, dlp.data]);
  const view = type ? rows.filter((r) => r.type === type) : rows;
  const detail = (e: IEnvRow): IDetail => ({
    title: e.displayName,
    subtitle: `${e.type} environment · ${e.region}`,
    fields: [
      { label: 'Environment ID', value: e.id },
      { label: 'State', value: e.state },
      { label: 'Dataverse', value: e.orgUrl || 'No database' },
      { label: 'Dataverse version', value: e.dataverseVersion },
      { label: 'Managed Environment', value: e.isManaged ? 'Yes' : 'No' },
      { label: 'Security group', value: e.securityGroupId || 'None' },
      { label: 'DLP coverage', value: e.dlpCovered === undefined ? 'Unknown' : e.dlpCovered ? 'Covered' : 'Not covered by any policy' },
      { label: 'Capacity', value: `Database ${fmtMb(e.capacity.databaseMb)} · File ${fmtMb(e.capacity.fileMb)} · Log ${fmtMb(e.capacity.logMb)}` },
      { label: 'Resources', value: `${e.apps} apps · ${e.flows} flows · ${e.agents} agents · ${e.makers} makers` },
      { label: 'Created', value: `${fmtDate(e.createdTime)}${e.createdBy ? ` by ${e.createdBy}` : ''}` },
      { label: 'Expires', value: e.expirationTime ? fmtDate(e.expirationTime) : undefined }
    ],
    links: [{ label: 'Open in Power Platform admin center', href: links.environment(e.id) }]
  });
  return (
    <PageFrame id="environments">
      <Gate loading={pp.loading} error={pp.error} onRetry={pp.reload} label="Environments">
        {pp.data && (
          <PageGrid>
            <KpiRow items={envSummary(rows, !!dlp.data)} />
            {pp.data.inv.errors.length > 0 && (
              <div style={{ gridColumn: 'span 12' }}>
                <MessageBar messageBarType={MessageBarType.warning} isMultiline>
                  <strong>{`${pp.data.inv.errors.length} environment request(s) failed.`}</strong>
                  {pp.data.inv.errors.slice(0, 8).map((e, i) => (
                    <div key={i}>{`${e.environmentName} · ${e.area}: ${e.message}`}</div>
                  ))}
                  {pp.data.inv.errors.length > 8 && <div>{`…and ${pp.data.inv.errors.length - 8} more (included in the export).`}</div>}
                </MessageBar>
              </div>
            )}
            <ChartCard title="Environments by type" span={4}>
              <Donut title="Environments by type" segments={topN(countBy(rows, (e) => e.type), 6)} centerLabel="environments" />
            </ChartCard>
            <Card title="Environments by region" span={4}>
              <BarList title="Environments by region" rows={topN(countBy(rows, (e) => e.region), 8)} />
            </Card>
            <Card title="Largest Dataverse databases" span={4}>
              <BarList
                title="Largest Dataverse databases"
                rows={rows.filter((e) => e.capacity.databaseMb > 0).sort((a, b) => b.capacity.databaseMb - a.capacity.databaseMb).slice(0, 8).map((e) => ({ label: e.displayName, value: e.capacity.databaseMb }))}
                format={fmtMb}
                colorIndex={2}
                emptyText="No Dataverse capacity reported."
              />
            </Card>
            <Card span={12}>
              <DataGrid
                title="Environments"
                items={view}
                columns={envColumns}
                onItemInvoked={(e) => openDetail(detail(e))}
                initialSort={{ key: 'apps', descending: true }}
                filters={<FilterDropdown label="All types" value={type} values={countBy(rows, (e) => e.type).map((c) => c.label)} onChange={setType} />}
              />
            </Card>
          </PageGrid>
        )}
      </Gate>
    </PageFrame>
  );
};

/* ---------- Apps ---------- */

export const AppsPage: React.FC = () => {
  const { openDetail, provider, notify, version } = useApp();
  const pp = usePP();
  const [kind, setKind] = React.useState('');
  const [owner, setOwner] = React.useState('');
  const [premiumOnly, setPremiumOnly] = React.useState(false);
  const [scan, setScan] = React.useState<IProgress | undefined>(undefined);
  const data = pp.data;
  const apps = React.useMemo(() => {
    if (!data) return [];
    return data.inv.apps.filter((a) => {
      if (kind && a.kind !== kind) return false;
      if (premiumOnly && !a.usesPremium) return false;
      if (owner) {
        const u = a.owner.id ? data.makers.users[a.owner.id] : undefined;
        const left = !!u && (u.status === 'Deleted' || u.status === 'Disabled');
        if (owner === 'Orphaned' && !left) return false;
        if (owner === 'Active owner' && left) return false;
      }
      return true;
    });
    // version: sharing scans update apps in place
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data, kind, owner, premiumOnly, version]);
  const columns = React.useMemo(() => (data ? appColumns(data.makers.users) : []), [data]);
  const runScan = async (): Promise<void> => {
    if (!data) return;
    setScan({ label: 'Checking app sharing', done: 0, total: data.inv.apps.filter((a) => a.kind === 'Canvas').length });
    try {
      const n = await provider.scanAppSharing(data.inv.apps, setScan);
      notify(`Sharing scan finished: ${n} app(s) are shared with everyone.`, n ? 'info' : 'success');
    } catch (e) {
      notify(`Sharing scan failed: ${describeError(e)}`, 'error');
    } finally {
      setScan(undefined);
    }
  };
  return (
    <PageFrame id="apps" actions={<DefaultButton text="Scan sharing" iconProps={{ iconName: 'Share' }} disabled={!data || !!scan} onClick={() => { runScan().catch(() => undefined); }} title="Check which canvas apps are shared with everyone (one request per app)" />}>
      <Gate loading={pp.loading} error={pp.error} onRetry={pp.reload} label="Apps">
        {data && (() => {
          const months = monthlyBuckets([apps.filter((a) => a.kind === 'Canvas').map((a) => a.createdTime), apps.filter((a) => a.kind !== 'Canvas').map((a) => a.createdTime)], 12);
          return (
            <PageGrid>
              <KpiRow items={appSummary(apps, data.makers)} />
              {scan && (
                <div style={{ gridColumn: 'span 12' }}>
                  <ProgressIndicator label={scan.label} description={`${scan.done} of ${scan.total} apps`} percentComplete={scan.total ? scan.done / scan.total : 0} />
                </div>
              )}
              <ChartCard title="Apps created per month" subtitle="Last 12 months" span={8}>
                <ColumnChart title="Apps created per month" categories={months.buckets.map(monthLabel)} categoryTitles={months.buckets.map(monthTitle)} series={[{ name: 'Canvas', values: months.values[0] }, { name: 'Model-driven', values: months.values[1] }]} stacked />
              </ChartCard>
              <Card title="Top connectors in apps" span={4}>
                <BarList title="Top connectors in apps" rows={topN(connectorUsage(apps, [], undefined).map((c) => ({ label: c.displayName, value: c.apps })), 8)} colorIndex={1} />
              </Card>
              <Card span={12}>
                <DataGrid
                  title="Apps"
                  items={apps}
                  columns={columns}
                  onItemInvoked={(a) => openDetail(appDetail(a, data))}
                  initialSort={{ key: 'modified', descending: true }}
                  filters={
                    <>
                      <FilterDropdown label="All types" value={kind} values={['Canvas', 'Model-driven']} onChange={setKind} width={140} />
                      <FilterDropdown label="Any owner" value={owner} values={['Orphaned', 'Active owner']} onChange={setOwner} width={140} />
                      <Checkbox label="Premium only" checked={premiumOnly} onChange={(_, v) => setPremiumOnly(!!v)} />
                    </>
                  }
                />
              </Card>
            </PageGrid>
          );
        })()}
      </Gate>
    </PageFrame>
  );
};

/* ---------- Flows ---------- */

export const FlowsPage: React.FC = () => {
  const { openDetail } = useApp();
  const pp = usePP();
  const [state, setState] = React.useState('');
  const [trigger, setTrigger] = React.useState('');
  const data = pp.data;
  const all = data ? data.inv.flows : [];
  const flows = React.useMemo(() => all.filter((f) => (!state || f.state === state) && (!trigger || f.triggerType === trigger)), [all, state, trigger]);
  const columns = React.useMemo(() => (data ? flowColumns(data.makers.users) : []), [data]);
  return (
    <PageFrame id="flows">
      <Gate loading={pp.loading} error={pp.error} onRetry={pp.reload} label="Cloud flows">
        {data && (() => {
          const months = monthlyBuckets([all.map((f) => f.createdTime)], 12);
          return (
            <PageGrid>
              <KpiRow items={flowSummary(all, data.makers)} />
              <ChartCard title="Flows by state" span={4}>
                <Donut title="Flows by state" segments={countBy(all, (f) => f.state)} centerLabel="flows" />
              </ChartCard>
              <ChartCard title="Flows by trigger" span={4}>
                <Donut title="Flows by trigger" segments={topN(countBy(all, (f) => f.triggerType), 6)} centerLabel="flows" />
              </ChartCard>
              <ChartCard title="Flows created per month" span={4}>
                <ColumnChart title="Flows created per month" categories={months.buckets.map(monthLabel)} categoryTitles={months.buckets.map(monthTitle)} series={[{ name: 'Flows', values: months.values[0] }]} height={180} />
              </ChartCard>
              <Card span={12}>
                <DataGrid
                  title="Cloud flows"
                  items={flows}
                  columns={columns}
                  onItemInvoked={(f) => openDetail(flowDetail(f, data))}
                  initialSort={{ key: 'modified', descending: true }}
                  filters={
                    <>
                      <FilterDropdown label="All states" value={state} values={['On', 'Off', 'Suspended']} onChange={setState} width={130} />
                      <FilterDropdown label="All triggers" value={trigger} values={countBy(all, (f) => f.triggerType).map((c) => c.label)} onChange={setTrigger} width={150} />
                    </>
                  }
                />
              </Card>
            </PageGrid>
          );
        })()}
      </Gate>
    </PageFrame>
  );
};

/* ---------- Agents ---------- */

export const AgentsPage: React.FC = () => {
  const { openDetail, provider } = useApp();
  const pp = usePP();
  const data = pp.data;
  const columns = React.useMemo(() => (data ? agentColumns(data.makers.users) : []), [data]);
  const dvSource = provider.getSources().filter((s) => s.id === 'dv')[0];
  return (
    <PageFrame id="agents">
      <Gate loading={pp.loading} error={pp.error} onRetry={pp.reload} label="Agents">
        {data && (
          <PageGrid>
            <KpiRow items={agentSummary(data.inv.agents, data.makers)} />
            {dvSource && dvSource.state !== 'ok' && (
              <Notice kind="warning">{`Agents are read from Dataverse in each environment. Dataverse status: ${dvSource.state}. ${dvSource.message || ''}`}</Notice>
            )}
            <ChartCard title="Agents by authentication" span={4}>
              <Donut title="Agents by authentication" segments={countBy(data.inv.agents, (b) => (b.authMode === 'None' ? 'No authentication' : b.authMode))} centerLabel="agents" />
            </ChartCard>
            <Card title="Agents by environment" span={8}>
              <BarList title="Agents by environment" rows={topN(countBy(data.inv.agents, (b) => b.environmentName), 8)} colorIndex={6} emptyText="No agents found." />
            </Card>
            <Card span={12}>
              <DataGrid title="Agents" items={data.inv.agents} columns={columns} onItemInvoked={(b) => openDetail(agentDetail(b, data))} initialSort={{ key: 'modified', descending: true }} emptyText="No Copilot Studio agents found." />
            </Card>
          </PageGrid>
        )}
      </Gate>
    </PageFrame>
  );
};

/* ---------- Makers ---------- */

export const MakersPage: React.FC = () => {
  const { openDetail } = useApp();
  const pp = usePP();
  const [status, setStatus] = React.useState('');
  const [dept, setDept] = React.useState('');
  const data = pp.data;
  const makers = data ? data.makers.makers : [];
  const view = makers.filter((m) => (!status || m.status === status) && (!dept || (m.department || 'Unknown') === dept));
  const detail = (m: IMaker): IDetail => ({
    title: m.displayName,
    subtitle: m.upn,
    fields: [
      { label: 'Status', value: m.status },
      { label: 'Department', value: m.department },
      { label: 'Country', value: m.country },
      { label: 'Last sign-in', value: m.lastSignIn ? fmtDate(m.lastSignIn) : 'Unknown' },
      { label: 'Owns', value: `${m.apps} apps · ${m.flows} flows · ${m.agents} agents in ${m.environments} environments` },
      { label: 'First created', value: fmtDate(m.firstCreated) },
      { label: 'Last created', value: fmtDate(m.lastCreated) },
      { label: 'Object ID', value: m.id }
    ],
    extra: data ? (
      <Card title="Resources">
        {data.inv.apps.filter((a) => a.owner.id === m.id).slice(0, 15).map((a) => <Text key={a.id} variant="small">{`App · ${a.displayName} (${a.environmentName})`}</Text>)}
        {data.inv.flows.filter((f) => f.owner.id === m.id).slice(0, 15).map((f) => <Text key={f.id} variant="small">{`Flow · ${f.displayName} (${f.environmentName}, ${f.state})`}</Text>)}
        {data.inv.agents.filter((b) => b.owner.id === m.id).slice(0, 15).map((b) => <Text key={b.id} variant="small">{`Agent · ${b.displayName} (${b.environmentName})`}</Text>)}
      </Card>
    ) : undefined
  });
  return (
    <PageFrame id="makers">
      <Gate loading={pp.loading} error={pp.error} onRetry={pp.reload} label="Makers">
        {data && (() => {
          const months = monthlyBuckets([makers.map((m) => m.firstCreated)], 12);
          return (
            <PageGrid>
              <KpiRow items={makerSummary(data.makers)} />
              <ChartCard title="New makers per month" subtitle="By month of their first app, flow or agent" span={6}>
                <ColumnChart title="New makers per month" categories={months.buckets.map(monthLabel)} categoryTitles={months.buckets.map(monthTitle)} series={[{ name: 'New makers', values: months.values[0] }]} height={200} />
              </ChartCard>
              <Card title="Makers by department" span={6}>
                <BarList title="Makers by department" rows={topN(countBy(makers, (m) => m.department || 'Unknown'), 8)} colorIndex={2} />
              </Card>
              <Card title="Top makers" subtitle="By apps + flows + agents owned" span={12}>
                <BarList
                  title="Top makers"
                  rows={makers.slice().sort((a, b) => b.apps + b.flows + b.agents - (a.apps + a.flows + a.agents)).slice(0, 10).map((m) => ({ label: m.displayName, value: m.apps + m.flows + m.agents, hint: `${m.apps} apps, ${m.flows} flows, ${m.agents} agents`, onClick: () => openDetail(detail(m)) }))}
                  colorIndex={0}
                />
              </Card>
              <Card span={12}>
                <DataGrid
                  title="Makers"
                  items={view}
                  columns={makerColumns}
                  onItemInvoked={(m) => openDetail(detail(m))}
                  initialSort={{ key: 'flows', descending: true }}
                  filters={
                    <>
                      <FilterDropdown label="Any status" value={status} values={['Active', 'Disabled', 'Deleted', 'Unknown']} onChange={setStatus} width={130} />
                      <FilterDropdown label="All departments" value={dept} values={countBy(makers, (m) => m.department || 'Unknown').map((c) => c.label)} onChange={setDept} />
                    </>
                  }
                />
              </Card>
            </PageGrid>
          );
        })()}
      </Gate>
    </PageFrame>
  );
};

/* ---------- Connectors & DLP ---------- */

export const ConnectorsPage: React.FC = () => {
  const pp = usePP();
  const dlp = useData((p) => p.getDlpPolicies());
  const [tier, setTier] = React.useState('');
  const data = pp.data;
  const usage = React.useMemo(() => (data ? connectorUsage(data.inv.apps, data.inv.flows, dlp.data) : []), [data, dlp.data]);
  const uncovered = data && dlp.data ? data.inv.environments.filter((e) => !isCoveredByDlp(e, dlp.data || [])) : [];
  return (
    <PageFrame id="connectors">
      <Gate loading={pp.loading} error={pp.error} onRetry={pp.reload} label="Connectors">
        {data && (
          <PageGrid>
            <KpiRow
              items={[
                { label: 'Connectors in use', value: fmtNumber(usage.length) },
                { label: 'Premium', value: fmtNumber(usage.filter((u) => u.tier === 'Premium').length) },
                { label: 'Custom', value: fmtNumber(usage.filter((u) => u.tier === 'Custom').length), tone: usage.some((u) => u.tier === 'Custom') ? 'warning' : undefined },
                { label: 'DLP policies', value: dlp.data ? fmtNumber(dlp.data.length) : '—' },
                { label: 'Environments without DLP', value: dlp.data ? fmtNumber(uncovered.length) : '—', tone: dlp.data ? (uncovered.length ? 'critical' : 'good') : undefined }
              ]}
            />
            <Card title="Most used connectors" subtitle="Apps + flows" span={8}>
              <BarList title="Most used connectors" rows={topN(usage.map((c) => ({ label: c.displayName, value: c.apps + c.flows, hint: c.tier })), 12)} colorIndex={1} />
            </Card>
            <ChartCard title="Connector usage by tier" span={4}>
              <Donut title="Connector usage by tier" segments={['Standard', 'Premium', 'Custom'].map((t) => ({ label: t, value: usage.filter((u) => u.tier === t).reduce((s, u) => s + u.apps + u.flows, 0) }))} centerLabel="uses" />
            </ChartCard>
            <Card title="DLP policies" span={12}>
              <Gate loading={dlp.loading} error={dlp.error} onRetry={dlp.reload} label="DLP policies">
                {dlp.data && (
                  <>
                    {uncovered.length > 0 && (
                      <MessageBar messageBarType={MessageBarType.severeWarning} isMultiline>
                        {`Not covered by any DLP policy: ${uncovered.map((e) => e.displayName).join(', ')}`}
                      </MessageBar>
                    )}
                    <DataGrid title="DLP policies" items={dlp.data} columns={dlpColumns(envNameMap(data.inv.environments))} height={300} emptyText="No DLP policies exist. Every connector can be combined with every other." />
                  </>
                )}
              </Gate>
            </Card>
            <Card span={12}>
              <DataGrid
                title="Connector usage"
                items={tier ? usage.filter((u) => u.tier === tier) : usage}
                columns={connectorColumns}
                initialSort={{ key: 'flows', descending: true }}
                filters={<FilterDropdown label="All tiers" value={tier} values={['Standard', 'Premium', 'Custom']} onChange={setTier} width={130} />}
              />
            </Card>
          </PageGrid>
        )}
      </Gate>
    </PageFrame>
  );
};

/* ---------- Licensing & capacity ---------- */

const PriceEditor: React.FC<{ skus: ISku[]; prices: Record<string, number>; onChange: (p: Record<string, number>) => void }> = ({ skus, prices, onChange }) => {
  const [draft, setDraft] = React.useState<Record<string, string>>(() => {
    const d: Record<string, string> = {};
    skus.forEach((s) => (d[s.skuPartNumber] = prices[s.skuPartNumber] !== undefined ? String(prices[s.skuPartNumber]) : ''));
    return d;
  });
  const save = (): void => {
    const out: Record<string, number> = { ...prices };
    Object.keys(draft).forEach((k) => {
      const n = Number(draft[k].replace(',', '.'));
      if (draft[k].trim() === '') delete out[k];
      else if (isFinite(n) && n >= 0) out[k] = n;
    });
    onChange(out);
  };
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
      {skus.map((s) => (
        <div key={s.skuId} style={{ display: 'grid', gridTemplateColumns: 'minmax(140px, 1fr) minmax(120px, 1fr) 110px', gap: 10, alignItems: 'center' }}>
          <Text variant="small">{s.friendlyName}</Text>
          <Meter used={s.consumed} total={s.enabled} label={`${s.consumed} of ${s.enabled} assigned`} />
          <TextField
            ariaLabel={`Monthly price for ${s.friendlyName}`}
            placeholder="Price / month"
            value={draft[s.skuPartNumber] || ''}
            onChange={(_, v) => setDraft({ ...draft, [s.skuPartNumber]: v || '' })}
            inputMode="decimal"
          />
        </div>
      ))}
      <div>
        <PrimaryButton text="Save prices" onClick={save} />
      </div>
      <Text variant="small">Prices are stored only in this browser and used to estimate the monthly cost of unassigned licences.</Text>
    </div>
  );
};

export const CapacityPage: React.FC = () => {
  const pp = usePP();
  const skus = useData((p) => p.getSkus());
  const [prices, setPrices] = React.useState<Record<string, number>>(loadPrices);
  const data = pp.data;
  const ppSkus = skus.data ? skus.data.filter((s) => s.isPowerPlatform) : [];
  const rows = skuRows(ppSkus, prices);
  const envs = data ? envRows(data.inv, undefined).filter((e) => e.hasDataverse) : [];
  const cost = unusedCost(rows);
  return (
    <PageFrame id="capacity">
      <Gate loading={pp.loading} error={pp.error} onRetry={pp.reload} label="Capacity">
        {data && (
          <PageGrid>
            <KpiRow
              items={[
                { label: 'Power Platform SKUs', value: skus.data ? fmtNumber(rows.length) : '—' },
                { label: 'Unassigned licences', value: skus.data ? fmtNumber(rows.filter((r) => r.appliesTo === 'User').reduce((s, r) => s + r.available, 0)) : '—', sub: 'User licences only' },
                { label: 'Unassigned cost / month', value: cost ? fmtNumber(cost) : '—', sub: cost ? 'From your unit prices' : 'Enter prices below' },
                { label: 'Database used', value: fmtMb(envs.reduce((s, e) => s + e.capacity.databaseMb, 0)) },
                { label: 'File used', value: fmtMb(envs.reduce((s, e) => s + e.capacity.fileMb, 0)) },
                { label: 'Log used', value: fmtMb(envs.reduce((s, e) => s + e.capacity.logMb, 0)) }
              ]}
            />
            <Card title="Power Platform licences" subtitle="Assigned against purchased, with optional unit prices" span={7}>
              <Gate loading={skus.loading} error={skus.error} onRetry={skus.reload} label="Licences">
                {ppSkus.length ? (
                  <PriceEditor skus={ppSkus} prices={prices} onChange={(p) => { savePrices(p); setPrices(p); }} />
                ) : (
                  <Empty text="No standalone Power Platform licences found." />
                )}
              </Gate>
            </Card>
            <ChartCard title="Dataverse capacity by type" span={5}>
              <Donut
                title="Dataverse capacity by type"
                segments={[
                  { label: 'Database', value: envs.reduce((s, e) => s + e.capacity.databaseMb, 0) },
                  { label: 'File', value: envs.reduce((s, e) => s + e.capacity.fileMb, 0) },
                  { label: 'Log', value: envs.reduce((s, e) => s + e.capacity.logMb, 0) }
                ]}
                format={fmtMb}
                centerLabel="used"
              />
            </ChartCard>
            {skus.data && (
              <Card span={12}>
                <DataGrid title="Power Platform licences" items={rows} columns={skuColumns} height={300} initialSort={{ key: 'consumed', descending: true }} />
              </Card>
            )}
            <Card span={12} title="Dataverse capacity by environment">
              <DataGrid title="Dataverse capacity" items={envs} columns={envColumns} initialSort={{ key: 'db', descending: true }} />
            </Card>
          </PageGrid>
        )}
      </Gate>
    </PageFrame>
  );
};

/* ---------- Governance ---------- */

export const GovernancePage: React.FC = () => {
  const { openDetail, version, theme } = useApp();
  const pp = usePP();
  const extras = useData((p) =>
    Promise.all([safeDlp(p), p.getTeams().then((t) => t, () => undefined), p.getAppCredentials().then((t) => t, () => undefined), p.getPowerBi().then((t) => t, () => undefined)])
  );
  const [severity, setSeverity] = React.useState('');
  const [category, setCategory] = React.useState('');
  const [rule, setRule] = React.useState('');
  const list = React.useMemo(
    () => (pp.data && extras.data ? findings(pp.data, extras.data[0], extras.data[1], extras.data[2], extras.data[3]) : []),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [pp.data, extras.data, version]
  );
  const view = list.filter((f) => (!severity || f.severity === severity) && (!category || f.category === category) && (!rule || f.ruleId === rule));
  const detail = (f: IFinding): IDetail => ({
    title: f.title,
    subtitle: `${f.severity} · ${f.category}`,
    fields: [
      { label: 'Resource', value: `${f.resourceType}: ${f.resourceName}` },
      { label: 'Environment', value: f.environmentName },
      { label: 'Owner', value: f.owner },
      { label: 'Detail', value: f.detail },
      { label: 'Recommended action', value: ruleAction(f.ruleId) }
    ],
    links: f.link ? [{ label: 'Open the resource', href: f.link }] : []
  });
  const loading = pp.loading || extras.loading;
  return (
    <PageFrame id="governance">
      <Gate loading={loading} error={pp.error} onRetry={pp.reload} label="Governance data">
        {pp.data && extras.data && (
          <PageGrid>
            <KpiRow items={findingSummary(list)} />
            {extras.data[0] === undefined && <Notice kind="warning">DLP policies could not be read, so the “environment not covered by DLP” rule was skipped.</Notice>}
            <Card title="Rules" subtitle="Click a rule to filter the worklist" span={12}>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(260px, 1fr))', gap: 8 }}>
                {RULES.map((r) => {
                  const n = list.filter((f) => f.ruleId === r.id).length;
                  return (
                    <button
                      key={r.id}
                      type="button"
                      aria-pressed={rule === r.id}
                      onClick={() => {
                        setRule(rule === r.id ? '' : r.id);
                        setSeverity('');
                        setCategory('');
                      }}
                      style={{ display: 'grid', gridTemplateColumns: '64px 1fr auto', gap: 8, alignItems: 'center', textAlign: 'left', background: rule === r.id ? theme.palette.neutralLighter : 'transparent', border: 'none', borderRadius: 4, padding: '4px 6px', cursor: 'pointer', color: 'inherit', font: 'inherit' }}
                    >
                      <Pill tone={severityTone(r.severity)} text={r.severity} />
                      <Text variant="small">{r.title}</Text>
                      <Text variant="small" style={{ fontWeight: 600, fontVariantNumeric: 'tabular-nums' }}>{fmtNumber(n)}</Text>
                    </button>
                  );
                })}
              </div>
            </Card>
            <Card span={12}>
              <DataGrid
                title="Governance findings"
                items={view}
                columns={findingColumns}
                onItemInvoked={(f) => openDetail(detail(f))}
                emptyText="No findings. Every rule passed."
                filters={
                  <>
                    <FilterDropdown label="Any severity" value={severity} values={['High', 'Medium', 'Low']} onChange={setSeverity} width={130} />
                    <FilterDropdown label="All categories" value={category} values={RULES.map((r) => r.category).filter((c, i, a) => a.indexOf(c) === i)} onChange={setCategory} />
                    {rule && <DefaultButton text="Show all rules" iconProps={{ iconName: 'ClearFilter' }} onClick={() => setRule('')} />}
                  </>
                }
              />
            </Card>
          </PageGrid>
        )}
      </Gate>
    </PageFrame>
  );
};

