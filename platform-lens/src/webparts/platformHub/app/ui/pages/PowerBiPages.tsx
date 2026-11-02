import * as React from 'react';
import { Checkbox, Dropdown, Pivot, PivotItem, Text } from '@fluentui/react';
import { IDetail, IDataState, useApp, useData } from '../AppContext';
import { IPbiActivity, IPbiWorkspace, IPowerBiData, ISku } from '../../models';
import { IDataProvider } from '../../providers/IDataProvider';
import { Card, Empty, Gate, KpiRow, Notice, PageGrid } from '../components/Layout';
import { BarList, ChartCard, ColumnChart, Donut, LineChart, Meter } from '../components/Charts';
import { DataGrid } from '../components/DataGrid';
import { PageFrame } from '../components/PageFrame';
import {
  pbiAppColumns, pbiCapacityColumns, pbiDashboardColumns, pbiDataflowColumns, pbiDatasetColumns, pbiEventColumns, pbiReportColumns,
  pbiUsageColumns, pbiUserColumns, pbiWorkspaceColumns
} from '../columns';
import { pbiActivitySummary, pbiContentSummary } from '../derive';
import { pbiReportRows, unusedReports } from '../pbiDerive';
import { loadPbiDays, PBI_DAY_OPTIONS, savePbiDays } from '../prefs';
import { countBy, topN } from '../../logic/aggregate';
import { describeError } from '../../http/ApiClient';
import { fmtDate, fmtNumber } from '../../logic/format';
import { links } from '../../logic/governance';

/* ---------- shared ---------- */

function usePeriod(): [number, (d: number) => void] {
  const [days, setDays] = React.useState<number>(loadPbiDays);
  return [days, (d: number) => { savePbiDays(d); setDays(d); }];
}

const PeriodPicker: React.FC<{ days: number; onChange: (d: number) => void }> = ({ days, onChange }) => (
  <Dropdown
    ariaLabel="Activity period"
    selectedKey={days}
    options={PBI_DAY_OPTIONS.map((d) => ({ key: d, text: d === 1 ? 'Activity: last 24 hours' : `Activity: last ${d} days` }))}
    onChange={(_, o) => o && onChange(Number(o.key))}
    styles={{ root: { width: 200 } }}
  />
);

const safeActivity = (days: number): ((p: IDataProvider) => Promise<{ a?: IPbiActivity; error?: unknown }>) => (p: IDataProvider) =>
  p.getPowerBiActivity(days).then((a) => ({ a }), (error) => ({ error }));

function useActivity(days: number): IDataState<{ a?: IPbiActivity; error?: unknown }> {
  return useData(safeActivity(days), [days]);
}

const ActivityNotice: React.FC<{ state: IDataState<{ a?: IPbiActivity; error?: unknown }> }> = ({ state }) => {
  if (state.loading || !state.data) return null;
  if (state.data.error) return <Notice kind="warning">{`The Power BI activity log could not be read, so usage figures are missing. ${describeError(state.data.error)}`}</Notice>;
  const a = state.data.a;
  if (!a) return null;
  return (
    <>
      {a.truncated && <Notice kind="warning">{`The activity log has more events than the row limit allows, so totals cover only part of the ${a.days}-day period. Raise "Maximum rows" in the web part settings or pick a shorter period.`}</Notice>}
      {a.failedDays.length > 0 && <Notice kind="warning">{`Some days could not be read (usually throttling): ${a.failedDays.join(', ')}. Refresh later to fill them in.`}</Notice>}
    </>
  );
};

const shortDay = (iso: string): string => {
  const d = new Date(Date.parse(iso + 'T00:00:00Z'));
  return `${d.getUTCDate()} ${['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][d.getUTCMonth()]}`;
};

function workspaceDetail(w: IPbiWorkspace): IDetail {
  return {
    title: w.name,
    subtitle: `${w.type} workspace · ${w.state}`,
    fields: [
      { label: 'Capacity', value: w.isOnDedicatedCapacity ? w.capacityName || w.capacityId : 'Shared capacity' },
      { label: 'Content', value: `${w.reports} reports · ${w.datasets} semantic models · ${w.dashboards} dashboards · ${w.dataflows} dataflows` },
      { label: 'Admins', value: w.users.filter((u) => u.right === 'Admin').map((u) => u.email || u.name).join(', ') || 'None' },
      { label: 'Workspace ID', value: w.id }
    ],
    links: [{ label: 'Open in Power BI', href: links.pbiWorkspace(w.id) }],
    extra: (
      <Card title={`Access (${w.users.length})`}>
        {w.users.length ? w.users.map((u, i) => <Text key={i} variant="small">{`${u.name}${u.email && u.email !== u.name ? ` (${u.email})` : ''} — ${u.right}${u.principalType !== 'User' ? `, ${u.principalType}` : ''}`}</Text>) : <Empty text="No access entries." />}
      </Card>
    )
  };
}

const isPbiSku = (s: ISku): boolean => /POWER_BI|PBI_|SPE_E5|ENTERPRISEPREMIUM|FABRIC/i.test(s.skuPartNumber);

/* ---------- Overview ---------- */

export const PowerBiOverviewPage: React.FC = () => {
  const { navigate } = useApp();
  const [days, setDays] = usePeriod();
  const content = useData((p) => p.getPowerBi());
  const activity = useActivity(days);
  const skus = useData((p) => p.getSkus().then((s) => s, () => [] as ISku[]));
  const data = content.data;
  const act = activity.data && activity.data.a;
  return (
    <PageFrame id="pbi" actions={<PeriodPicker days={days} onChange={setDays} />}>
      <Gate loading={content.loading} error={content.error} onRetry={content.reload} label="Power BI admin data">
        {data && (
          <PageGrid>
            <KpiRow items={pbiContentSummary(data).map((k, i) => ({ ...k, onClick: i < 5 ? () => navigate('pbiContent') : undefined }))} />
            <ActivityNotice state={activity} />
            {activity.loading && <Notice>{`Reading the Power BI activity log for the last ${days} day(s)…`}</Notice>}
            {act && <KpiRow items={pbiActivitySummary(act).map((k) => ({ ...k, onClick: () => navigate('pbiActivity') }))} />}
            {act && (
              <>
                <ChartCard title="Daily active users" subtitle="Distinct users with any Power BI activity" span={6}>
                  <LineChart title="Daily active users" dates={act.daily.map((d) => d.date)} series={[{ name: 'Active users', values: act.daily.map((d) => d.activeUsers) }]} area height={200} />
                </ChartCard>
                <ChartCard title="Report and dashboard views per day" span={6}>
                  <ColumnChart title="Views per day" categories={act.daily.map((d) => shortDay(d.date))} series={[{ name: 'Views', values: act.daily.map((d) => d.views) }]} height={200} />
                </ChartCard>
                <Card title="Most viewed reports" subtitle={`Last ${act.days} day(s)`} span={7}>
                  <BarList title="Most viewed reports" rows={act.reports.slice(0, 10).map((r) => ({ label: r.name, value: r.views, hint: `${r.viewers} viewers · ${r.workspace}` }))} emptyText="No views in this period." />
                </Card>
                <ChartCard title="Views by app" subtitle="Where people open Power BI content" span={5}>
                  <Donut title="Views by app" segments={topN(act.byConsumption, 5)} centerLabel="views" />
                </ChartCard>
              </>
            )}
            <ChartCard title="Workspaces by type" span={4}>
              <Donut title="Workspaces by type" segments={countBy(data.workspaces, (w) => w.type)} centerLabel="workspaces" />
            </ChartCard>
            <Card title="Largest workspaces" subtitle="By reports + semantic models" span={4}>
              <BarList
                title="Largest workspaces"
                rows={data.workspaces.filter((w) => w.type !== 'Personal').sort((a, b) => b.reports + b.datasets - (a.reports + a.datasets)).slice(0, 8).map((w) => ({ label: w.name, value: w.reports + w.datasets }))}
                colorIndex={2}
              />
            </Card>
            <Card title="Power BI licences" span={4}>
              {skus.data && skus.data.filter(isPbiSku).length ? (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                  {skus.data.filter(isPbiSku).filter((s) => s.enabled < 9000).map((s) => (
                    <div key={s.skuId} style={{ display: 'grid', gridTemplateColumns: 'minmax(110px, 1fr) minmax(90px, 1fr)', gap: 10, alignItems: 'center' }}>
                      <Text variant="small">{`${s.friendlyName} (${fmtNumber(s.consumed)}/${fmtNumber(s.enabled)})`}</Text>
                      <Meter used={s.consumed} total={s.enabled} />
                    </div>
                  ))}
                  <Text variant="small">Microsoft 365 E5 includes Power BI Pro.</Text>
                </div>
              ) : (
                <Empty text="No Power BI licences found, or licences could not be read." />
              )}
            </Card>
            <Card title="Capacities" span={12}>
              {data.capacities.length ? <DataGrid title="Capacities" items={data.capacities} columns={pbiCapacityColumns} height={240} /> : <Empty text="No Premium or Fabric capacities. All workspaces use shared capacity." />}
            </Card>
          </PageGrid>
        )}
      </Gate>
    </PageFrame>
  );
};

/* ---------- Workspaces & content ---------- */

export const PowerBiContentPage: React.FC = () => {
  const { openDetail } = useApp();
  const [days, setDays] = usePeriod();
  const content = useData((p) => p.getPowerBi());
  const activity = useActivity(days);
  const [type, setType] = React.useState('');
  const [includePersonal, setIncludePersonal] = React.useState(false);
  const [noAdmin, setNoAdmin] = React.useState(false);
  const [failedOnly, setFailedOnly] = React.useState(false);
  const data: IPowerBiData | undefined = content.data;
  const act = activity.data && activity.data.a;
  const reports = React.useMemo(() => (data ? pbiReportRows(data, act) : []), [data, act]);
  const personal = React.useMemo(() => {
    const m: Record<string, boolean> = {};
    if (data) data.workspaces.forEach((w) => w.type === 'Personal' && (m[w.id] = true));
    return m;
  }, [data]);
  const workspaces = data
    ? data.workspaces.filter((w) => (includePersonal || w.type !== 'Personal') && (!type || w.type === type) && (!noAdmin || (w.type !== 'Personal' && (w.admins === 0 || w.state === 'Orphaned'))))
    : [];
  const scope = <T extends { workspaceId?: string }>(list: T[]): T[] => (includePersonal ? list : list.filter((x) => !x.workspaceId || !personal[x.workspaceId]));
  const personalBox = <Checkbox label="Include personal workspaces" checked={includePersonal} onChange={(_, v) => setIncludePersonal(!!v)} />;
  return (
    <PageFrame id="pbiContent" actions={<PeriodPicker days={days} onChange={setDays} />}>
      <Gate loading={content.loading} error={content.error} onRetry={content.reload} label="Power BI admin data">
        {data && (
          <PageGrid>
            <KpiRow items={pbiContentSummary(data)} />
            <ActivityNotice state={activity} />
            {data.refreshUnavailable && <Notice>Refresh status comes from Premium and Fabric capacities only, and could not be read. Models on shared capacity show no refresh status.</Notice>}
            <Card span={12}>
              <Pivot aria-label="Power BI content">
                <PivotItem headerText={`Workspaces (${workspaces.length})`}>
                  <DataGrid
                    title="Power BI workspaces"
                    items={workspaces}
                    columns={pbiWorkspaceColumns}
                    onItemInvoked={(w) => openDetail(workspaceDetail(w))}
                    initialSort={{ key: 'reports', descending: true }}
                    filters={
                      <>
                        <Dropdown ariaLabel="Workspace type" selectedKey={type} options={[{ key: '', text: 'All types' }, { key: 'Workspace', text: 'Workspace' }, { key: 'Classic', text: 'Classic' }, { key: 'Personal', text: 'Personal' }]} onChange={(_, o) => o && setType(String(o.key))} styles={{ root: { width: 140 } }} />
                        {personalBox}
                        <Checkbox label="Without admin" checked={noAdmin} onChange={(_, v) => setNoAdmin(!!v)} />
                      </>
                    }
                  />
                </PivotItem>
                <PivotItem headerText={`Reports (${scope(reports).length})`}>
                  <DataGrid title="Power BI reports" items={scope(reports)} columns={pbiReportColumns} initialSort={{ key: act ? 'views' : 'modified', descending: true }} filters={personalBox} />
                </PivotItem>
                <PivotItem headerText={`Semantic models (${scope(data.datasets).length})`}>
                  <DataGrid
                    title="Semantic models"
                    items={scope(data.datasets).filter((d) => !failedOnly || d.lastRefreshStatus === 'Failed')}
                    columns={pbiDatasetColumns}
                    initialSort={{ key: 'failures', descending: true }}
                    filters={<>{personalBox}<Checkbox label="Failed refresh only" checked={failedOnly} onChange={(_, v) => setFailedOnly(!!v)} /></>}
                  />
                </PivotItem>
                <PivotItem headerText={`Dashboards (${scope(data.dashboards).length})`}>
                  <DataGrid title="Dashboards" items={scope(data.dashboards)} columns={pbiDashboardColumns} filters={personalBox} />
                </PivotItem>
                <PivotItem headerText={`Dataflows (${data.dataflows.length})`}>
                  <DataGrid title="Dataflows" items={data.dataflows} columns={pbiDataflowColumns} initialSort={{ key: 'modified', descending: true }} />
                </PivotItem>
                <PivotItem headerText={`Apps (${data.apps.length})`}>
                  {data.appsUnavailable ? <Empty text="Power BI apps could not be read." /> : <DataGrid title="Power BI apps" items={data.apps} columns={pbiAppColumns} initialSort={{ key: 'updated', descending: true }} />}
                </PivotItem>
                <PivotItem headerText={`Capacities (${data.capacities.length})`}>
                  <DataGrid title="Capacities" items={data.capacities} columns={pbiCapacityColumns} emptyText="No Premium or Fabric capacities." />
                </PivotItem>
              </Pivot>
            </Card>
          </PageGrid>
        )}
      </Gate>
    </PageFrame>
  );
};

/* ---------- Activity ---------- */

export const PowerBiActivityPage: React.FC = () => {
  const [days, setDays] = usePeriod();
  const activity = useData((p) => p.getPowerBiActivity(days), [days]);
  const content = useData((p) => p.getPowerBi().then((d) => d, () => undefined));
  const a = activity.data;
  const unused = React.useMemo(() => (a && content.data ? unusedReports(pbiReportRows(content.data, a), content.data) : undefined), [a, content.data]);
  return (
    <PageFrame id="pbiActivity" actions={<PeriodPicker days={days} onChange={setDays} />}>
      <Gate loading={activity.loading} error={activity.error} onRetry={activity.reload} label="Power BI activity log">
        {a && (
          <PageGrid>
            <KpiRow items={pbiActivitySummary(a)} />
            {a.truncated && <Notice kind="warning">{`The activity log has more events than the row limit allows, so totals cover only part of the period. Raise "Maximum rows" in the web part settings or pick a shorter period.`}</Notice>}
            {a.failedDays.length > 0 && <Notice kind="warning">{`Some days could not be read (usually throttling): ${a.failedDays.join(', ')}.`}</Notice>}
            <ChartCard title="Daily active users" subtitle={`${fmtDate(a.from)} to ${fmtDate(a.to)} (UTC days)`} span={6}>
              <LineChart title="Daily active users" dates={a.daily.map((d) => d.date)} series={[{ name: 'Active users', values: a.daily.map((d) => d.activeUsers) }]} area height={200} />
            </ChartCard>
            <ChartCard title="Activities per day" subtitle="Views and all other activities" span={6}>
              <ColumnChart
                title="Activities per day"
                categories={a.daily.map((d) => shortDay(d.date))}
                series={[{ name: 'Views', values: a.daily.map((d) => d.views) }, { name: 'Other activities', values: a.daily.map((d) => d.events - d.views) }]}
                stacked
                height={200}
              />
            </ChartCard>
            <Card title="Activity types" span={7}>
              <BarList title="Activity types" rows={topN(a.byActivity, 12)} colorIndex={6} />
            </Card>
            <ChartCard title="Views by app" span={5}>
              <Donut title="Views by app" segments={topN(a.byConsumption, 5)} centerLabel="views" />
            </ChartCard>
            <Card span={12}>
              <Pivot aria-label="Power BI activity tables">
                <PivotItem headerText={`Most viewed (${a.reports.length})`}>
                  <DataGrid title="Most viewed reports and dashboards" items={a.reports} columns={pbiUsageColumns} initialSort={{ key: 'views', descending: true }} />
                </PivotItem>
                <PivotItem headerText={`Users (${a.users.length})`}>
                  <DataGrid title="Power BI users" items={a.users} columns={pbiUserColumns} initialSort={{ key: 'events', descending: true }} />
                </PivotItem>
                <PivotItem headerText={`Exports & sharing (${a.sensitive.length})`}>
                  <DataGrid title="Exports, sharing and publish to web" items={a.sensitive} columns={pbiEventColumns} initialSort={{ key: 'time', descending: true }} emptyText="No exports, shares or publish-to-web events in this period." />
                </PivotItem>
                <PivotItem headerText={unused ? `Unused reports (${unused.length})` : 'Unused reports'}>
                  {unused ? (
                    <DataGrid title="Reports with no views" items={unused} columns={pbiReportColumns} initialSort={{ key: 'modified', descending: false }} emptyText="Every report in a shared workspace was viewed in this period." />
                  ) : (
                    <Empty text="Workspace content could not be read, so unused reports cannot be listed." />
                  )}
                </PivotItem>
                <PivotItem headerText="Recent events">
                  <DataGrid title="Recent Power BI events" items={a.recent} columns={pbiEventColumns} initialSort={{ key: 'time', descending: true }} />
                </PivotItem>
              </Pivot>
            </Card>
          </PageGrid>
        )}
      </Gate>
    </PageFrame>
  );
};
