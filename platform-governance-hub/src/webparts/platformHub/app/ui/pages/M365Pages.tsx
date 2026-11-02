import * as React from 'react';
import { Checkbox, Pivot, PivotItem, Text } from '@fluentui/react';
import { IDetail, useApp, useData } from '../AppContext';
import { IDataProvider } from '../../providers/IDataProvider';
import { ISite, ISku } from '../../models';
import { Card, Empty, Gate, KpiRow, Notice, PageGrid, Pill } from '../components/Layout';
import { BarList, ChartCard, ColumnChart, Donut, LineChart, Meter } from '../components/Charts';
import { DataGrid } from '../components/DataGrid';
import { PageFrame } from '../components/PageFrame';
import {
  accountColumns, credentialColumns, flattenRoles, groupColumns, inactiveUserColumns, issueColumns, messageColumns, mfaColumns, roleColumns, siteColumns,
  skuColumns, teamActivityColumns
} from '../columns';
import { adoptionSummary, identitySummary, inactiveLicensed, siteSummary, skuRows, teamsSummary } from '../derive';
import { loadPrices } from '../exportHelpers';
import { countBy, monthlyBuckets, topN } from '../../logic/aggregate';
import { daysSince, fmtBytes, fmtDate, fmtNumber } from '../../logic/format';

const concealed = (p: IDataProvider): Promise<boolean> => p.getReportsConcealed().then((c) => c, () => false);

const ConcealedNotice: React.FC = () => {
  const c = useData(concealed);
  return c.data ? (
    <Notice kind="warning">
      Usage reports show hidden (hashed) user, group and site names because “Display concealed user, group, and site names in all reports” is on.
      Turn it off in Microsoft 365 admin center &gt; Settings &gt; Org settings &gt; Reports to see real names.
    </Notice>
  ) : null;
};

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const monthLabel = (iso: string): string => MONTHS[new Date(Date.parse(iso)).getUTCMonth()];
const monthTitle = (iso: string): string => `${monthLabel(iso)} ${new Date(Date.parse(iso)).getUTCFullYear()}`;

/* ---------- SharePoint sites ---------- */

export const SitesPage: React.FC = () => {
  const { openDetail } = useApp();
  const sites = useData((p) => p.getSites());
  const [inactiveOnly, setInactiveOnly] = React.useState(false);
  const [externalOnly, setExternalOnly] = React.useState(false);
  const data = sites.data;
  const list = data ? data.sites.filter((s) => !s.isDeleted) : [];
  const view = list.filter((s) => {
    const d = daysSince(s.lastActivity);
    if (inactiveOnly && !(d === undefined || d > 90)) return false;
    if (externalOnly && !s.externalSharing) return false;
    return true;
  });
  const detail = (s: ISite): IDetail => ({
    title: s.title,
    subtitle: s.url,
    fields: [
      { label: 'Template', value: s.template },
      { label: 'Owner', value: s.ownerName || s.ownerUpn },
      { label: 'Storage used', value: fmtBytes(s.storageUsedBytes) },
      { label: 'Files', value: `${fmtNumber(s.fileCount)} (${fmtNumber(s.activeFileCount)} active in 30 days)` },
      { label: 'Page views (30 days)', value: fmtNumber(s.pageViews) },
      { label: 'External sharing', value: s.externalSharing ? 'On' : 'Off' },
      { label: 'Anyone links', value: fmtNumber(s.anonymousLinks) },
      { label: 'Last activity', value: s.lastActivity ? fmtDate(s.lastActivity) : 'None recorded' }
    ],
    links: s.url ? [{ label: 'Open site', href: s.url }] : []
  });
  return (
    <PageFrame id="sites">
      <Gate loading={sites.loading} error={sites.error} onRetry={sites.reload} label="SharePoint usage report">
        {data && (
          <PageGrid>
            <ConcealedNotice />
            <KpiRow items={siteSummary(data)} />
            {data.sharing && (
              <Notice>{`Tenant external sharing level: ${data.sharing.sharingCapability.replace(/([A-Z])/g, ' $1').toLowerCase()}.`}</Notice>
            )}
            <ChartCard title="Total SharePoint storage" subtitle="Last 90 days" span={8}>
              <LineChart title="Total SharePoint storage" dates={data.storageTrend.map((p) => p.date)} series={[{ name: 'Storage used', values: data.storageTrend.map((p) => p.bytes) }]} format={fmtBytes} axisFormat={fmtBytes} area />
            </ChartCard>
            <ChartCard title="Sites by template" span={4}>
              <Donut title="Sites by template" segments={topN(countBy(list, (s) => s.template), 6)} centerLabel="sites" />
            </ChartCard>
            <Card title="Largest sites" span={6}>
              <BarList title="Largest sites" rows={list.slice().sort((a, b) => b.storageUsedBytes - a.storageUsedBytes).slice(0, 10).map((s) => ({ label: s.title, value: s.storageUsedBytes, onClick: () => openDetail(detail(s)) }))} format={fmtBytes} colorIndex={2} />
            </Card>
            <Card title="Most viewed sites" subtitle="Page views in 30 days" span={6}>
              <BarList title="Most viewed sites" rows={list.slice().sort((a, b) => b.pageViews - a.pageViews).slice(0, 10).map((s) => ({ label: s.title, value: s.pageViews, onClick: () => openDetail(detail(s)) }))} colorIndex={0} />
            </Card>
            <Card span={12}>
              <DataGrid
                title="SharePoint sites"
                items={view}
                columns={siteColumns}
                onItemInvoked={(s) => openDetail(detail(s))}
                initialSort={{ key: 'storage', descending: true }}
                filters={
                  <>
                    <Checkbox label="Inactive 90 days" checked={inactiveOnly} onChange={(_, v) => setInactiveOnly(!!v)} />
                    <Checkbox label="External sharing on" checked={externalOnly} onChange={(_, v) => setExternalOnly(!!v)} />
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

/* ---------- Teams & groups ---------- */

export const TeamsPage: React.FC = () => {
  const teams = useData((p) => p.getTeams());
  const [owners, setOwners] = React.useState(false);
  const data = teams.data;
  return (
    <PageFrame id="teams">
      <Gate loading={teams.loading} error={teams.error} onRetry={teams.reload} label="Groups and teams">
        {data && (() => {
          const months = monthlyBuckets([data.groups.filter((g) => g.isTeam).map((g) => g.createdTime), data.groups.filter((g) => !g.isTeam).map((g) => g.createdTime)], 12);
          return (
            <PageGrid>
              <ConcealedNotice />
              <KpiRow items={teamsSummary(data)} />
              <ChartCard title="Groups created per month" span={8}>
                <ColumnChart title="Groups created per month" categories={months.buckets.map(monthLabel)} categoryTitles={months.buckets.map(monthTitle)} series={[{ name: 'Teams', values: months.values[0] }, { name: 'Groups without Teams', values: months.values[1] }]} stacked />
              </ChartCard>
              <ChartCard title="Groups by privacy" span={4}>
                <Donut title="Groups by privacy" segments={countBy(data.groups, (g) => g.visibility)} centerLabel="groups" />
              </ChartCard>
              <Card span={12} title="Microsoft 365 groups">
                <DataGrid
                  title="Microsoft 365 groups"
                  items={owners ? data.groups.filter((g) => g.ownerCount < 2) : data.groups}
                  columns={groupColumns}
                  initialSort={{ key: 'owners', descending: false }}
                  filters={<Checkbox label="Fewer than 2 owners" checked={owners} onChange={(_, v) => setOwners(!!v)} />}
                />
              </Card>
              <Card span={12} title="Team activity (30 days)">
                <DataGrid title="Team activity" items={data.activity} columns={teamActivityColumns} initialSort={{ key: 'active', descending: true }} emptyText="The Teams activity report returned no rows." />
              </Card>
            </PageGrid>
          );
        })()}
      </Gate>
    </PageFrame>
  );
};

/* ---------- Licensing & adoption ---------- */

export const AdoptionPage: React.FC = () => {
  const adoption = useData((p) => p.getAdoption());
  const skus = useData((p) => p.getSkus().then((s) => s, () => undefined as ISku[] | undefined));
  const data = adoption.data;
  const rows = skuRows(skus.data || [], loadPrices());
  return (
    <PageFrame id="adoption">
      <Gate loading={adoption.loading} error={adoption.error} onRetry={adoption.reload} label="Usage reports">
        {data && (
          <PageGrid>
            <ConcealedNotice />
            <KpiRow items={adoptionSummary(data, skus.data)} />
            <ChartCard title="Daily active users by workload" subtitle="Last 90 days" span={12}>
              <LineChart
                title="Daily active users by workload"
                dates={data.counts.map((c) => c.date)}
                series={[
                  { name: 'Microsoft 365', values: data.counts.map((c) => c.office365) },
                  { name: 'Exchange', values: data.counts.map((c) => c.exchange) },
                  { name: 'Teams', values: data.counts.map((c) => c.teams) },
                  { name: 'SharePoint', values: data.counts.map((c) => c.sharePoint) },
                  { name: 'OneDrive', values: data.counts.map((c) => c.oneDrive) }
                ]}
                height={240}
              />
            </ChartCard>
            <Card title="Licence use" subtitle="Assigned against purchased" span={6}>
              {rows.length ? (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                  {rows.filter((r) => r.enabled > 0 && r.enabled < 9000).slice(0, 12).map((r) => (
                    <div key={r.skuId} style={{ display: 'grid', gridTemplateColumns: 'minmax(120px, 1fr) minmax(120px, 1fr) 90px', gap: 10, alignItems: 'center' }}>
                      <Text variant="small">{r.friendlyName}</Text>
                      <Meter used={r.consumed} total={r.enabled} />
                      <Text variant="small" style={{ textAlign: 'right', fontVariantNumeric: 'tabular-nums' }}>{`${fmtNumber(r.consumed)} / ${fmtNumber(r.enabled)}`}</Text>
                    </div>
                  ))}
                </div>
              ) : (
                <Empty text="Licences could not be read." />
              )}
            </Card>
            <Card title="Microsoft 365 Copilot" subtitle="Users active in the last 30 days" span={6}>
              {data.copilotUnavailable ? (
                <Empty text="The Copilot usage report is not available in this tenant or for this account." />
              ) : (
                <Donut
                  title="Microsoft 365 Copilot activity"
                  segments={[
                    { label: 'Active in 30 days', value: data.copilotUsers.filter((u) => (daysSince(u.lastActivity) || 9999) <= 30).length },
                    { label: 'Licensed, not active', value: data.copilotUsers.filter((u) => (daysSince(u.lastActivity) || 9999) > 30).length }
                  ]}
                  centerLabel="Copilot users"
                />
              )}
            </Card>
            <Card span={12} title="All licences">
              <DataGrid title="All licences" items={rows} columns={skuColumns} height={320} initialSort={{ key: 'consumed', descending: true }} />
            </Card>
            <Card span={12} title="Licensed users with no activity in 30 days" subtitle="Candidates to reclaim licences">
              <DataGrid title="Inactive licensed users" items={inactiveLicensed(data)} columns={inactiveUserColumns} emptyText="Every licensed user was active in the last 30 days." />
            </Card>
          </PageGrid>
        )}
      </Gate>
    </PageFrame>
  );
};

/* ---------- Identity & roles ---------- */

export const IdentityPage: React.FC = () => {
  const identity = useData((p) => p.getIdentity());
  const data = identity.data;
  return (
    <PageFrame id="identity">
      <Gate loading={identity.loading} error={identity.error} onRetry={identity.reload} label="Directory data">
        {data && (() => {
          const guests = data.users.filter((u) => u.userType === 'Guest');
          const stale = data.users.filter((u) => u.accountEnabled && (daysSince(u.lastSignIn) === undefined || (daysSince(u.lastSignIn) || 0) > 90));
          const months = monthlyBuckets([guests.map((g) => g.createdTime)], 12);
          return (
            <PageGrid>
              <KpiRow items={identitySummary(data)} />
              {data.signInUnavailable && <Notice kind="warning">Sign-in activity is not available. It needs AuditLog.Read.All and an Entra ID P1 licence, so stale-account checks are skipped.</Notice>}
              <Card title="Admin role holders" span={6}>
                <BarList title="Admin role holders" rows={data.roles.map((r) => ({ label: r.displayName, value: r.members.length })).sort((a, b) => b.value - a.value).slice(0, 12)} colorIndex={6} emptyText="No role assignments could be read." />
              </Card>
              <ChartCard title="MFA registration (members)" span={6}>
                {data.mfaUnavailable ? (
                  <Empty text="MFA registration details need AuditLog.Read.All and Entra ID P1." />
                ) : (
                  <Donut
                    title="MFA registration"
                    segments={[
                      { label: 'Registered', value: data.mfa.filter((m) => m.isMfaRegistered && m.userType.toLowerCase() === 'member').length },
                      { label: 'Not registered', value: data.mfa.filter((m) => !m.isMfaRegistered && m.userType.toLowerCase() === 'member').length }
                    ]}
                    centerLabel="members"
                  />
                )}
              </ChartCard>
              <ChartCard title="Guests invited per month" span={12}>
                <ColumnChart title="Guests invited per month" categories={months.buckets.map(monthLabel)} categoryTitles={months.buckets.map(monthTitle)} series={[{ name: 'Guests', values: months.values[0] }]} height={180} />
              </ChartCard>
              <Card span={12}>
                <Pivot aria-label="Identity tables">
                  <PivotItem headerText={`Admin roles (${flattenRoles(data.roles).length})`}>
                    <DataGrid title="Admin roles" items={flattenRoles(data.roles)} columns={roleColumns} initialSort={{ key: 'role', descending: false }} />
                  </PivotItem>
                  <PivotItem headerText={`App credentials (${data.credentials.length})`}>
                    <DataGrid title="App credentials" items={data.credentials} columns={credentialColumns} initialSort={{ key: 'days', descending: false }} emptyText="No app registrations with secrets or certificates." />
                  </PivotItem>
                  <PivotItem headerText={`Guests (${guests.length})`}>
                    <DataGrid title="Guests" items={guests} columns={accountColumns} initialSort={{ key: 'signin', descending: false }} />
                  </PivotItem>
                  {!data.signInUnavailable && (
                    <PivotItem headerText={`Stale accounts (${stale.length})`}>
                      <DataGrid title="Stale accounts" items={stale} columns={accountColumns} initialSort={{ key: 'signin', descending: false }} emptyText="No enabled account is older than 90 days without sign-in." />
                    </PivotItem>
                  )}
                  {!data.mfaUnavailable && (
                    <PivotItem headerText={`Without MFA (${data.mfa.filter((m) => !m.isMfaRegistered).length})`}>
                      <DataGrid title="MFA registration" items={data.mfa.filter((m) => !m.isMfaRegistered)} columns={mfaColumns} initialSort={{ key: 'admin', descending: true }} />
                    </PivotItem>
                  )}
                </Pivot>
              </Card>
            </PageGrid>
          );
        })()}
      </Gate>
    </PageFrame>
  );
};

/* ---------- Service health (embedded in Export centre and Data sources) ---------- */

export const HealthTables: React.FC = () => {
  const health = useData((p) => p.getServiceHealth());
  return (
    <Gate loading={health.loading} error={health.error} label="Service health">
      {health.data && (
        <PageGrid>
          <Card span={12} title="Open service issues">
            {health.data.issues.length ? <DataGrid title="Service issues" items={health.data.issues} columns={issueColumns} height={260} /> : <Pill tone="good" text="No open issues" />}
          </Card>
          <Card span={12} title="Message center">
            <DataGrid title="Message center" items={health.data.messages} columns={messageColumns} height={320} />
          </Card>
        </PageGrid>
      )}
    </Gate>
  );
};
