import * as React from 'react';
import { DefaultButton, DetailsList, DetailsListLayoutMode, Dropdown, IColumn, MessageBar, MessageBarType, PrimaryButton, SearchBox, SelectionMode, TextField, Toggle } from '@fluentui/react';
import { downloadCsv } from '../../services/exportCsv';
import { formatBytes, formatCompact, formatDate, formatDateTime, relativeTime } from '../../services/format';
import { siteType } from '../../services/siteTypes';
import { STATE_NAMES } from '../../services/StorageImport';
import { IBucket, ITenantOverview, ITopSite } from '../../services/TenantGrowth';
import { ITenantIndex } from '../../services/TenantStore';
import { capacityOf, formatTB, ICapacity } from '../../services/capacity';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { useLoader } from '../shared/hooks';
import { Card, Empty, ErrorBar, Loading, Pill, SERIES, StatTile, ViewHeader } from '../shared/ui';
import { ImportPanel } from './ImportPanel';

/**
 * Tenant dashboard: every site collection from the latest imported CSV (the admin center
 * site list), so it covers all 17,000+ sites with real storage figures. The page reads one
 * small index file; the full site table loads one snapshot file on request.
 */
export const TenantDashboardView: React.FC = () => {
  const ctx = useAdmin();
  const loader = useLoader(async () => {
    const [index, doc] = await Promise.all([ctx.tenant.loadIndex(), ctx.growth.load().catch(() => undefined)]);
    return { index, settings: doc ? doc.settings : undefined };
  }, []);
  const [importOpen, setImportOpen] = React.useState(false);
  const [editCap, setEditCap] = React.useState(false);

  if (loader.loading && !loader.data) {
    return <Loading text="Loading the tenant overview…" />;
  }
  if (loader.error || !loader.data) {
    return <ErrorBar error={loader.error || 'Nothing loaded.'} onRetry={loader.reload} />;
  }
  const index = loader.data.index;
  const o = index.overview;
  const last = index.snapshots[index.snapshots.length - 1];
  const countDeleted = index.countDeleted !== false;

  const header = (
    <ViewHeader
      title="Tenant dashboard"
      hint={o && last ? `All site collections from ${last.source || 'the last import'} · data as of ${formatDateTime(new Date(o.snapshotT * 1000))}` : 'Every site collection in the tenant, with storage, from the admin center site list.'}
    >
      <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={loader.reload}>
        Refresh
      </DefaultButton>
      <PrimaryButton iconProps={{ iconName: 'ExcelDocument' }} onClick={() => setImportOpen(true)}>
        Import CSV…
      </PrimaryButton>
    </ViewHeader>
  );
  const panel = loader.data.settings ? (
    <ImportPanel open={importOpen} onDismiss={() => setImportOpen(false)} index={index} settings={loader.data.settings} onImported={() => loader.reload()} />
  ) : null;

  if (!o) {
    return (
      <div className={styles.view}>
        {header}
        <Card>
          <Empty text={index.snapshots.length ? 'Recalculate or import again to build the tenant overview.' : 'No site list imported yet.'} />
          <p className={styles.muted} style={{ textAlign: 'center', maxWidth: 640, margin: '0 auto' }}>
            Import the CSV your flow saves from the SharePoint admin center site list (URL and storage used are required; template, Teams, archive status, deleted and last activity columns add more detail). The dashboard then covers every site collection, not just the ones search returns.
          </p>
        </Card>
        {panel}
      </div>
    );
  }

  const cap = capacityOf(o, index.capacityTB, countDeleted);
  const nonDeleted = o.activeSites + o.archivedSites;
  const teamsKnown = o.teamsSites !== undefined;
  const groupKnown = o.groupSites !== undefined;

  return (
    <div className={styles.view}>
      {header}

      <div className={styles.statGrid}>
        <StatTile icon="Globe" label="Site collections" value={nonDeleted.toLocaleString()} sub={`${o.deletedSites.toLocaleString()} more in the recycle bin`} />
        {teamsKnown ? (
          <>
            <StatTile icon="TeamsLogo" label="Teams-connected sites" value={(o.teamsSites as number).toLocaleString()} sub={`${Math.round(((o.teamsSites as number) / Math.max(1, nonDeleted)) * 100)}% of sites`} />
            <StatTile icon="SharepointLogo" label="Non-Teams sites" value={(nonDeleted - (o.teamsSites as number)).toLocaleString()} />
          </>
        ) : groupKnown ? (
          <>
            <StatTile icon="Group" label="Group-connected sites" value={(o.groupSites as number).toLocaleString()} sub="map a Teams column for Teams counts" />
            <StatTile icon="SharepointLogo" label="Other sites" value={(nonDeleted - (o.groupSites as number)).toLocaleString()} />
          </>
        ) : null}
        <StatTile icon="Database" label="Storage used" value={formatTB(cap.usedBytes)} sub={countDeleted ? 'active + recycle bin' : 'active sites'} tone="brand" />
        {cap.leftBytes !== undefined && cap.leftBytes < 0 ? (
          <StatTile icon="Warning" label="Over capacity by" value={formatTB(-cap.leftBytes)} sub={`beyond ${formatTB(cap.capacityBytes || 0)} allocated`} tone="critical" onClick={() => setEditCap(true)} />
        ) : (
          <StatTile
            icon="CloudUpload"
            label="Storage left"
            value={cap.leftBytes !== undefined ? formatTB(cap.leftBytes) : '—'}
            sub={cap.capacityBytes ? `of ${formatTB(cap.capacityBytes)} allocated` : 'set the tenant capacity'}
            tone={cap.fraction !== undefined && cap.fraction >= 0.9 ? 'critical' : cap.fraction !== undefined && cap.fraction >= 0.75 ? 'warning' : 'good'}
            onClick={() => setEditCap(true)}
          />
        )}
        <StatTile
          icon="Diagnostic"
          label="Capacity used"
          value={cap.fraction !== undefined ? `${(cap.fraction * 100).toFixed(2)}%` : '—'}
          sub={cap.capacityBytes ? `of ${formatTB(cap.capacityBytes)}` : 'set the tenant capacity'}
          tone={cap.fraction !== undefined && cap.fraction >= 1 ? 'critical' : undefined}
          onClick={() => setEditCap(true)}
        />
      </div>

      <CapacityCard overview={o} cap={cap} index={index} editing={editCap} setEditing={setEditCap} onSaved={loader.reload} />

      <div className={styles.grid}>
        <Card title="Sites by state">
          <BucketBars
            buckets={[
              { label: 'Active', count: o.activeSites, bytes: o.activeBytes },
              { label: 'Archived', count: o.archivedSites, bytes: o.archivedBytes },
              { label: 'Deleted (recycle bin)', count: o.deletedSites, bytes: o.deletedBytes }
            ]}
          />
        </Card>
        <Card title="Sites by size">
          <BucketBars buckets={o.sizeBands} />
        </Card>
        {o.activity && (
          <Card title="Last activity">
            <BucketBars buckets={o.activity} />
          </Card>
        )}
        {o.byType && (
          <Card title="Sites by type">
            <BucketBars buckets={o.byType.slice(0, 8)} />
          </Card>
        )}
      </div>

      <TopSites top={o.top} total={o.activeBytes + o.archivedBytes} />

      <AllSites index={index} />
      {panel}
    </div>
  );
};

// ---- capacity meter ---------------------------------------------------------

const CapacityCard: React.FC<{ overview: ITenantOverview; cap: ICapacity; index: ITenantIndex; editing: boolean; setEditing: (b: boolean) => void; onSaved: () => void }> = ({
  overview,
  cap,
  index,
  editing,
  setEditing,
  onSaved
}) => {
  const ctx = useAdmin();
  const [tb, setTb] = React.useState(index.capacityTB ? String(index.capacityTB) : '');
  const [countDeleted, setCountDeleted] = React.useState(index.countDeleted !== false);
  const save = async (): Promise<void> => {
    const v = Number(tb.replace(',', '.'));
    try {
      await ctx.tenant.saveCapacity(isFinite(v) && v > 0 ? v : undefined, countDeleted);
      ctx.log('Set tenant storage capacity', `${tb} TB`, true);
      setEditing(false);
      onSaved();
    } catch (e) {
      ctx.notify((e as Error).message, 'error');
    }
  };
  const usedPct = cap.fraction !== undefined ? Math.min(1, cap.fraction) : undefined;
  return (
    <Card
      title="Storage consumption"
      right={
        <DefaultButton iconProps={{ iconName: 'Edit' }} onClick={() => setEditing(!editing)}>
          {index.capacityTB ? 'Change capacity' : 'Set tenant capacity'}
        </DefaultButton>
      }
    >
      {usedPct !== undefined && cap.capacityBytes ? (
        <>
          <div className={styles.meter2} role="img" aria-label={`${formatTB(cap.usedBytes)} used of ${formatTB(cap.capacityBytes)}`}>
            <div className={styles.meterUsed} style={{ width: `${usedPct * 100}%`, background: usedPct >= 0.9 ? '#d03b3b' : usedPct >= 0.75 ? '#c27c00' : SERIES[0] }} title={`Used ${formatTB(cap.usedBytes)}`} />
            <div className={styles.meterTick} style={{ left: '90%' }} title="90%" />
          </div>
          <div className={styles.meterLegend}>
            <span>
              <span className={styles.swatch} style={{ background: SERIES[0] }} /> Used <strong>{formatTB(cap.usedBytes)}</strong> ({((cap.fraction || 0) * 100).toFixed(1)}%)
            </span>
            {(cap.leftBytes || 0) >= 0 ? (
              <span>
                <span className={styles.swatch} style={{ background: '#d9dee7' }} /> Left <strong>{formatTB(cap.leftBytes || 0)}</strong>
              </span>
            ) : (
              <span style={{ color: '#a11d1d' }}>
                Over capacity by <strong>{formatTB(-(cap.leftBytes || 0))}</strong> ({((cap.fraction || 0) * 100).toFixed(1)}% of allocated)
              </span>
            )}
            <span>
              Allocated <strong>{formatTB(cap.capacityBytes)}</strong>
            </span>
            {overview.archivedBytes > 0 && (
              <span className={styles.muted}>
                + {formatTB(overview.archivedBytes)} in Microsoft 365 Archive (billed separately)
              </span>
            )}
          </div>
        </>
      ) : (
        <MessageBar messageBarType={MessageBarType.info}>
          Enter the tenant&rsquo;s total SharePoint storage (SharePoint admin center › Reports › Storage, or Get-SPOTenant StorageQuota) to see storage left and percent used.
          {' '}Used so far: <strong>{formatTB(cap.usedBytes)}</strong>.
        </MessageBar>
      )}
      {editing && (
        <div className={styles.filters} style={{ marginTop: 12 }}>
          <TextField className={styles.field} label="Tenant storage capacity (TB)" type="number" min={0} step={0.01} value={tb} onChange={(_, v) => setTb(v || '')} />
          <Toggle label="Count deleted sites in the recycle bin as used" inlineLabel checked={countDeleted} onChange={(_, c) => setCountDeleted(!!c)} />
          <PrimaryButton onClick={save}>Save</PrimaryButton>
          <DefaultButton onClick={() => setEditing(false)}>Cancel</DefaultButton>
        </div>
      )}
    </Card>
  );
};

// ---- breakdown bars -----------------------------------------------------------

/** One-hue bars for counts with storage beside them; hover shows both figures. */
const BucketBars: React.FC<{ buckets: IBucket[] }> = ({ buckets }) => {
  const max = buckets.reduce((m, b) => Math.max(m, b.count), 0) || 1;
  const total = buckets.reduce((s, b) => s + b.count, 0) || 1;
  return (
    <div className={styles.bucketList}>
      {buckets.map((b) => (
        <div key={b.label} className={styles.bucketRow} title={`${b.label}: ${b.count.toLocaleString()} sites · ${formatBytes(b.bytes)}`}>
          <span className={styles.bucketLabel}>{b.label}</span>
          <div className={styles.barTrack}>
            <div className={styles.barFill} style={{ width: `${b.count ? Math.max(1.5, (b.count / max) * 100) : 0}%`, background: SERIES[0] }} />
          </div>
          <span className={styles.bucketValue}>
            {b.count.toLocaleString()} <span className={styles.muted}>· {Math.round((b.count / total) * 100)}%</span>
          </span>
          <span className={styles.bucketBytes}>{formatBytes(b.bytes)}</span>
        </div>
      ))}
    </div>
  );
};

// ---- top 50 -----------------------------------------------------------------

const TopSites: React.FC<{ top: ITopSite[]; total: number }> = ({ top, total }) => {
  const ctx = useAdmin();
  const [showAll, setShowAll] = React.useState(false);
  const rows = showAll ? top : top.slice(0, 15);
  const max = top.length ? top[0].bytes : 1;
  const sumTop = top.reduce((s, t) => s + t.bytes, 0);
  return (
    <Card
      title={`Top ${top.length} sites by storage`}
      right={
        <span className={styles.muted} style={{ fontWeight: 400 }}>
          together {formatTB(sumTop)} · {total ? Math.round((sumTop / total) * 100) : 0}% of all site storage
        </span>
      }
    >
      <ol className={styles.rankList}>
        {rows.map((t, i) => (
          <li key={t.url} className={styles.rankRow} title={`${t.title}\n${t.url}\n${formatBytes(t.bytes)}${t.pm ? ` · ${(t.pm / 10).toFixed(1)}% of its quota` : ''}`}>
            <span className={styles.rankNo}>{i + 1}</span>
            <div className={styles.rankMain}>
              <div className={styles.rankTop}>
                <span className={styles.rankTitle}>{t.title}</span>
                {t.state === 1 && <Pill kind="info">Archived</Pill>}
                <span className={styles.rankValue}>{formatBytes(t.bytes)}</span>
              </div>
              <div className={styles.barTrack}>
                <div className={styles.barFill} style={{ width: `${Math.max(1, (t.bytes / max) * 100)}%`, background: SERIES[0] }} />
              </div>
              <div className={styles.rankUrl}>
                {t.url}
                {t.type && <span> · {t.type}</span>}
              </div>
            </div>
            <button type="button" className={styles.link} onClick={() => ctx.setTarget({ webUrl: t.url, siteUrl: t.url, title: t.title })}>
              Manage
            </button>
          </li>
        ))}
      </ol>
      {top.length > 15 && (
        <DefaultButton onClick={() => setShowAll(!showAll)} style={{ marginTop: 8 }}>
          {showAll ? 'Show top 15' : `Show all ${top.length}`}
        </DefaultButton>
      )}
    </Card>
  );
};

// ---- all site collections ----------------------------------------------------

interface ISiteLine {
  url: string;
  title: string;
  bytes: number;
  pm: number;
  state: number;
  template: string;
  type: string;
  teams?: boolean;
  lastActivity?: Date;
}

type SortKey = 'bytes' | 'title' | 'activity';

const AllSites: React.FC<{ index: ITenantIndex }> = ({ index }) => {
  const ctx = useAdmin();
  const [open, setOpen] = React.useState(false);
  const snap = useLoader(async (): Promise<ISiteLine[]> => {
    const s = await ctx.tenant.loadLatestSnapshot();
    if (!s) {
      return [];
    }
    const origin = (/^(https:\/\/[^/]+)/i.exec(ctx.homeWebUrl) || [''])[1];
    const titles = s.titles || {};
    return s.u.map((p, i) => {
      const template = s.tp && s.tpl ? s.tpl[s.tp[i]] || '' : '';
      return {
        url: p.charAt(0) === '/' ? (p === '/' ? origin : origin + p) : p,
        title: titles[p.toLowerCase()] || p.substring(p.lastIndexOf('/') + 1) || p,
        bytes: s.b[i] || 0,
        pm: s.q[i] || 0,
        state: s.s ? s.s[i] || 0 : 0,
        template,
        type: template ? siteType(template, false) : '',
        teams: s.tm ? s.tm[i] === 1 : undefined,
        lastActivity: s.la && s.la[i] ? new Date(s.la[i] * 86400000) : undefined
      };
    });
  }, [index.snapshots.length], open);
  const [text, setText] = React.useState('');
  const [state, setState] = React.useState('all');
  const [type, setType] = React.useState('all');
  const [sort, setSort] = React.useState<SortKey>('bytes');

  const all = snap.data || [];
  const types = React.useMemo(() => {
    const m: { [k: string]: number } = {};
    all.forEach((l) => l.type && (m[l.type] = (m[l.type] || 0) + 1));
    return Object.keys(m).sort((a, b) => m[b] - m[a]);
  }, [all]);
  const q = text.trim().toLowerCase();
  const rows = React.useMemo(() => {
    const r = all.filter((l) => (state === 'all' || String(l.state) === state) && (type === 'all' || l.type === type) && (!q || l.url.toLowerCase().indexOf(q) >= 0 || l.title.toLowerCase().indexOf(q) >= 0));
    return r.sort((a, b) =>
      sort === 'title' ? a.title.localeCompare(b.title) : sort === 'activity' ? (a.lastActivity ? a.lastActivity.getTime() : 0) - (b.lastActivity ? b.lastActivity.getTime() : 0) : b.bytes - a.bytes
    );
  }, [all, state, type, q, sort]);

  const latest = index.snapshots[index.snapshots.length - 1];
  if (!open) {
    return (
      <Card title="All site collections">
        <p className={styles.muted} style={{ marginTop: 0 }}>
          Browse, search, filter and export every site collection in the latest import ({latest ? latest.sites.toLocaleString() : 0} sites). Loads one file.
        </p>
        <PrimaryButton iconProps={{ iconName: 'Table' }} onClick={() => setOpen(true)}>
          Browse all {latest ? latest.sites.toLocaleString() : ''} sites
        </PrimaryButton>
      </Card>
    );
  }
  const columns: IColumn[] = [
    {
      key: 'title',
      name: 'Site',
      minWidth: 220,
      isResizable: true,
      isSorted: sort === 'title',
      onColumnClick: () => setSort('title'),
      onRender: (l: ISiteLine) => (
        <div>
          <div style={{ fontWeight: 600 }}>{l.title}</div>
          <div className={styles.muted} style={{ fontSize: 12 }}>
            {l.url}
          </div>
        </div>
      )
    },
    { key: 'bytes', name: 'Storage', minWidth: 90, isSorted: sort === 'bytes', isSortedDescending: true, onColumnClick: () => setSort('bytes'), onRender: (l: ISiteLine) => <strong>{formatBytes(l.bytes)}</strong> },
    { key: 'pm', name: 'Of quota', minWidth: 70, onRender: (l: ISiteLine) => (l.pm ? `${(l.pm / 10).toFixed(1)}%` : '-') },
    { key: 'state', name: 'State', minWidth: 90, onRender: (l: ISiteLine) => (l.state === 2 ? <Pill kind="warning">Deleted</Pill> : l.state === 1 ? <Pill kind="info">Archived</Pill> : <Pill kind="good">Active</Pill>) },
    { key: 'type', name: 'Type', minWidth: 150, isResizable: true, onRender: (l: ISiteLine) => l.type || <span className={styles.muted}>-</span> },
    { key: 'teams', name: 'Teams', minWidth: 60, onRender: (l: ISiteLine) => (l.teams === undefined ? '-' : l.teams ? 'Yes' : 'No') },
    { key: 'activity', name: 'Last activity', minWidth: 110, isSorted: sort === 'activity', onColumnClick: () => setSort('activity'), onRender: (l: ISiteLine) => (l.lastActivity ? relativeTime(l.lastActivity) : '-') }
  ];

  return (
    <Card
      title={`All site collections (${all.length.toLocaleString()})`}
      right={
        <DefaultButton
          iconProps={{ iconName: 'Download' }}
          disabled={!rows.length}
          onClick={() =>
            downloadCsv(
              'site-collections.csv',
              ['Title', 'URL', 'Storage (bytes)', 'Storage (GB)', 'Quota used %', 'State', 'Template', 'Type', 'Teams connected', 'Last activity'],
              rows.map((l) => [l.title, l.url, l.bytes, (l.bytes / 1073741824).toFixed(2), l.pm ? (l.pm / 10).toFixed(1) : '', STATE_NAMES[l.state], l.template, l.type, l.teams === undefined ? '' : l.teams ? 'Yes' : 'No', l.lastActivity ? formatDate(l.lastActivity) : ''])
            )
          }
        >
          Export CSV
        </DefaultButton>
      }
    >
      {snap.loading && !snap.data && <Loading text="Loading the site list…" />}
      {snap.error && <ErrorBar error={snap.error} onRetry={snap.reload} />}
      {snap.data && (
        <>
          <div className={styles.filters}>
            <SearchBox className={styles.field} placeholder="Search title or URL" value={text} onChange={(_, v) => setText(v || '')} />
            <Dropdown
              className={styles.field}
              label="State"
              selectedKey={state}
              onChange={(_, o) => setState(String(o ? o.key : 'all'))}
              options={[
                { key: 'all', text: 'All states' },
                { key: '0', text: 'Active' },
                { key: '1', text: 'Archived' },
                { key: '2', text: 'Deleted' }
              ]}
            />
            {types.length > 0 && (
              <Dropdown className={styles.field} label="Type" selectedKey={type} onChange={(_, o) => setType(String(o ? o.key : 'all'))} options={[{ key: 'all', text: 'All types' }].concat(types.map((t) => ({ key: t, text: t })))} />
            )}
            <span className={styles.muted}>
              {rows.length.toLocaleString()} shown · {formatCompact(rows.reduce((s, l) => s + l.bytes, 0) / 1073741824)} GB
            </span>
          </div>
          <div style={{ maxHeight: 640, overflowY: 'auto', marginTop: 8 }} data-is-scrollable="true">
            <DetailsList items={rows} columns={columns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} getKey={(l: ISiteLine) => l.url} />
          </div>
        </>
      )}
    </Card>
  );
};
