import * as React from 'react';
import { DefaultButton, DetailsList, DetailsListLayoutMode, Dropdown, IColumn, MessageBar, MessageBarType, PrimaryButton, SearchBox, SelectionMode } from '@fluentui/react';
import { ISiteRow } from '../../models';
import { downloadCsv } from '../../services/exportCsv';
import { siteType } from '../../services/siteTypes';
import { daysAgo, formatBytes, formatDate, relativeTime } from '../../services/format';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { Card, Donut, Empty, ErrorBar, Kpi, Loading, Pill, ViewHeader } from '../shared/ui';


/**
 * Tenant inventory from the search index: up to 5,000 site collections in at
 * most 10 paced search requests. Security-trimmed, so it lists the sites the
 * signed-in admin can access. "Manage" retargets every other view at a site.
 */
export const SitesView: React.FC = () => {
  const ctx = useAdmin();
  const [sites, setSites] = React.useState<ISiteRow[] | undefined>();
  const [truncated, setTruncated] = React.useState(false);
  const [busy, setBusy] = React.useState(false);
  const [loaded, setLoaded] = React.useState(0);
  const [sizes, setSizes] = React.useState<{ [urlLower: string]: number } | undefined>();
  const [sortBySize, setSortBySize] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>();
  const [text, setText] = React.useState('');
  const [type, setType] = React.useState('all');
  const [inactive, setInactive] = React.useState(0);

  const load = async (): Promise<void> => {
    setBusy(true);
    setError(undefined);
    setLoaded(0);
    try {
      ctx.client.clearCache();
      const r = await ctx.search.listSites(ctx.homeWebUrl, 100000, setLoaded);
      setSites(r.sites);
      setTruncated(r.truncated);
      // Storage per site from the latest CSV import, if there is one (one request).
      try {
        const snap = await ctx.tenant.loadLatestSnapshot();
        if (snap) {
          const origin = (/^(https:\/\/[^/]+)/i.exec(ctx.homeWebUrl) || [''])[1];
          const m: { [k: string]: number } = {};
          snap.u.forEach((p, i) => (m[(p.charAt(0) === '/' ? (p === '/' ? origin : origin + p) : p).toLowerCase()] = snap.b[i]));
          setSizes(m);
        }
      } catch {
        /* storage column is optional */
      }
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const typed = React.useMemo(() => (sites || []).map((s) => ({ ...s, type: siteType(s.template, s.groupConnected) })), [sites]);
  const types = React.useMemo(() => {
    const m: { [k: string]: number } = {};
    typed.forEach((s) => {
      m[s.type] = (m[s.type] || 0) + 1;
    });
    return Object.keys(m)
      .map((k) => ({ label: k, value: m[k] }))
      .sort((a, b) => b.value - a.value);
  }, [typed]);

  const sizeOf = (s: ISiteRow): number | undefined => (sizes ? sizes[s.url.toLowerCase()] : undefined);
  const filteredRaw = typed.filter((s) => {
    if (type !== 'all' && s.type !== type) {
      return false;
    }
    if (inactive > 0) {
      const d = daysAgo(s.lastModified);
      if (d === undefined || d < inactive) {
        return false;
      }
    }
    const q = text.trim().toLowerCase();
    return !q || s.title.toLowerCase().indexOf(q) >= 0 || s.url.toLowerCase().indexOf(q) >= 0;
  });

  const filtered = sortBySize && sizes ? filteredRaw.slice().sort((a, b) => (sizeOf(b) || 0) - (sizeOf(a) || 0)) : filteredRaw;
  const stale180 = typed.filter((s) => (daysAgo(s.lastModified) || 0) >= 180).length;
  const noGroup = typed.filter((s) => s.type === 'Team site (no group)' || s.type === 'Classic team site').length;

  const columns: IColumn[] = [
    {
      key: 'title',
      name: 'Site',
      minWidth: 180,
      isResizable: true,
      onRender: (s: ISiteRow) => (
        <div>
          <div style={{ fontWeight: 600 }}>{s.title}</div>
          <div className={styles.muted} style={{ fontSize: 12 }}>
            {s.url}
          </div>
        </div>
      )
    },
    { key: 'type', name: 'Type', minWidth: 140, isResizable: true, onRender: (s: ISiteRow & { type: string }) => s.type },
    ...(sizes
      ? [
          {
            key: 'size',
            name: 'Storage',
            minWidth: 80,
            isSorted: sortBySize,
            isSortedDescending: true,
            onColumnClick: () => setSortBySize(!sortBySize),
            onRender: (s: ISiteRow) => {
              const b = sizeOf(s);
              return b === undefined ? <span className={styles.muted}>-</span> : formatBytes(b);
            }
          } as IColumn
        ]
      : []),
    { key: 'created', name: 'Created', minWidth: 90, onRender: (s: ISiteRow) => formatDate(s.created) },
    {
      key: 'last',
      name: 'Last activity',
      minWidth: 110,
      onRender: (s: ISiteRow) => {
        const d = daysAgo(s.lastModified);
        return d !== undefined && d >= 180 ? <Pill kind="warning">{relativeTime(s.lastModified)}</Pill> : relativeTime(s.lastModified);
      }
    },
    {
      key: 'act',
      name: '',
      minWidth: 110,
      onRender: (s: ISiteRow) =>
        s.manageable ? (
          <DefaultButton
            text="Manage"
            iconProps={{ iconName: 'Settings' }}
            onClick={() => ctx.setTarget({ webUrl: s.url, siteUrl: s.url, title: s.title })}
          />
        ) : (
          <span className={styles.muted}>Other domain</span>
        )
    }
  ];

  return (
    <div className={styles.view}>
      <ViewHeader title="Sites" hint="Every site collection you can see in the tenant, from the search index. Pick one to administer it from this same screen.">
        <PrimaryButton iconProps={{ iconName: sites ? 'Refresh' : 'Search' }} onClick={load} disabled={busy}>
          {sites ? 'Reload inventory' : 'Load site inventory'}
        </PrimaryButton>
        {sites && (
          <DefaultButton
            iconProps={{ iconName: 'Download' }}
            onClick={() =>
              downloadCsv(
                'sites.csv',
                ['Title', 'URL', 'Type', 'Template', 'Storage (bytes)', 'Created', 'Last activity'],
                filtered.map((s) => [s.title, s.url, s.type, s.template, sizeOf(s), s.created, s.lastModified])
              )
            }
          >
            Export CSV
          </DefaultButton>
        )}
      </ViewHeader>

      {error && <ErrorBar error={error} onRetry={load} />}
      {busy && <Loading text={`Reading the search index… ${loaded.toLocaleString()} sites so far (500 per request)`} />}

      {!sites && !busy && !error && (
        <Card>
          <Empty text="Load the inventory to see every site collection you can access. It takes one search request per 500 sites (about 35 for 17,000), paced to stay well inside SharePoint's limits." />
        </Card>
      )}

      {sites && (
        <>
          {truncated && (
            <MessageBar messageBarType={MessageBarType.warning}>The search index stopped returning results before the end of the list. Showing {typed.length.toLocaleString()} sites.</MessageBar>
          )}
          <div className={styles.kpiRow}>
            <Kpi label="Sites" value={typed.length.toLocaleString()} sub="visible to you" />
            <Kpi label="Inactive 180+ days" value={stale180.toLocaleString()} sub="archive candidates" />
            <Kpi label="Without an M365 group" value={noGroup.toLocaleString()} sub="harder to govern" />
          </div>
          <div className={styles.grid}>
            <Card title="By type">
              <Donut slices={types.slice(0, 7)} centre={typed.length.toLocaleString()} centreSub="sites" />
            </Card>
          </div>
          <div className={styles.filters}>
            <SearchBox className={styles.field} placeholder="Search title or URL" value={text} onChange={(_, v) => setText(v || '')} />
            <Dropdown
              className={styles.field}
              label="Type"
              selectedKey={type}
              onChange={(_, o) => setType(String(o ? o.key : 'all'))}
              options={[{ key: 'all', text: 'All types' }].concat(types.map((t) => ({ key: t.label, text: `${t.label} (${t.value})` })))}
            />
            <Dropdown
              className={styles.field}
              label="Inactive for"
              selectedKey={inactive}
              onChange={(_, o) => setInactive(Number(o ? o.key : 0))}
              options={[
                { key: 0, text: 'Any' },
                { key: 90, text: '90+ days' },
                { key: 180, text: '180+ days' },
                { key: 365, text: '1+ year' }
              ]}
            />
            <span className={styles.muted}>
              {filtered.length.toLocaleString()} of {typed.length.toLocaleString()}
            </span>
          </div>
          <div className={styles.tableWrap}>
            {/* DetailsList renders only the rows on screen, so all 17,000+ sites scroll smoothly. */}
            <DetailsList items={filtered} columns={columns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} />
          </div>
        </>
      )}
    </div>
  );
};
