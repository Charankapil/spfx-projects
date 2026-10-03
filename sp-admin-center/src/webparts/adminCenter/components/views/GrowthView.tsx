import * as React from 'react';
import {
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  Icon,
  IColumn,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
  PrimaryButton,
  ProgressIndicator,
  SearchBox,
  Selection,
  SelectionMode,
  TextField
} from '@fluentui/react';
import { ISiteRow } from '../../models';
import { downloadCsv } from '../../services/exportCsv';
import { formatBytes, formatDateTime, relativeTime } from '../../services/format';
import { analyzeAll, anomaliesOf, DEFAULT_SETTINGS, IGrowthSettings, ISiteGrowth, MAX_TRACKED } from '../../services/GrowthEngine';
import { ICaptureResult } from '../../services/GrowthStore';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { useLoader } from '../shared/hooks';
import { Card, Empty, ErrorBar, Kpi, Loading, Pill, Sparkline, ViewHeader, severityIcon } from '../shared/ui';

const REQS_PER_SEC = 6;

export function growthSummary(s: ISiteGrowth): string {
  return s.reasons.length ? s.reasons.join(' ') : s.status === 'baseline' ? 'Collecting a baseline. Capture again in a day or more.' : 'Normal growth.';
}

const statusPill = (s: ISiteGrowth): React.ReactNode =>
  s.status === 'critical' ? <Pill kind="critical">Growing fast</Pill> : s.status === 'warning' ? <Pill kind="warning">Watch</Pill> : s.status === 'baseline' ? <Pill>Baseline</Pill> : <Pill kind="good">Normal</Pill>;

/**
 * Tracks the storage of a watchlist of site collections over time and flags
 * unusual growth. History is one JSON file (see GrowthStore); a capture costs
 * one light request per tracked site, so the watchlist is capped.
 */
export const GrowthView: React.FC = () => {
  const ctx = useAdmin();
  const { growth, target } = ctx;
  const loader = useLoader(() => growth.load(), []);
  const [capturing, setCapturing] = React.useState(false);
  const [progress, setProgress] = React.useState({ done: 0, total: 0, failed: 0 });
  const [result, setResult] = React.useState<ICaptureResult | undefined>();
  const [pickerOpen, setPickerOpen] = React.useState(false);
  const [draft, setDraft] = React.useState<{ [k in keyof IGrowthSettings]?: string }>({});
  const cancelRef = React.useRef(false);

  const doc = loader.data;
  const all = React.useMemo(() => (doc ? analyzeAll(doc) : []), [doc]);
  const anomalies = React.useMemo(() => anomaliesOf(all), [all]);
  const anomalyCount = anomalies.length;
  React.useEffect(() => {
    ctx.setGrowthBadge(anomalyCount);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [anomalyCount]);

  if (loader.loading && !loader.data) {
    return <Loading text="Loading the storage history…" />;
  }
  if (loader.error || !doc) {
    return <ErrorBar error={loader.error || 'Nothing loaded.'} onRetry={loader.reload} />;
  }

  const tracked = all.length;
  const lastCapture = doc.lastCapture ? new Date(doc.lastCapture * 1000) : undefined;
  const staleCapture = lastCapture ? Date.now() - lastCapture.getTime() > 3 * 86400000 : false;
  const total = all.reduce((s, x) => s + x.latestBytes, 0);

  const capture = async (force: boolean): Promise<void> => {
    if (tracked > 100) {
      const ok = await ctx.confirm({
        title: 'Capture storage for ' + tracked + ' sites',
        message: `This sends one small request per site (about ${tracked}), paced to stay within SharePoint's limits. It takes roughly ${Math.ceil(tracked / REQS_PER_SEC / 60)} minute(s) and you can cancel at any time. Continue?`,
        confirmText: 'Capture'
      });
      if (!ok) {
        return;
      }
    }
    cancelRef.current = false;
    setCapturing(true);
    setResult(undefined);
    setProgress({ done: 0, total: tracked, failed: 0 });
    try {
      const r = await growth.capture({
        force,
        shouldCancel: () => cancelRef.current,
        onProgress: (done, totalSites, failed) => setProgress({ done, total: totalSites, failed })
      });
      setResult(r);
      ctx.log('Capture storage snapshot', `${r.captured} sites${r.failed.length ? `, ${r.failed.length} failed` : ''}`, !r.aborted, r.aborted ? 'Stopped after repeated failures' : undefined);
      loader.reload();
    } catch (e) {
      ctx.notify((e as Error).message, 'error');
    } finally {
      setCapturing(false);
    }
  };

  const untrack = async (s: ISiteGrowth): Promise<void> => {
    const ok = await ctx.confirm({ title: 'Stop tracking', message: `Stop tracking "${s.title}" and delete its ${s.points} stored snapshot(s)?`, confirmText: 'Stop tracking', danger: true });
    if (!ok) {
      return;
    }
    try {
      await growth.untrack([s.url]);
      ctx.log('Stop tracking storage', s.url, true);
      loader.reload();
    } catch (e) {
      ctx.notify((e as Error).message, 'error');
    }
  };

  const num = (k: keyof IGrowthSettings): string => (draft[k] !== undefined ? (draft[k] as string) : String(doc.settings[k]));
  const dirty = Object.keys(draft).length > 0;
  const saveSettings = async (): Promise<void> => {
    const next = { ...doc.settings };
    (Object.keys(draft) as Array<keyof IGrowthSettings>).forEach((k) => {
      const v = Number(draft[k]);
      if (isFinite(v) && v > 0) {
        next[k] = v;
      }
    });
    try {
      await growth.saveSettings(next);
      ctx.log('Change growth thresholds', JSON.stringify(next), true);
      setDraft({});
      loader.reload();
    } catch (e) {
      ctx.notify((e as Error).message, 'error');
    }
  };

  const columns: IColumn[] = [
    {
      key: 'site',
      name: 'Site',
      minWidth: 180,
      isResizable: true,
      onRender: (s: ISiteGrowth) => (
        <div>
          <div style={{ fontWeight: 600 }}>{s.title}</div>
          <div className={styles.muted} style={{ fontSize: 12 }}>
            {s.url}
          </div>
        </div>
      )
    },
    { key: 'size', name: 'Size now', minWidth: 80, onRender: (s: ISiteGrowth) => (s.points ? <strong>{formatBytes(s.latestBytes)}</strong> : '-') },
    { key: 'quota', name: 'Of quota', minWidth: 70, onRender: (s: ISiteGrowth) => (s.fraction ? `${Math.round(s.fraction * 100)}%` : '-') },
    {
      key: 'g7',
      name: 'Growth / week',
      minWidth: 120,
      onRender: (s: ISiteGrowth) =>
        s.status === 'baseline' ? '-' : (
          <span>
            {s.growth7Bytes >= 0 ? '+' : '-'}
            {formatBytes(Math.abs(s.growth7Bytes))} <span className={styles.muted}>({s.growth7Pct >= 0 ? '+' : ''}{Math.round(s.growth7Pct)}%)</span>
          </span>
        )
    },
    { key: 'trend', name: 'Trend', minWidth: 120, onRender: (s: ISiteGrowth) => <Sparkline values={s.spark} color={s.status === 'critical' ? '#c4314b' : s.status === 'warning' ? '#b87400' : '#2563eb'} /> },
    { key: 'status', name: 'Status', minWidth: 100, onRender: (s: ISiteGrowth) => statusPill(s) },
    {
      key: 'act',
      name: '',
      minWidth: 130,
      onRender: (s: ISiteGrowth) => (
        <div className={styles.actions}>
          <button type="button" className={styles.link} onClick={() => ctx.setTarget({ webUrl: s.url, siteUrl: s.url, title: s.title })}>
            Manage
          </button>
          <button type="button" className={styles.link} onClick={() => untrack(s)}>
            Remove
          </button>
        </div>
      )
    }
  ];

  return (
    <div className={styles.view}>
      <ViewHeader title="Storage growth" hint="Snapshots of each tracked site collection's storage over time, with alerts when growth looks unusual.">
        <PrimaryButton iconProps={{ iconName: 'Camera' }} onClick={() => capture(false)} disabled={capturing || tracked === 0}>
          Capture now
        </PrimaryButton>
        <DefaultButton iconProps={{ iconName: 'Add' }} onClick={() => setPickerOpen(true)} disabled={capturing}>
          Choose sites
        </DefaultButton>
        <DefaultButton
          iconProps={{ iconName: 'Download' }}
          disabled={tracked === 0}
          onClick={() => downloadCsv('storage-growth.csv', ['Site', 'URL', 'Size (bytes)', 'Quota used', 'Growth per week (bytes)', 'Growth per week %', 'Status', 'Snapshots', 'Notes'], all.map((s) => [s.title, s.url, s.latestBytes, s.fraction ? Math.round(s.fraction * 100) + '%' : '', Math.round(s.growth7Bytes), Math.round(s.growth7Pct), s.status, s.points, growthSummary(s)]))}
        >
          Export CSV
        </DefaultButton>
      </ViewHeader>

      {capturing && (
        <Card>
          <ProgressIndicator label="Capturing storage snapshots…" description={`${progress.done} of ${progress.total} sites${progress.failed ? ` · ${progress.failed} could not be read` : ''}`} percentComplete={progress.total ? progress.done / progress.total : 0} />
          <DefaultButton onClick={() => (cancelRef.current = true)}>Cancel (keeps what is captured)</DefaultButton>
        </Card>
      )}

      {result && !capturing && (
        <MessageBar messageBarType={result.aborted ? MessageBarType.error : result.failed.length ? MessageBarType.warning : MessageBarType.success} onDismiss={() => setResult(undefined)} isMultiline>
          {result.aborted
            ? `Stopped after repeated failures (${result.failed.length} sites). Check your permissions and try again later. `
            : `Captured ${result.captured} site${result.captured === 1 ? '' : 's'}${result.skipped ? `, ${result.skipped} already up to date` : ''}${result.cancelled ? ' (cancelled)' : ''}. `}
          {result.failed.length > 0 && `${result.failed.length} could not be read, for example ${result.failed[0].url}: ${result.failed[0].message}`}
        </MessageBar>
      )}

      {tracked === 0 ? (
        <Card>
          <Empty text="No sites are tracked yet." />
          <div className={styles.actions} style={{ justifyContent: 'center' }}>
            <PrimaryButton
              onClick={async () => {
                try {
                  await growth.track([{ url: target.siteUrl, title: target.title }]);
                  loader.reload();
                } catch (e) {
                  ctx.notify((e as Error).message, 'error');
                }
              }}
            >
              Track {target.title}
            </PrimaryButton>
            <DefaultButton onClick={() => setPickerOpen(true)}>Choose sites…</DefaultButton>
          </div>
          <p className={styles.muted} style={{ textAlign: 'center' }}>
            Capture a snapshot, then capture again after a day or more. Growth is compared across snapshots, so alerts start once there is history.
          </p>
        </Card>
      ) : (
        <>
          <div className={styles.kpiRow}>
            <Kpi label="Tracked sites" value={String(tracked)} sub={`of ${MAX_TRACKED} maximum`} />
            <Kpi label="Tracked storage" value={formatBytes(total)} />
            <Kpi label="Unusual growth" value={String(anomalies.length)} sub={anomalies.length ? 'need a look' : 'all normal'} />
            <Kpi label="Last capture" value={lastCapture ? relativeTime(lastCapture) : 'never'} sub={lastCapture ? formatDateTime(lastCapture) : undefined} />
          </div>

          {staleCapture && (
            <MessageBar messageBarType={MessageBarType.warning}>The last snapshot is more than 3 days old, so recent growth may not be visible yet. Capture again.</MessageBar>
          )}

          {anomalies.length > 0 && (
            <Card title="Sites that grew unusually fast">
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                {anomalies.map((s) => {
                  const si = severityIcon(s.status === 'critical' ? 'critical' : 'warning');
                  return (
                    <div key={s.url} className={styles.finding}>
                      <Icon iconName={si.icon} className={styles.findingIcon} style={{ color: si.color }} aria-hidden="true" />
                      <div style={{ minWidth: 0 }}>
                        <div className={styles.findingTitle}>{s.title}</div>
                        <div className={styles.findingDetail}>
                          {formatBytes(s.latestBytes)} now · {growthSummary(s)}
                        </div>
                      </div>
                      <div className={styles.findingGo} style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
                        <Sparkline values={s.spark} color={si.color} />
                        <DefaultButton onClick={() => ctx.setTarget({ webUrl: s.url, siteUrl: s.url, title: s.title })}>Manage</DefaultButton>
                      </div>
                    </div>
                  );
                })}
              </div>
            </Card>
          )}

          <div className={styles.tableWrap}>
            <DetailsList items={all} columns={columns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} />
          </div>
        </>
      )}

      <Card
        title="When to raise an alert"
        right={
          <div className={styles.actions}>
            <DefaultButton disabled={!dirty} onClick={saveSettings}>
              Save thresholds
            </DefaultButton>
            <DefaultButton onClick={() => setDraft(Object.keys(DEFAULT_SETTINGS).reduce((a, k) => ({ ...a, [k]: String(DEFAULT_SETTINGS[k as keyof IGrowthSettings]) }), {}))}>Reset to defaults</DefaultButton>
          </div>
        }
      >
        <div className={styles.filters}>
          <TextField className={styles.field} label="Growth of at least (% per week)" type="number" min={1} value={num('pctPerWeek')} onChange={(_, v) => setDraft({ ...draft, pctPerWeek: v || '' })} />
          <TextField className={styles.field} label="…and at least (GB per week)" type="number" min={1} value={num('gbPerWeek')} onChange={(_, v) => setDraft({ ...draft, gbPerWeek: v || '' })} />
          <TextField className={styles.field} label="Always alert above (GB per week)" type="number" min={1} value={num('hugeGbPerWeek')} onChange={(_, v) => setDraft({ ...draft, hugeGbPerWeek: v || '' })} />
          <TextField className={styles.field} label="Quota full within (days)" type="number" min={1} value={num('forecastDays')} onChange={(_, v) => setDraft({ ...draft, forecastDays: v || '' })} />
        </div>
        <p className={styles.muted} style={{ marginBottom: 0 }}>
          Rates are scaled to a week from the snapshots you have, so uneven gaps between captures are fine. A site is also flagged when its latest daily growth is far above its own usual pattern. Thresholds are shared by every admin. The history is stored in <em>Site Assets/admin-center-storage-history.json</em> on this site.
        </p>
      </Card>

      <SitePicker open={pickerOpen} onDismiss={() => setPickerOpen(false)} trackedUrls={all.map((s) => s.url)} onTracked={() => loader.reload()} />
    </div>
  );
};

// ---- choose which sites to track -----------------------------------------------

const SitePicker: React.FC<{ open: boolean; onDismiss: () => void; trackedUrls: string[]; onTracked: () => void }> = ({ open, onDismiss, trackedUrls, onTracked }) => {
  const ctx = useAdmin();
  const [sites, setSites] = React.useState<ISiteRow[] | undefined>();
  const [busy, setBusy] = React.useState(false);
  const [text, setText] = React.useState('');
  const [url, setUrl] = React.useState('');
  const [error, setError] = React.useState<string | undefined>();
  const [selected, setSelected] = React.useState<ISiteRow[]>([]);
  const selection = React.useMemo(() => new Selection({ getKey: (i) => (i as ISiteRow).url, onSelectionChanged: () => setSelected(selection.getSelection() as ISiteRow[]) }), []);

  const q = text.trim().toLowerCase();
  // Memoised: DetailsList resets its selection whenever it receives a new items array.
  const rows = React.useMemo(
    () => (sites || []).filter((s) => s.manageable && trackedUrls.indexOf(s.url) < 0 && (!q || s.title.toLowerCase().indexOf(q) >= 0 || s.url.toLowerCase().indexOf(q) >= 0)).slice(0, 500),
    [sites, trackedUrls, q]
  );

  const loadSites = async (): Promise<void> => {
    setBusy(true);
    setError(undefined);
    try {
      const r = await ctx.search.listSites(ctx.homeWebUrl, 10);
      setSites(r.sites);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const trackSelected = async (): Promise<void> => {
    setBusy(true);
    try {
      const r = await ctx.growth.track(selected.map((s) => ({ url: s.url, title: s.title })));
      ctx.log('Track storage growth', `${r.added} site(s)`, true);
      ctx.notify(`${r.added} site${r.added === 1 ? '' : 's'} added.${r.capped ? ` The limit of ${MAX_TRACKED} tracked sites was reached.` : ''}`, r.capped ? 'info' : 'success');
      selection.setAllSelected(false);
      onTracked();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const addByUrl = async (): Promise<void> => {
    setError(undefined);
    const u = url.trim().replace(/\/+$/, '');
    try {
      ctx.client.assertSameOrigin(u);
    } catch {
      setError('Enter a SharePoint site collection address on this tenant.');
      return;
    }
    setBusy(true);
    try {
      await ctx.api.getSiteUsage(u); // verifies that you can read this site's storage
      const r = await ctx.growth.track([{ url: u, title: u.substring(u.lastIndexOf('/') + 1) || u }]);
      ctx.notify(r.added ? 'Site added.' : 'That site is already tracked.', 'success');
      setUrl('');
      onTracked();
    } catch (e) {
      setError('Could not add that site: ' + (e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  return (
    <Panel isOpen={open} onDismiss={onDismiss} type={PanelType.medium} headerText="Choose sites to track" closeButtonAriaLabel="Close" isLightDismiss>
      <p className={styles.muted}>
        Each tracked site costs one small request per capture, so up to {MAX_TRACKED} sites can be tracked. Start with the biggest or busiest.
      </p>
      <div className={styles.filters}>
        <TextField className={styles.field} label="Add by address" placeholder="https://tenant.sharepoint.com/sites/finance" value={url} onChange={(_, v) => setUrl(v || '')} onKeyDown={(e) => e.key === 'Enter' && !busy && url.trim() && addByUrl()} />
        <PrimaryButton onClick={addByUrl} disabled={busy || !url.trim()}>
          Add
        </PrimaryButton>
      </div>
      {error && <ErrorBar error={error} />}
      <h3 style={{ fontSize: 14, marginTop: 20 }}>Or pick from the sites you can see</h3>
      {!sites ? (
        <DefaultButton iconProps={{ iconName: 'Search' }} onClick={loadSites} disabled={busy}>
          {busy ? 'Reading the search index…' : 'Load site list'}
        </DefaultButton>
      ) : (
        <>
          <div className={styles.filters}>
            <SearchBox className={styles.field} placeholder="Filter by title or URL" value={text} onChange={(_, v) => setText(v || '')} />
            <PrimaryButton onClick={trackSelected} disabled={busy || selected.length === 0}>
              Track {selected.length || ''} selected
            </PrimaryButton>
          </div>
          <div className={styles.tableWrap} style={{ marginTop: 10 }}>
            <DetailsList
              items={rows}
              selection={selection}
              selectionMode={SelectionMode.multiple}
              layoutMode={DetailsListLayoutMode.justified}
              columns={[
                {
                  key: 't',
                  name: 'Site',
                  minWidth: 200,
                  isResizable: true,
                  onRender: (s: ISiteRow) => (
                    <div>
                      <div style={{ fontWeight: 600 }}>{s.title}</div>
                      <div className={styles.muted} style={{ fontSize: 12 }}>
                        {s.url}
                      </div>
                    </div>
                  )
                }
              ]}
            />
            {rows.length === 0 && <div className={styles.empty}>No more sites to add.</div>}
          </div>
        </>
      )}
    </Panel>
  );
};
