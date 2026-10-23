import * as React from 'react';
import { DefaultButton, DetailsList, DetailsListLayoutMode, IColumn, MessageBar, MessageBarType, PrimaryButton, SelectionMode } from '@fluentui/react';
import { parseCsv } from '../../services/csv';
import { downloadCsv } from '../../services/exportCsv';
import { formatPercent, formatBytes, formatCompact, formatDateTime, relativeTime } from '../../services/format';
import { IGrowthDoc, IGrowthSettings } from '../../services/GrowthEngine';
import { ISegment, ISummaryRow } from '../../services/TenantGrowth';
import { ITenantIndex } from '../../services/TenantStore';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { BarList, Card, Donut, Empty, ErrorBar, Kpi, Loading, Pill, Sparkline } from '../shared/ui';
import { growthSummary, statusPill } from './growthShared';
import { ImportPanel } from './ImportPanel';

/**
 * All sites at once, from a CSV your flow writes. Analysis happens once at
 * import time and is stored with the index, so opening this tab is one small
 * file read however many sites the tenant has.
 */
export const TenantTab: React.FC<{ doc: IGrowthDoc; index?: ITenantIndex; indexError?: string; reload: () => void }> = ({ doc, index, indexError, reload }) => {
  const ctx = useAdmin();
  const { tenant } = ctx;
  const [importOpen, setImportOpen] = React.useState(false);
  const [busy, setBusy] = React.useState('');

  if (indexError) {
    return <ErrorBar error={indexError} onRetry={reload} />;
  }
  if (!index) {
    return <Loading />;
  }
  const summary = index.summary;
  const snapshots = index.snapshots;
  const last = snapshots.length ? new Date(snapshots[snapshots.length - 1].t * 1000) : undefined;

  // Read the saved thresholds at the moment of use: the page may not have reloaded them after a save yet.
  const freshSettings = async (): Promise<IGrowthSettings> => (await ctx.growth.load(true)).settings;

  const importNewest = async (): Promise<void> => {
    if (!index.folder || !index.mapping) {
      setImportOpen(true);
      return;
    }
    setBusy('Looking for the newest CSV…');
    try {
      const files = await tenant.listCsvFiles(index.folder);
      const f = files[0];
      if (!f) {
        ctx.notify('There is no CSV in the folder the last import came from.', 'info');
        return;
      }
      const modified = f.modified ? Math.floor(f.modified.getTime() / 1000) : Math.floor(Date.now() / 1000);
      if (index.lastFile && index.lastFile.name === f.name && index.lastFile.modified >= modified) {
        ctx.notify(`${f.name} has not changed since the last import, so there is nothing new to import.`, 'info');
        return;
      }
      setBusy(`Reading ${f.name}…`);
      const text = await tenant.readCsv(f.serverRelativeUrl);
      const headers = parseCsv(text)[0] || [];
      if (headers.indexOf(index.mapping.url) < 0 || headers.indexOf(index.mapping.storage) < 0) {
        ctx.notify('The columns in this file differ from the last import, so please review the mapping.', 'info');
        setImportOpen(true);
        return;
      }
      const r = await tenant.importCsv(text, index.mapping, { t: modified, fileName: f.name, folder: index.folder, modified }, await freshSettings(), setBusy);
      ctx.log('Import storage CSV', `${f.name}: ${r.stats.kept} sites`, true);
      ctx.notify(`Imported ${r.stats.kept.toLocaleString()} sites from ${f.name}.`, 'success');
      reload();
    } catch (e) {
      ctx.log('Import storage CSV', index.folder || '', false, (e as Error).message);
      ctx.notify((e as Error).message, 'error');
    } finally {
      setBusy('');
    }
  };

  const recalc = async (): Promise<void> => {
    setBusy('Recalculating…');
    try {
      await tenant.recalculate(await freshSettings());
      ctx.notify('Results recalculated with the current thresholds.', 'success');
      reload();
    } catch (e) {
      ctx.notify((e as Error).message, 'error');
    } finally {
      setBusy('');
    }
  };

  const columns: IColumn[] = [
    {
      key: 'site',
      name: 'Site',
      minWidth: 200,
      isResizable: true,
      onRender: (s: ISummaryRow) => (
        <div>
          <div style={{ fontWeight: 600 }}>{s.title}</div>
          <div className={styles.muted} style={{ fontSize: 12 }}>
            {s.url}
          </div>
        </div>
      )
    },
    { key: 'size', name: 'Size now', minWidth: 80, onRender: (s: ISummaryRow) => <strong>{formatBytes(s.latestBytes)}</strong> },
    { key: 'quota', name: 'Of quota', minWidth: 70, onRender: (s: ISummaryRow) => formatPercent(s.fraction) },
    {
      key: 'g7',
      name: 'Growth / week',
      minWidth: 120,
      onRender: (s: ISummaryRow) => (
        <span>
          {s.growth7Bytes >= 0 ? '+' : '-'}
          {formatBytes(Math.abs(s.growth7Bytes))} <span className={styles.muted}>({s.growth7Pct >= 0 ? '+' : ''}{Math.round(s.growth7Pct)}%)</span>
        </span>
      )
    },
    { key: 'trend', name: 'Trend', minWidth: 120, onRender: (s: ISummaryRow) => <Sparkline values={s.spark} color={s.status === 'critical' ? '#c4314b' : '#b87400'} /> },
    {
      key: 'status',
      name: 'Status',
      minWidth: 110,
      onRender: (s: ISummaryRow) => (
        <span>
          {statusPill(s)} {s.state === 1 && <Pill kind="info">Archived</Pill>}
        </span>
      )
    },
    { key: 'why', name: 'Why', minWidth: 260, isResizable: true, onRender: (s: ISummaryRow) => <span className={styles.muted}>{growthSummary(s)}</span> },
    {
      key: 'act',
      name: '',
      minWidth: 70,
      onRender: (s: ISummaryRow) =>
        s.url.toLowerCase().indexOf((/^(https:\/\/[^/]+)/i.exec(ctx.homeWebUrl) || [''])[1].toLowerCase()) === 0 ? (
          <button type="button" className={styles.link} onClick={() => ctx.setTarget({ webUrl: s.url, siteUrl: s.url, title: s.title })}>
            Manage
          </button>
        ) : null
    }
  ];

  return (
    <div className={styles.view}>
      <div className={styles.actions}>
        <PrimaryButton iconProps={{ iconName: 'ExcelDocument' }} onClick={() => setImportOpen(true)} disabled={!!busy}>
          Import CSV…
        </PrimaryButton>
        {index.folder && index.mapping && (
          <DefaultButton iconProps={{ iconName: 'Download' }} onClick={importNewest} disabled={!!busy}>
            Import newest from last folder
          </DefaultButton>
        )}
        {snapshots.length > 0 && (
          <DefaultButton iconProps={{ iconName: 'Calculator' }} onClick={recalc} disabled={!!busy}>
            Recalculate with current thresholds
          </DefaultButton>
        )}
        {summary && (
          <DefaultButton
            iconProps={{ iconName: 'Download' }}
            onClick={() => downloadCsv('storage-growth-tenant.csv', ['Site', 'URL', 'Size (bytes)', 'Quota used', 'Growth per week (bytes)', 'Growth per week %', 'Status', 'Notes'], summary.anomalies.map((s) => [s.title, s.url, s.latestBytes, s.fraction ? Math.round(s.fraction * 100) + '%' : '', Math.round(s.growth7Bytes), Math.round(s.growth7Pct), s.status, growthSummary(s)]))}
          >
            Export alerts CSV
          </DefaultButton>
        )}
      </div>

      {busy && <Loading text={busy} />}

      {snapshots.length === 0 ? (
        <Card>
          <Empty text="No storage snapshot imported yet." />
          <p className={styles.muted} style={{ textAlign: 'center' }}>
            Import the CSV your flow saves from the SharePoint admin center site list. Import it again on a later day (or after the flow runs again) and growth between the snapshots is analysed for every site in the file.
          </p>
        </Card>
      ) : (
        summary && (
          <>
            {snapshots.length < 2 && (
              <MessageBar messageBarType={MessageBarType.info}>One snapshot gives a baseline. Alerts start once there is a second snapshot at least a day later.</MessageBar>
            )}
            {last && Date.now() - last.getTime() > 8 * 86400000 && (
              <MessageBar messageBarType={MessageBarType.warning}>The newest snapshot is more than 8 days old, so recent growth may not show yet. Import the latest CSV.</MessageBar>
            )}
            {summary.segments ? (
              <SegmentsCard segments={summary.segments} total={summary.totalBytes} />
            ) : (
              <MessageBar messageBarType={MessageBarType.info}>Press &ldquo;Recalculate with current thresholds&rdquo; (or import again) to see storage split by active, archived and deleted sites.</MessageBar>
            )}
            <div className={styles.kpiRow}>
              <Kpi label="Sites" value={formatCompact(summary.sites)} sub="in the latest snapshot" />
              <Kpi label="Total storage" value={formatBytes(summary.totalBytes)} />
              <Kpi label="Net growth / week" value={(summary.growth7Bytes >= 0 ? '+' : '-') + formatBytes(Math.abs(summary.growth7Bytes))} sub="all sites combined" />
              <Kpi label="Unusual growth" value={String(summary.anomalies.length)} sub={summary.anomalies.length ? 'sites to review' : 'all normal'} />
              <Kpi label="Last snapshot" value={last ? relativeTime(last) : '-'} sub={`${snapshots.length} snapshot${snapshots.length === 1 ? '' : 's'} kept`} />
            </div>

            {summary.anomalies.length > 0 ? (
              <div className={styles.tableWrap}>
                <DetailsList items={summary.anomalies} columns={columns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} getKey={(s: ISummaryRow) => s.url} />
              </div>
            ) : (
              <Card>
                <Empty text="No site is growing unusually fast." />
              </Card>
            )}

            <div className={styles.grid}>
              <Card title="Fastest growing sites (per week)">
                {summary.growers.length ? <BarList rows={summary.growers.slice(0, 10).map((s) => ({ label: s.title, value: s.growth7Bytes, display: '+' + formatBytes(s.growth7Bytes) }))} /> : <Empty text="Needs a second snapshot." />}
              </Card>
              <Card title="Largest sites">
                <BarList rows={summary.largest.slice(0, 10).map((s) => ({ label: s.title, value: s.latestBytes, display: formatBytes(s.latestBytes) }))} />
              </Card>
            </div>
            <p className={styles.muted}>
              Results analysed {formatDateTime(new Date(summary.at * 1000))} from {summary.basedOnSnapshots} snapshot{summary.basedOnSnapshots === 1 ? '' : 's'}. Showing the {summary.anomalies.length >= 300 ? 'first 300' : summary.anomalies.length} flagged sites.
            </p>
          </>
        )
      )}

      <ImportPanel open={importOpen} onDismiss={() => setImportOpen(false)} index={index} settings={doc.settings} onImported={() => reload()} />
    </div>
  );
};

const STATE_COLORS = { active: '#2563eb', archived: '#0d9488', deleted: '#b45309' };

/** Storage split by site state: where the space actually goes, and what deleting or archiving could free. */
const SegmentsCard: React.FC<{ segments: { active: ISegment; archived: ISegment; deleted: ISegment }; total: number }> = ({ segments, total }) => {
  const parts: Array<{ key: 'active' | 'archived' | 'deleted'; label: string; hint: string }> = [
    { key: 'active', label: 'Active sites', hint: 'in normal use' },
    { key: 'archived', label: 'Archived sites', hint: 'Microsoft 365 Archive' },
    { key: 'deleted', label: 'Deleted sites', hint: 'still in the site recycle bin' }
  ];
  const pct = (b: number): string => (total > 0 ? `${((b / total) * 100).toFixed(b / total < 0.1 ? 1 : 0)}%` : '-');
  return (
    <Card title="Storage by site state">
      <div className={styles.donutWrap} style={{ alignItems: 'stretch' }}>
        <Donut
          slices={parts.map((p) => ({ label: p.label, value: segments[p.key].bytes, color: STATE_COLORS[p.key] }))}
          centre={formatBytes(total)}
          centreSub="in total"
          format={formatBytes}
        />
        <div className={styles.kpiRow} style={{ flex: 1, minWidth: 260 }}>
          {parts.map((p) => {
            const s = segments[p.key];
            return (
              <div key={p.key} className={styles.kpi} style={{ borderTop: `3px solid ${STATE_COLORS[p.key]}` }}>
                <div className={styles.kpiLabel}>{p.label}</div>
                <div className={styles.kpiValue}>{formatCompact(s.sites)}</div>
                <div className={styles.kpiSub}>
                  {formatBytes(s.bytes)} · {pct(s.bytes)} of storage
                </div>
                <div className={styles.kpiSub}>
                  {s.growth7Bytes >= 0 ? '+' : '-'}
                  {formatBytes(Math.abs(s.growth7Bytes))} / week · {p.hint}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </Card>
  );
};
