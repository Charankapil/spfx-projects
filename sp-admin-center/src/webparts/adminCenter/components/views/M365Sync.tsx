import * as React from 'react';
import { DefaultButton, MessageBar, MessageBarType } from '@fluentui/react';
import { DEFAULT_SETTINGS, IGrowthSettings } from '../../services/GrowthEngine';
import { ITrendPoint, reportTime, USAGE_MAPPING } from '../../services/GraphApi';
import { formatTB } from '../../services/capacity';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { graphErrorText } from '../shared/graph';
import { Card, SERIES } from '../shared/ui';

/**
 * Pulls the SharePoint site usage report (every site collection with storage, quota, last
 * activity, template and deleted state) from Microsoft Graph and stores it as a tenant snapshot,
 * the same as a CSV import: the dashboard, growth alerts and history all use it.
 * About one request per 200 sites (17,000 sites = ~85 paced requests).
 */
export const SyncFromM365Button: React.FC<{ settings?: IGrowthSettings; onSynced: () => void }> = ({ settings, onSynced }) => {
  const ctx = useAdmin();
  const [busy, setBusy] = React.useState<string | undefined>();

  const sync = async (): Promise<void> => {
    setBusy('Reading the usage report…');
    try {
      const table = await ctx.graph.siteUsage('D7', (n) => setBusy(`Reading the usage report… ${n.toLocaleString()} sites`));
      if (table.sites === 0) {
        throw new Error(
          table.concealed
            ? `Microsoft 365 returned ${table.concealed.toLocaleString()} sites but their addresses are hidden. A Global or Reports administrator must clear "Display concealed user, group, and site names in all reports" (Microsoft 365 admin center > Settings > Org settings > Reports).`
            : 'The usage report returned no sites.'
        );
      }
      const nowSec = Math.floor(Date.now() / 1000);
      const t = reportTime(table.refreshDate, nowSec);
      const r = await ctx.tenant.importCsv(table.rows, USAGE_MAPPING, { t, fileName: 'Microsoft 365 usage report', keepMapping: true }, settings || DEFAULT_SETTINGS, (step) => setBusy(step));
      ctx.log('Sync tenant dashboard from Microsoft 365', `${r.stats.kept} sites`, true);
      ctx.notify(
        `Synced ${r.stats.kept.toLocaleString()} sites from Microsoft 365 (data as of ${table.refreshDate || 'the latest report'})${r.summary.anomalies.length ? `; ${r.summary.anomalies.length} growing unusually fast` : ''}.${table.concealed ? ` ${table.concealed} rows had hidden addresses and were skipped.` : ''}`,
        'success'
      );
      ctx.setGrowthBadge(r.summary.anomalies.length);
      onSynced();
    } catch (e) {
      ctx.log('Sync tenant dashboard from Microsoft 365', 'usage report', false, graphErrorText(e));
      ctx.notify(graphErrorText(e), 'error');
    } finally {
      setBusy(undefined);
    }
  };

  return (
    <DefaultButton iconProps={{ iconName: 'CloudDownload' }} onClick={sync} disabled={!!busy} title="Needs the Reports.Read.All Graph permission (see API access)">
      {busy || 'Sync from Microsoft 365'}
    </DefaultButton>
  );
};

/** 180-day tenant storage and active-site trends from the usage reports; loaded on request. */
export const M365Trends: React.FC = () => {
  const ctx = useAdmin();
  const [state, setState] = React.useState<{ loading: boolean; storage?: ITrendPoint[]; sites?: ITrendPoint[]; error?: string }>({ loading: false });

  const load = async (): Promise<void> => {
    setState({ loading: true });
    try {
      const storage = await ctx.graph.storageTrend();
      const sites = await ctx.graph.siteCountTrend();
      setState({ loading: false, storage, sites });
    } catch (e) {
      setState({ loading: false, error: graphErrorText(e) });
    }
  };

  const st = state.storage || [];
  const si = state.sites || [];
  const growth = st.length > 1 ? st[st.length - 1].value - st[0].value : undefined;
  return (
    <Card
      title="Microsoft 365 trends (last 180 days)"
      right={
        <DefaultButton iconProps={{ iconName: 'AreaChart' }} onClick={load} disabled={state.loading}>
          {state.loading ? 'Loading…' : state.storage ? 'Reload' : 'Load trends'}
        </DefaultButton>
      }
    >
      {state.error && <MessageBar messageBarType={MessageBarType.warning} isMultiline>{state.error}</MessageBar>}
      {!state.storage && !state.error && <p className={styles.muted} style={{ margin: 0 }}>Tenant storage and active sites per day, from Microsoft 365 usage reports (Reports.Read.All). Two requests.</p>}
      {state.storage && (
        <div className={styles.grid}>
          <div>
            <h4 style={{ margin: '0 0 6px' }}>
              Storage used: {st.length ? formatTB(st[st.length - 1].value) : '—'}{' '}
              {growth !== undefined && <span className={styles.muted} style={{ fontWeight: 400 }}>({growth >= 0 ? '+' : '−'}{formatTB(Math.abs(growth))} in {st.length} days)</span>}
            </h4>
            <LineChart points={st.map((p) => ({ date: p.date, a: p.value }))} format={formatTB} labelA="Storage used" />
          </div>
          <div>
            <h4 style={{ margin: '0 0 6px' }}>
              Sites: {si.length ? si[si.length - 1].value.toLocaleString() : '—'} total, {si.length && si[si.length - 1].value2 !== undefined ? (si[si.length - 1].value2 as number).toLocaleString() : '—'} active
            </h4>
            <LineChart points={si.map((p) => ({ date: p.date, a: p.value, b: p.value2 }))} format={(n) => Math.round(n).toLocaleString()} labelA="Total sites" labelB="Active sites" />
          </div>
        </div>
      )}
    </Card>
  );
};

const LineChart: React.FC<{ points: Array<{ date: string; a: number; b?: number }>; format: (n: number) => string; labelA: string; labelB?: string }> = ({ points, format, labelA, labelB }) => {
  if (points.length < 2) {
    return <p className={styles.muted}>Not enough data yet.</p>;
  }
  const W = 520;
  const H = 150;
  const PAD = 6;
  const values = points.map((p) => p.a).concat(points.filter((p) => p.b !== undefined).map((p) => p.b as number));
  let min = Math.min.apply(null, values);
  let max = Math.max.apply(null, values);
  if (max === min) {
    max += 1;
    min -= 1;
  }
  const x = (i: number): number => PAD + (i / (points.length - 1)) * (W - 2 * PAD);
  const y = (v: number): number => H - PAD - ((v - min) / (max - min)) * (H - 2 * PAD);
  const line = (get: (p: { a: number; b?: number }) => number | undefined): string =>
    points
      .map((p, i) => [i, get(p)] as [number, number | undefined])
      .filter((q) => q[1] !== undefined)
      .map((q) => `${x(q[0]).toFixed(1)},${y(q[1] as number).toFixed(1)}`)
      .join(' ');
  return (
    <figure style={{ margin: 0 }}>
      <svg viewBox={`0 0 ${W} ${H}`} width="100%" height={H} preserveAspectRatio="none" role="img" aria-label={`${labelA}${labelB ? ' and ' + labelB : ''} from ${points[0].date} to ${points[points.length - 1].date}`}>
        <line x1={PAD} x2={W - PAD} y1={H - PAD} y2={H - PAD} stroke="#d9dee7" />
        <polyline points={line((p) => p.a)} fill="none" stroke={SERIES[0]} strokeWidth={2} vectorEffect="non-scaling-stroke" />
        {labelB && <polyline points={line((p) => p.b)} fill="none" stroke={SERIES[1]} strokeWidth={2} vectorEffect="non-scaling-stroke" />}
      </svg>
      <figcaption className={styles.muted} style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, gap: 8, flexWrap: 'wrap' }}>
        <span>{points[0].date}</span>
        <span>
          <span className={styles.swatch} style={{ background: SERIES[0] }} /> {labelA} {format(points[points.length - 1].a)}
          {labelB && points[points.length - 1].b !== undefined && (
            <>
              {' · '}
              <span className={styles.swatch} style={{ background: SERIES[1] }} /> {labelB} {format(points[points.length - 1].b as number)}
            </>
          )}
          {' · '}range {format(min)}–{format(max)}
        </span>
        <span>{points[points.length - 1].date}</span>
      </figcaption>
    </figure>
  );
};
