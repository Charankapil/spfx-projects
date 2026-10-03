import { GrowthPoint, GrowthStatus, IGrowthSettings, analyzeSite } from './GrowthEngine';
import { ISnapshotData } from './StorageImport';

/**
 * Tenant-wide growth analysis over imported snapshots. It runs once, when an
 * admin imports a CSV (or presses Recalculate), and the small result is stored
 * with the index, so the Dashboard never has to load the snapshots themselves.
 */

export interface ISnapshotMeta {
  t: number;
  file: string;
  sites: number;
  bytes: number;
  source?: string;
}

export interface ISummaryRow {
  url: string;
  title: string;
  latestBytes: number;
  fraction: number;
  status: GrowthStatus;
  reasons: string[];
  growth7Bytes: number;
  growth7Pct: number;
  daysToFull?: number;
  spark: number[];
}

export interface ITenantSummary {
  at: number;
  sites: number;
  totalBytes: number;
  /** Net growth of all sites over a week, extrapolated from the available snapshots. */
  growth7Bytes: number;
  basedOnSnapshots: number;
  anomalies: ISummaryRow[];
  largest: ISummaryRow[];
  growers: ISummaryRow[];
}

export const MAX_ANOMALIES = 300;
export const TOP_N = 25;
export const MAX_SNAPSHOTS_KEPT = 60;
const DAY = 86400;

/** The latest snapshot plus those closest to 1, 2, 3, 5, 7, 10, 14, 21 and 30 days earlier (at most `max`). */
export function selectSnapshots(metas: ISnapshotMeta[], max = 8): ISnapshotMeta[] {
  const sorted = metas.slice().sort((a, b) => a.t - b.t);
  if (sorted.length <= 1) {
    return sorted;
  }
  const latest = sorted[sorted.length - 1];
  const chosen: { [t: number]: ISnapshotMeta } = { [latest.t]: latest };
  const rest = sorted.slice(0, -1);
  [7, 1, 14, 2, 3, 5, 10, 21, 30].forEach((d) => {
    if (Object.keys(chosen).length >= max || rest.length === 0) {
      return;
    }
    const target = latest.t - d * DAY;
    const best = rest.filter((m) => !chosen[m.t]).reduce<ISnapshotMeta | undefined>((b, m) => (!b || Math.abs(m.t - target) < Math.abs(b.t - target) ? m : b), undefined);
    if (best) {
      chosen[best.t] = best;
    }
  });
  return Object.keys(chosen)
    .map((k) => chosen[Number(k)])
    .sort((a, b) => a.t - b.t);
}

const rank: { [k in GrowthStatus]: number } = { baseline: 0, ok: 1, warning: 2, critical: 3 };

export function analyzeTenant(snaps: ISnapshotData[], settings: IGrowthSettings, titles: { [pathLower: string]: string }, origin: string, now: number): ITenantSummary {
  const sorted = snaps.slice().sort((a, b) => a.t - b.t);
  const latest = sorted[sorted.length - 1];
  const summary: ITenantSummary = { at: now, sites: 0, totalBytes: 0, growth7Bytes: 0, basedOnSnapshots: sorted.length, anomalies: [], largest: [], growers: [] };
  if (!latest) {
    return summary;
  }
  // Index every earlier snapshot by lower-cased path.
  const lookups = sorted.map((s) => {
    const m: { [k: string]: number } = {};
    for (let i = 0; i < s.u.length; i++) {
      m[s.u[i].toLowerCase()] = i;
    }
    return m;
  });
  const rows: ISummaryRow[] = [];
  for (let i = 0; i < latest.u.length; i++) {
    const path = latest.u[i];
    const key = path.toLowerCase();
    const points: GrowthPoint[] = [];
    for (let s = 0; s < sorted.length; s++) {
      const j = s === sorted.length - 1 ? i : lookups[s][key];
      if (j !== undefined) {
        points.push([sorted[s].t, sorted[s].b[j], sorted[s].q[j] || 0]);
      }
    }
    const url = path.charAt(0) === '/' ? (path === '/' ? origin : origin + path) : path;
    const title = titles[key] || path.substring(path.lastIndexOf('/') + 1) || url;
    const a = analyzeSite(url, title, points, settings);
    summary.sites++;
    summary.totalBytes += a.latestBytes;
    if (a.status !== 'baseline') {
      summary.growth7Bytes += a.growth7Bytes;
    }
    rows.push({ url, title, latestBytes: a.latestBytes, fraction: a.fraction, status: a.status, reasons: a.reasons, growth7Bytes: a.growth7Bytes, growth7Pct: a.growth7Pct, daysToFull: a.daysToFull, spark: a.spark });
  }
  summary.anomalies = rows
    .filter((r) => r.status === 'warning' || r.status === 'critical')
    .sort((a, b) => rank[b.status] - rank[a.status] || b.growth7Bytes - a.growth7Bytes)
    .slice(0, MAX_ANOMALIES);
  summary.largest = rows.slice().sort((a, b) => b.latestBytes - a.latestBytes).slice(0, TOP_N);
  summary.growers = rows
    .filter((r) => r.status !== 'baseline' && r.growth7Bytes > 0)
    .sort((a, b) => b.growth7Bytes - a.growth7Bytes)
    .slice(0, TOP_N);
  return summary;
}
