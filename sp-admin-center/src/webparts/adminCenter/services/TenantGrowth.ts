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
  /** 0 active, 1 archived, 2 deleted (from the latest snapshot). */
  state?: number;
}

export interface ISegment {
  sites: number;
  bytes: number;
  growth7Bytes: number;
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
  /** Totals per site state; absent in summaries saved before v1.3. */
  segments?: { active: ISegment; archived: ISegment; deleted: ISegment };
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
  const seg = (): ISegment => ({ sites: 0, bytes: 0, growth7Bytes: 0 });
  const segments = { active: seg(), archived: seg(), deleted: seg() };
  const summary: ITenantSummary = { at: now, sites: 0, totalBytes: 0, growth7Bytes: 0, basedOnSnapshots: sorted.length, anomalies: [], largest: [], growers: [], segments };
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
    const state = latest.s && latest.s[i] ? latest.s[i] : 0;
    const a = analyzeSite(url, title, points, settings);
    const bucket = state === 2 ? segments.deleted : state === 1 ? segments.archived : segments.active;
    bucket.sites++;
    bucket.bytes += a.latestBytes;
    summary.sites++;
    summary.totalBytes += a.latestBytes;
    if (a.status !== 'baseline') {
      bucket.growth7Bytes += a.growth7Bytes;
      summary.growth7Bytes += a.growth7Bytes;
    }
    if (state === 2) {
      // Deleted sites are counted in their group but never raise growth alerts.
      continue;
    }
    rows.push({ url, title, latestBytes: a.latestBytes, fraction: a.fraction, status: a.status, reasons: a.reasons, growth7Bytes: a.growth7Bytes, growth7Pct: a.growth7Pct, daysToFull: a.daysToFull, spark: a.spark, state });
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

// ---- tenant inventory overview (the "tenant dashboard") -----------------------------

export interface IBucket {
  label: string;
  count: number;
  bytes: number;
}

export interface ITopSite {
  url: string;
  title: string;
  bytes: number;
  /** Per-mille of the site's own quota (0 unknown). */
  pm: number;
  state: number;
  type: string;
}

export interface ITenantOverview {
  at: number;
  /** Snapshot this overview describes. */
  snapshotT: number;
  sites: number;
  activeSites: number;
  archivedSites: number;
  deletedSites: number;
  activeBytes: number;
  archivedBytes: number;
  deletedBytes: number;
  /** Sites connected to Teams; undefined when the CSV has no Teams column. */
  teamsSites?: number;
  /** Sites on a group-connected template; undefined when the CSV has no template column. */
  groupSites?: number;
  byType?: IBucket[];
  sizeBands: IBucket[];
  activity?: IBucket[];
  top: ITopSite[];
}

const GB = 1073741824;
export const SIZE_BANDS: Array<{ label: string; max: number }> = [
  { label: 'Under 1 GB', max: GB },
  { label: '1 – 10 GB', max: 10 * GB },
  { label: '10 – 100 GB', max: 100 * GB },
  { label: '100 GB – 1 TB', max: 1024 * GB },
  { label: '1 TB and more', max: Infinity }
];
export const ACTIVITY_BANDS: Array<{ label: string; maxDays: number }> = [
  { label: 'Last 30 days', maxDays: 30 },
  { label: '1 – 3 months', maxDays: 90 },
  { label: '3 – 6 months', maxDays: 180 },
  { label: '6 – 12 months', maxDays: 365 },
  { label: 'Over a year', maxDays: Infinity }
];
export const TOP_SITES = 50;

/**
 * Counts and breakdowns over every site in one snapshot. Deleted sites are counted in
 * their own figures but left out of the type / size / activity breakdowns and the top list.
 */
export function summarizeInventory(
  snap: ISnapshotData,
  titles: { [pathLower: string]: string },
  origin: string,
  now: number,
  typeOf: (template: string) => string
): ITenantOverview {
  const o: ITenantOverview = {
    at: now,
    snapshotT: snap.t,
    sites: 0,
    activeSites: 0,
    archivedSites: 0,
    deletedSites: 0,
    activeBytes: 0,
    archivedBytes: 0,
    deletedBytes: 0,
    sizeBands: SIZE_BANDS.map((b) => ({ label: b.label, count: 0, bytes: 0 })),
    top: []
  };
  const hasTpl = !!(snap.tp && snap.tpl);
  const hasTeams = !!snap.tm;
  const hasActivity = !!snap.la;
  const types: { [label: string]: IBucket } = {};
  const activity = hasActivity ? ACTIVITY_BANDS.map((b) => ({ label: b.label, count: 0, bytes: 0 })).concat([{ label: 'Unknown', count: 0, bytes: 0 }]) : undefined;
  if (hasTeams) {
    o.teamsSites = 0;
  }
  if (hasTpl) {
    o.groupSites = 0;
  }
  const today = Math.floor((now * 1000) / 86400000);
  const candidates: ITopSite[] = [];
  for (let i = 0; i < snap.u.length; i++) {
    const state = snap.s && snap.s[i] ? snap.s[i] : 0;
    const bytes = snap.b[i] || 0;
    o.sites++;
    if (state === 2) {
      o.deletedSites++;
      o.deletedBytes += bytes;
      continue;
    }
    if (state === 1) {
      o.archivedSites++;
      o.archivedBytes += bytes;
    } else {
      o.activeSites++;
      o.activeBytes += bytes;
    }
    const template = hasTpl ? (snap.tpl as string[])[(snap.tp as number[])[i]] || '' : '';
    const type = hasTpl ? typeOf(template) : '';
    if (hasTpl) {
      const t = (types[type] = types[type] || { label: type, count: 0, bytes: 0 });
      t.count++;
      t.bytes += bytes;
      if (template.indexOf('GROUP') === 0 || template.indexOf('TEAMCHANNEL') === 0) {
        o.groupSites = (o.groupSites || 0) + 1;
      }
    }
    if (hasTeams && (snap.tm as number[])[i] === 1) {
      o.teamsSites = (o.teamsSites || 0) + 1;
    }
    for (let k = 0; k < SIZE_BANDS.length; k++) {
      if (bytes < SIZE_BANDS[k].max) {
        o.sizeBands[k].count++;
        o.sizeBands[k].bytes += bytes;
        break;
      }
    }
    if (activity) {
      const day = (snap.la as number[])[i] || 0;
      const age = day ? today - day : -1;
      const k = age < 0 ? ACTIVITY_BANDS.length : ACTIVITY_BANDS.findIndex((b) => age < b.maxDays);
      activity[k].count++;
      activity[k].bytes += bytes;
    }
    const path = snap.u[i];
    candidates.push({
      url: path.charAt(0) === '/' ? (path === '/' ? origin : origin + path) : path,
      title: titles[path.toLowerCase()] || path.substring(path.lastIndexOf('/') + 1) || path,
      bytes,
      pm: snap.q[i] || 0,
      state,
      type
    });
  }
  o.top = candidates.sort((a, b) => b.bytes - a.bytes).slice(0, TOP_SITES);
  if (hasTpl) {
    o.byType = Object.keys(types)
      .map((k) => types[k])
      .sort((a, b) => b.count - a.count);
  }
  o.activity = activity;
  return o;
}
