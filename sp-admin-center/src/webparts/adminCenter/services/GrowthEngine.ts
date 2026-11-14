import { formatBytes } from './format';

/**
 * Storage-growth history and anomaly rules. Pure functions only (no network),
 * so every rule is unit-tested with synthetic histories.
 */

/** [epoch seconds, bytes used, per-mille of quota used (0 = unknown)] */
export type GrowthPoint = [number, number, number];

export interface IGrowthSettings {
  /** Flag when projected 7-day growth is at least this many percent... */
  pctPerWeek: number;
  /** ...and at least this many GB. */
  gbPerWeek: number;
  /** Always flag projected 7-day growth of at least this many GB. */
  hugeGbPerWeek: number;
  /** Flag when the quota is projected to fill within this many days. */
  forecastDays: number;
  /** The latest growth rate is a spike when it exceeds the usual rate by this many standard deviations. */
  sigma: number;
}

export const DEFAULT_SETTINGS: IGrowthSettings = { pctPerWeek: 25, gbPerWeek: 5, hugeGbPerWeek: 50, forecastDays: 30, sigma: 3 };

export interface IGrowthSite {
  title: string;
  s: GrowthPoint[];
}

export interface IGrowthDoc {
  v: 1;
  settings: IGrowthSettings;
  sites: { [url: string]: IGrowthSite };
  /** Epoch seconds of the most recent capture. */
  lastCapture?: number;
}

export const MAX_POINTS = 90;
export const MIN_GAP_SEC = 6 * 3600;
export const MAX_TRACKED = 500;
const GB = 1073741824;
const DAY = 86400;

export function emptyDoc(): IGrowthDoc {
  return { v: 1, settings: { ...DEFAULT_SETTINGS }, sites: {} };
}

/** Accepts whatever is in the file and returns a safe document (drops anything malformed). */
export function normalizeDoc(raw: unknown): IGrowthDoc {
  const doc = emptyDoc();
  const r = raw as { settings?: { [k: string]: unknown }; sites?: { [k: string]: { title?: unknown; s?: unknown } }; lastCapture?: unknown } | null;
  if (!r || typeof r !== 'object') {
    return doc;
  }
  if (r.settings && typeof r.settings === 'object') {
    (Object.keys(DEFAULT_SETTINGS) as Array<keyof IGrowthSettings>).forEach((k) => {
      const v = Number(r.settings && r.settings[k]);
      if (isFinite(v) && v > 0) {
        doc.settings[k] = v;
      }
    });
  }
  if (r.sites && typeof r.sites === 'object') {
    Object.keys(r.sites).forEach((url) => {
      const site = r.sites && r.sites[url];
      if (!site || !Array.isArray(site.s)) {
        return;
      }
      const pts = (site.s as unknown[])
        .filter((p) => Array.isArray(p) && isFinite(Number(p[0])) && isFinite(Number(p[1])) && Number(p[1]) >= 0)
        .map((p) => [Number((p as number[])[0]), Number((p as number[])[1]), Number((p as number[])[2]) || 0] as GrowthPoint)
        .sort((a, b) => a[0] - b[0])
        .slice(-MAX_POINTS);
      doc.sites[url] = { title: typeof site.title === 'string' ? site.title : url, s: pts };
    });
  }
  if (isFinite(Number(r.lastCapture))) {
    doc.lastCapture = Number(r.lastCapture);
  }
  return doc;
}

/** Appends a point; a point within MIN_GAP_SEC of the last one replaces it (keeps the file small). */
export function addPoint(points: GrowthPoint[], p: GrowthPoint): GrowthPoint[] {
  const out = points.slice();
  const last = out[out.length - 1];
  if (last && p[0] - last[0] < MIN_GAP_SEC && p[0] >= last[0]) {
    out[out.length - 1] = p;
  } else {
    out.push(p);
  }
  out.sort((a, b) => a[0] - b[0]);
  return out.slice(-MAX_POINTS);
}

export type GrowthStatus = 'baseline' | 'ok' | 'warning' | 'critical';

export interface ISiteGrowth {
  url: string;
  title: string;
  points: number;
  latestBytes: number;
  /** 0..1 of quota used, 0 if unknown. */
  fraction: number;
  status: GrowthStatus;
  reasons: string[];
  /** Observed growth between the baseline snapshot and the latest one. */
  deltaBytes: number;
  spanDays: number;
  /** Growth rate scaled to 7 days. */
  growth7Bytes: number;
  growth7Pct: number;
  daysToFull?: number;
  spark: number[];
}

const rank: { [k in GrowthStatus]: number } = { baseline: 0, ok: 1, warning: 2, critical: 3 };

export function analyzeSite(url: string, title: string, points: GrowthPoint[], settings: IGrowthSettings): ISiteGrowth {
  const pts = points.slice().sort((a, b) => a[0] - b[0]);
  const latest = pts[pts.length - 1];
  const result: ISiteGrowth = {
    url,
    title,
    points: pts.length,
    latestBytes: latest ? latest[1] : 0,
    fraction: latest && latest[2] > 0 ? latest[2] / 1000 : 0,
    status: 'baseline',
    reasons: [],
    deltaBytes: 0,
    spanDays: 0,
    growth7Bytes: 0,
    growth7Pct: 0,
    spark: pts.map((p) => p[1])
  };
  if (!latest) {
    return result;
  }

  // Baseline: the snapshot closest to 7 days before the latest, at least ~20 h older.
  const target = latest[0] - 7 * DAY;
  const candidates = pts.slice(0, -1).filter((p) => latest[0] - p[0] >= 20 * 3600);
  if (candidates.length === 0) {
    return result;
  }
  const base = candidates.reduce((best, p) => (Math.abs(p[0] - target) < Math.abs(best[0] - target) ? p : best));
  const days = (latest[0] - base[0]) / DAY;
  const delta = latest[1] - base[1];
  const perDay = delta / days;
  const g7 = perDay * 7;
  result.status = 'ok';
  result.deltaBytes = delta;
  result.spanDays = days;
  result.growth7Bytes = g7;
  result.growth7Pct = base[1] > 0 ? (g7 / base[1]) * 100 : g7 > 0 ? 100 : 0;

  const raise = (s: GrowthStatus, reason: string): void => {
    result.reasons.push(reason);
    if (rank[s] > rank[result.status]) {
      result.status = s;
    }
  };
  const span = days >= 1 ? `${days.toFixed(days < 10 ? 1 : 0)} days` : 'a day';

  if (delta > 0 && g7 >= settings.hugeGbPerWeek * GB) {
    raise('critical', `Grew ${formatBytes(delta)} in ${span} (about ${formatBytes(g7)} a week).`);
  } else if (delta > 0 && g7 >= settings.gbPerWeek * GB && result.growth7Pct >= settings.pctPerWeek) {
    const strong = result.growth7Pct >= settings.pctPerWeek * 2 && g7 >= settings.gbPerWeek * GB * 2;
    raise(strong ? 'critical' : 'warning', `Grew ${formatBytes(delta)} (+${Math.round((delta / Math.max(1, base[1])) * 100)}%) in ${span}.`);
  }

  // Spike: the latest interval is far above the site's usual daily growth.
  const rates: number[] = [];
  for (let i = 1; i < pts.length; i++) {
    const dt = (pts[i][0] - pts[i - 1][0]) / DAY;
    if (dt >= 0.25) {
      rates.push((pts[i][1] - pts[i - 1][1]) / dt);
    }
  }
  if (rates.length >= 5) {
    const last = rates[rates.length - 1];
    const prev = rates.slice(0, -1);
    const mean = prev.reduce((s, x) => s + x, 0) / prev.length;
    const sd = Math.sqrt(prev.reduce((s, x) => s + (x - mean) * (x - mean), 0) / prev.length);
    if (last > 0 && last * 7 >= settings.gbPerWeek * GB && last > mean + settings.sigma * Math.max(sd, 1) && last > 2 * Math.max(mean, 0)) {
      raise('warning', `Latest growth is ${formatBytes(last)} a day, far above its usual ${formatBytes(Math.max(mean, 0))} a day.`);
    }
  }

  // Forecast: when would the quota fill at the current rate?
  if (result.fraction > 0 && perDay > 0) {
    const quota = latest[1] / result.fraction;
    const dtf = Math.max(0, quota - latest[1]) / perDay;
    result.daysToFull = dtf;
    if (dtf <= 14) {
      raise('critical', `At this rate the storage quota is full in about ${Math.ceil(dtf)} days.`);
    } else if (dtf <= settings.forecastDays) {
      raise('warning', `At this rate the storage quota is full in about ${Math.ceil(dtf)} days.`);
    }
  }
  return result;
}

export function analyzeAll(doc: IGrowthDoc): ISiteGrowth[] {
  return Object.keys(doc.sites)
    .map((url) => analyzeSite(url, doc.sites[url].title, doc.sites[url].s, doc.settings))
    .sort((a, b) => rank[b.status] - rank[a.status] || b.growth7Bytes - a.growth7Bytes);
}

export function anomaliesOf(all: ISiteGrowth[]): ISiteGrowth[] {
  return all.filter((s) => s.status === 'warning' || s.status === 'critical');
}
