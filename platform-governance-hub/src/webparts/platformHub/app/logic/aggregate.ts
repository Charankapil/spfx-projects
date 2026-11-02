const DAY = 86400000;

export interface ICount {
  label: string;
  value: number;
}

/** Counts items by key, sorted descending. */
export function countBy<T>(items: T[], key: (item: T) => string | undefined): ICount[] {
  const map: Record<string, number> = {};
  items.forEach((i) => {
    const k = key(i) || 'Unknown';
    map[k] = (map[k] || 0) + 1;
  });
  return Object.keys(map)
    .map((label) => ({ label, value: map[label] }))
    .sort((a, b) => b.value - a.value || (a.label < b.label ? -1 : 1));
}

export function sumBy<T>(items: T[], value: (item: T) => number): number {
  let s = 0;
  items.forEach((i) => {
    s += value(i) || 0;
  });
  return s;
}

/** Keeps the top `n` entries and folds the rest into "Other". */
export function topN(counts: ICount[], n: number, otherLabel: string = 'Other'): ICount[] {
  if (counts.length <= n) return counts;
  const head = counts.slice(0, n - 1);
  const rest = counts.slice(n - 1).reduce((s, c) => s + c.value, 0);
  return head.concat([{ label: otherLabel, value: rest }]);
}

/** Monday 00:00 (local) of the week containing `t`. */
export function weekStart(t: number): number {
  const d = new Date(t);
  d.setHours(0, 0, 0, 0);
  const day = (d.getDay() + 6) % 7; // Monday = 0
  return d.getTime() - day * DAY;
}

export interface IBucketSeries {
  /** ISO dates of bucket starts, oldest first. */
  buckets: string[];
  /** One array per series, aligned with `buckets`. */
  values: number[][];
}

/**
 * Buckets creation dates into the last `weeks` weeks for several series.
 * Items outside the window are ignored.
 */
export function weeklyBuckets(series: (string | undefined)[][], weeks: number, now: number = Date.now()): IBucketSeries {
  const end = weekStart(now);
  const starts: number[] = [];
  for (let i = weeks - 1; i >= 0; i--) starts.push(end - i * 7 * DAY);
  const values = series.map((dates) => {
    const row = starts.map(() => 0);
    dates.forEach((iso) => {
      if (!iso) return;
      const t = Date.parse(iso);
      if (isNaN(t) || t < starts[0]) return;
      const idx = Math.min(starts.length - 1, Math.floor((weekStart(t) - starts[0]) / (7 * DAY)));
      if (idx >= 0) row[idx]++;
    });
    return row;
  });
  return { buckets: starts.map((s) => new Date(s).toISOString()), values };
}

/** Monthly buckets ("2026-03") for the last `months` months. */
export function monthlyBuckets(series: (string | undefined)[][], months: number, now: number = Date.now()): IBucketSeries {
  const d = new Date(now);
  const keys: string[] = [];
  for (let i = months - 1; i >= 0; i--) {
    const m = new Date(d.getFullYear(), d.getMonth() - i, 1);
    keys.push(`${m.getFullYear()}-${String(m.getMonth() + 1).padStart(2, '0')}`);
  }
  const values = series.map((dates) => {
    const row = keys.map(() => 0);
    dates.forEach((iso) => {
      if (!iso) return;
      const t = Date.parse(iso);
      if (isNaN(t)) return;
      const m = new Date(t);
      const k = `${m.getFullYear()}-${String(m.getMonth() + 1).padStart(2, '0')}`;
      const idx = keys.indexOf(k);
      if (idx >= 0) row[idx]++;
    });
    return row;
  });
  return { buckets: keys.map((k) => `${k}-01T00:00:00.000Z`), values };
}

export function uniqueStrings(values: (string | undefined)[]): string[] {
  const seen: Record<string, boolean> = {};
  const out: string[] = [];
  values.forEach((v) => {
    if (v && !seen[v]) {
      seen[v] = true;
      out.push(v);
    }
  });
  return out;
}
