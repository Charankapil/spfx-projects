/**
 * Turns the CSV produced by a flow (or a Get-SPOSite export) into a compact
 * storage snapshot. Column names are detected automatically but can always be
 * mapped by hand; units are guessed from the header or the size of the values
 * and confirmed by the admin in a preview. Pure functions, no network.
 */

export type Unit = 'bytes' | 'kb' | 'mb' | 'gb' | 'tb';

export const UNIT_BYTES: { [u in Unit]: number } = { bytes: 1, kb: 1024, mb: 1048576, gb: 1073741824, tb: 1099511627776 };

export interface IMapping {
  url: string;
  storage: string;
  storageUnit: Unit;
  quota?: string;
  quotaUnit: Unit;
  title?: string;
  deleted?: string;
  excludeOneDrive: boolean;
}

/** One snapshot: parallel arrays so 17,000 sites stay around half a megabyte of JSON. */
export interface ISnapshotData {
  /** Epoch seconds. */
  t: number;
  /** Site addresses relative to the tenant origin (e.g. /sites/finance); other origins stay absolute. */
  u: string[];
  /** Storage used, bytes. */
  b: number[];
  /** Per-mille of quota used (0 = unknown). */
  q: number[];
}

const norm = (s: string): string => s.toLowerCase().replace(/[^a-z0-9]/g, '');

const CANDIDATES = {
  url: ['siteurl', 'url', 'siteaddress', 'weburl', 'sitecollectionurl', 'path', 'site'],
  storage: ['storageusagecurrent', 'storageused', 'storageusedmb', 'storageusedgb', 'storageusedbytes', 'storageusage', 'storageinuse', 'currentstorage', 'sizemb', 'sizegb', 'sizebytes', 'storage', 'usage', 'size'],
  quota: ['storagequota', 'storagequotamb', 'storagequotagb', 'quota', 'quotamb', 'storagemaximumlevel', 'storagelimit', 'maxstorage', 'allocated', 'storageallocated'],
  title: ['title', 'sitename', 'name', 'displayname'],
  deleted: ['timedeleted', 'isdeleted', 'deleted', 'deleteddate', 'deletiontime']
};

function pick(headers: string[], names: string[]): string | undefined {
  const normalized = headers.map(norm);
  for (const n of names) {
    const i = normalized.indexOf(n);
    if (i >= 0) {
      return headers[i];
    }
  }
  return undefined;
}

export function unitFromHeader(header: string | undefined): Unit | undefined {
  const h = (header || '').toLowerCase();
  if (/\b(tb|terabytes?)\b|\(tb\)|tb$/.test(h) || /\btb\b/.test(h)) {
    return 'tb';
  }
  if (/gb$|\bgb\b|gigabytes?|\(gb\)/.test(h)) {
    return 'gb';
  }
  if (/mb$|\bmb\b|megabytes?|\(mb\)/.test(h)) {
    return 'mb';
  }
  if (/kb$|\bkb\b|kilobytes?|\(kb\)/.test(h)) {
    return 'kb';
  }
  if (/bytes?\b|\(b\)/.test(h)) {
    return 'bytes';
  }
  return undefined;
}

export function parseNumber(raw: string | undefined): number {
  if (raw === undefined || raw === null) {
    return NaN;
  }
  let t = String(raw).trim().replace(/[^\d.,\-+eE]/g, '');
  if (!t || !/\d/.test(t)) {
    return NaN;
  }
  const hasDot = t.indexOf('.') >= 0;
  const hasComma = t.indexOf(',') >= 0;
  if (hasDot && hasComma) {
    // The separator that comes last is the decimal mark; the other groups thousands.
    if (t.lastIndexOf(',') > t.lastIndexOf('.')) {
      t = t.replace(/\./g, '').replace(',', '.');
    } else {
      t = t.replace(/,/g, '');
    }
  } else if (hasComma) {
    t = /^[-+]?\d{1,3}(,\d{3})+$/.test(t) ? t.replace(/,/g, '') : t.replace(',', '.');
  }
  const n = Number(t);
  return isFinite(n) ? n : NaN;
}

function median(values: number[]): number {
  const v = values.filter((x) => isFinite(x)).sort((a, b) => a - b);
  return v.length ? v[Math.floor(v.length / 2)] : 0;
}

/** Guess a unit from the magnitude of the values when the header does not say. */
export function guessUnitFromValues(values: number[], kind: 'storage' | 'quota'): Unit {
  const m = median(values.filter((x) => x > 0));
  if (kind === 'quota') {
    return m > 1e9 ? 'bytes' : 'mb';
  }
  return m > 5e6 ? 'bytes' : 'mb';
}

export function detectMapping(headers: string[], rows: string[][]): IMapping | undefined {
  const url = pick(headers, CANDIDATES.url);
  const storage = pick(headers, CANDIDATES.storage);
  if (!url || !storage) {
    return undefined;
  }
  const quota = pick(headers, CANDIDATES.quota);
  const col = (h: string): number[] => rows.slice(0, 500).map((r) => parseNumber(r[headers.indexOf(h)]));
  return {
    url,
    storage,
    storageUnit: unitFromHeader(storage) || guessUnitFromValues(col(storage), 'storage'),
    quota,
    quotaUnit: quota ? unitFromHeader(quota) || guessUnitFromValues(col(quota), 'quota') : 'mb',
    title: pick(headers, CANDIDATES.title),
    deleted: pick(headers, CANDIDATES.deleted),
    excludeOneDrive: true
  };
}

export function toPath(url: string, origin: string): string {
  const u = url.trim().replace(/\/+$/, '');
  return u.toLowerCase().indexOf(origin.toLowerCase()) === 0 ? u.substring(origin.length) || '/' : u;
}

export function fromPath(path: string, origin: string): string {
  return path.charAt(0) === '/' ? (path === '/' ? origin : origin + path) : path;
}

export interface IBuildStats {
  rows: number;
  kept: number;
  noUrl: number;
  badNumber: number;
  deleted: number;
  oneDrive: number;
  duplicates: number;
}

export interface IBuildResult {
  data: ISnapshotData;
  /** Titles by lower-cased path, when the CSV has a title column. */
  titles: { [pathLower: string]: string };
  stats: IBuildStats;
}

export function buildSnapshot(rows: string[][], mapping: IMapping, origin: string, t: number): IBuildResult {
  const headers = rows[0] || [];
  const idx = (h: string | undefined): number => (h ? headers.indexOf(h) : -1);
  const iUrl = idx(mapping.url);
  const iStorage = idx(mapping.storage);
  const iQuota = idx(mapping.quota);
  const iTitle = idx(mapping.title);
  const iDeleted = idx(mapping.deleted);
  const stats: IBuildStats = { rows: Math.max(0, rows.length - 1), kept: 0, noUrl: 0, badNumber: 0, deleted: 0, oneDrive: 0, duplicates: 0 };
  const seen: { [k: string]: number } = {};
  const data: ISnapshotData = { t, u: [], b: [], q: [] };
  const titles: { [k: string]: string } = {};
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const rawUrl = (row[iUrl] || '').trim();
    if (!/^https?:\/\//i.test(rawUrl) && rawUrl.charAt(0) !== '/') {
      stats.noUrl++;
      continue;
    }
    if (iDeleted >= 0) {
      const d = (row[iDeleted] || '').trim().toLowerCase();
      if (d && d !== 'false' && d !== '0' && d !== 'no' && d !== 'null') {
        stats.deleted++;
        continue;
      }
    }
    if (mapping.excludeOneDrive && /-my\.sharepoint\.com|\/personal\//i.test(rawUrl)) {
      stats.oneDrive++;
      continue;
    }
    const used = parseNumber(row[iStorage]);
    if (!isFinite(used) || used < 0) {
      stats.badNumber++;
      continue;
    }
    const bytes = Math.round(used * UNIT_BYTES[mapping.storageUnit]);
    let pm = 0;
    if (iQuota >= 0) {
      const quota = parseNumber(row[iQuota]) * UNIT_BYTES[mapping.quotaUnit];
      if (isFinite(quota) && quota > 0) {
        pm = Math.min(2000, Math.round((bytes / quota) * 1000));
      }
    }
    const path = toPath(rawUrl.charAt(0) === '/' ? origin + rawUrl : rawUrl, origin);
    const key = path.toLowerCase();
    if (seen[key] !== undefined) {
      stats.duplicates++;
      data.b[seen[key]] = bytes;
      data.q[seen[key]] = pm;
      continue;
    }
    seen[key] = data.u.length;
    data.u.push(path);
    data.b.push(bytes);
    data.q.push(pm);
    if (iTitle >= 0 && row[iTitle]) {
      titles[key] = row[iTitle].trim();
    }
    stats.kept++;
  }
  return { data, titles, stats };
}
