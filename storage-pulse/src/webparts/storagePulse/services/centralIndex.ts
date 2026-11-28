import { AGE_BUCKET_COUNT, IAgeHistogram, IScanResult } from '../models/IScanResult';
import { totalHistogram } from './activity';
import { safeHref } from './safeData';
import { coverage, estimatedCount, forDisplay } from './siteMap';

/**
 * One line per site collection scanned from the central page: enough for the
 * overview table (and its active / inactive split for any threshold) without
 * loading every site's full results.
 */
export interface ISiteSummary {
  url: string;
  title: string;
  scanStartedAt: string;
  scanCompletedAt: string;
  scannedBy: string;
  /** Site collection storage used, as SharePoint reports it. */
  siteStorageBytes?: number;
  /** Files shown on the site's dashboard (read file by file, or estimated from search), by month of age. */
  histogram: IAgeHistogram;
  versionBytes: number;
  libraries: number;
  librariesRead: number;
  librariesEstimated: number;
  /** Libraries and sites that could not be read completely. */
  failures: number;
}

export interface ICentralIndex {
  schemaVersion: 1;
  sites: ISiteSummary[];
}

const MAX_SITES = 5000;

export function emptyIndex(): ICentralIndex {
  return { schemaVersion: 1, sites: [] };
}

/** Site URLs compared without case or a trailing slash. */
export function siteKey(url: string): string {
  return url.replace(/\/+$/, '').toLowerCase();
}

export function summarize(result: IScanResult): ISiteSummary {
  const view = forDisplay(result);
  const cover = coverage(result);
  let versionBytes = 0;
  for (const lib of view.libraries) {
    if (lib.metrics) {
      versionBytes += Math.max(0, lib.metrics.totalSize - lib.metrics.fileStreamSize);
    }
  }
  return {
    url: result.rootUrl.replace(/\/+$/, ''),
    title: result.rootTitle,
    scanStartedAt: result.scanStartedAt,
    scanCompletedAt: result.scanCompletedAt,
    scannedBy: result.scannedBy,
    siteStorageBytes: result.siteStorageBytes,
    histogram: totalHistogram(view.libraries),
    versionBytes,
    libraries: cover.total,
    librariesRead: cover.scanned,
    librariesEstimated: estimatedCount(result),
    failures: result.libraries.filter((l) => l.error).length + result.webs.filter((w) => w.error).length
  };
}

/** Adds or replaces the site's line. */
export function upsert(index: ICentralIndex, summary: ISiteSummary): ICentralIndex {
  const key = siteKey(summary.url);
  const sites = index.sites.filter((s) => siteKey(s.url) !== key);
  sites.push(summary);
  sites.sort((a, b) => a.title.localeCompare(b.title));
  return { schemaVersion: 1, sites: sites.slice(0, MAX_SITES) };
}

export function removeSite(index: ICentralIndex, url: string): ICentralIndex {
  const key = siteKey(url);
  return { schemaVersion: 1, sites: index.sites.filter((s) => siteKey(s.url) !== key) };
}

function count(value: unknown): number {
  return typeof value === 'number' && isFinite(value) && value >= 0 ? value : 0;
}

function text(value: unknown, max = 500): string {
  return typeof value === 'string' ? value.substring(0, max) : '';
}

function isDate(value: unknown): value is string {
  return typeof value === 'string' && !isNaN(new Date(value).getTime());
}

/** The index file lives in Site Assets, which members can edit: everything read back is checked. */
export function sanitizeIndex(value: unknown, origin: string): ICentralIndex {
  const raw = value as ICentralIndex | undefined;
  const index = emptyIndex();
  if (!raw || typeof raw !== 'object' || !Array.isArray(raw.sites)) {
    return index;
  }
  for (const s of raw.sites.slice(0, MAX_SITES)) {
    const url = s ? safeHref(text(s.url), origin) : undefined;
    const h = s && s.histogram;
    if (!url || !isDate(s.scanStartedAt) || !isDate(s.scanCompletedAt) || !h) {
      continue;
    }
    if (!Array.isArray(h.counts) || !Array.isArray(h.bytes) || h.counts.length !== AGE_BUCKET_COUNT || h.bytes.length !== AGE_BUCKET_COUNT) {
      continue;
    }
    index.sites.push({
      url: url.replace(/\/+$/, ''),
      title: text(s.title) || url,
      scanStartedAt: s.scanStartedAt,
      scanCompletedAt: s.scanCompletedAt,
      scannedBy: text(s.scannedBy),
      siteStorageBytes: s.siteStorageBytes === undefined ? undefined : count(s.siteStorageBytes),
      histogram: { counts: h.counts.map(count), bytes: h.bytes.map(count) },
      versionBytes: count(s.versionBytes),
      libraries: Math.floor(count(s.libraries)),
      librariesRead: Math.floor(count(s.librariesRead)),
      librariesEstimated: Math.floor(count(s.librariesEstimated)),
      failures: Math.floor(count(s.failures))
    });
  }
  return index;
}
