import {
  AGE_BUCKET_COUNT,
  IAgeHistogram,
  IFileTypeStat,
  ILargeFile,
  ILibraryResult,
  IScanResult,
  IWebResult,
  MAX_AGE_MONTHS
} from '../models/IScanResult';
import { AGE_BANDS } from './activity';

/*
 * The saved scan lives in Site Assets, which site members can usually edit.
 * Everything read back from it is therefore treated as untrusted: shapes and
 * numbers are checked, sizes are capped, and URLs are only ever used as links
 * when they are https URLs on the tenant's own SharePoint host.
 */

const MAX_LIBRARIES = 20000;
const MAX_WEBS = 20000;
const MAX_FILE_TYPES = 400;
const MAX_LARGE_FILES = 200;
const MAX_TEXT = 1000;

/**
 * Percent-encodes each segment of a server-relative path, so names with
 * "#", "%" or "?" make working links (FileRef and ServerRelativeUrl come
 * back from SharePoint unencoded).
 */
export function encodePath(serverRelativePath: string): string {
  return serverRelativePath
    .split('/')
    .map((segment) => {
      let decoded = segment;
      try {
        decoded = decodeURIComponent(segment);
      } catch {
        // A lone "%" in a name: encode the raw segment.
      }
      return encodeURIComponent(decoded);
    })
    .join('/');
}

/** The URL if it is https on `origin`, otherwise undefined. Use for every href built from stored data. */
export function safeHref(url: string | undefined, origin: string): string | undefined {
  if (!url) {
    return undefined;
  }
  try {
    const parsed = new URL(url);
    return parsed.protocol === 'https:' && parsed.origin === origin ? parsed.href : undefined;
  } catch {
    return undefined;
  }
}

/** A link to a server-relative path on `origin`, or undefined if the path is not server-relative. */
export function serverRelativeHref(path: string | undefined, origin: string): string | undefined {
  if (!path || path.charAt(0) !== '/' || path.charAt(1) === '/') {
    return undefined;
  }
  return safeHref(`${origin}${encodePath(path)}`, origin);
}

function text(value: unknown, fallback = ''): string {
  return typeof value === 'string' ? value.substring(0, MAX_TEXT) : fallback;
}

function count(value: unknown): number {
  return typeof value === 'number' && isFinite(value) && value >= 0 ? value : 0;
}

function numbers(value: unknown, length: number): number[] | undefined {
  if (!Array.isArray(value) || value.length !== length) {
    return undefined;
  }
  return value.map(count);
}

function histogram(value: unknown): IAgeHistogram | undefined {
  const h = value as IAgeHistogram | undefined;
  if (!h) {
    return undefined;
  }
  const counts = numbers(h.counts, AGE_BUCKET_COUNT);
  const bytes = numbers(h.bytes, AGE_BUCKET_COUNT);
  return counts && bytes ? { counts, bytes } : undefined;
}

function fileTypes(value: unknown): IFileTypeStat[] | undefined {
  if (!Array.isArray(value)) {
    return undefined;
  }
  const out: IFileTypeStat[] = [];
  for (const raw of value.slice(0, MAX_FILE_TYPES)) {
    const counts = numbers(raw && raw.counts, AGE_BANDS.length);
    const bytes = numbers(raw && raw.bytes, AGE_BANDS.length);
    const extension = text(raw && raw.extension).substring(0, 20);
    if (counts && bytes && extension) {
      out.push({ extension, counts, bytes });
    }
  }
  return out;
}

function library(raw: ILibraryResult, origin: string): ILibraryResult | undefined {
  const h = raw && typeof raw === 'object' ? histogram(raw.histogram) : undefined;
  if (!h) {
    return undefined;
  }
  const lib: ILibraryResult = {
    id: text(raw.id),
    title: text(raw.title, '(untitled)'),
    webTitle: text(raw.webTitle),
    webUrl: safeHref(text(raw.webUrl), origin) || '',
    url: safeHref(text(raw.url), origin) || '',
    itemCount: count(raw.itemCount),
    files: count(raw.files),
    bytes: count(raw.bytes),
    histogram: h,
    fileTypes: fileTypes(raw.fileTypes)
  };
  if (raw.metrics) {
    lib.metrics = { totalSize: count(raw.metrics.totalSize), fileStreamSize: count(raw.metrics.fileStreamSize) };
  }
  if (raw.error) {
    lib.error = text(raw.error);
    lib.partial = raw.partial === true;
  }
  if (raw.unreadItems) {
    lib.unreadItems = count(raw.unreadItems);
  }
  return lib;
}

function largeFile(raw: ILargeFile): ILargeFile | undefined {
  if (!raw || typeof raw.serverRelativeUrl !== 'string' || raw.serverRelativeUrl.charAt(0) !== '/') {
    return undefined;
  }
  const modified = text(raw.modified);
  return {
    name: text(raw.name, '(unnamed)'),
    serverRelativeUrl: text(raw.serverRelativeUrl),
    bytes: count(raw.bytes),
    modified: isNaN(new Date(modified).getTime()) ? '' : modified,
    ageMonths: Math.min(MAX_AGE_MONTHS, Math.floor(count(raw.ageMonths))),
    libraryTitle: text(raw.libraryTitle),
    webTitle: text(raw.webTitle)
  };
}

function isDate(value: unknown): value is string {
  return typeof value === 'string' && !isNaN(new Date(value).getTime());
}

/**
 * Returns a clean copy of a saved scan, or undefined if it is not a scan this
 * web part could have written. `origin` is the tenant host the page runs on.
 */
export function sanitizeResult(value: unknown, origin: string): IScanResult | undefined {
  const raw = value as IScanResult | undefined;
  if (!raw || typeof raw !== 'object' || !Array.isArray(raw.libraries) || !Array.isArray(raw.webs)) {
    return undefined;
  }
  if (!isDate(raw.scanStartedAt) || !isDate(raw.scanCompletedAt)) {
    return undefined;
  }
  const rootUrl = safeHref(text(raw.rootUrl), origin);
  if (!rootUrl) {
    return undefined;
  }
  const libraries: ILibraryResult[] = [];
  for (const lib of raw.libraries.slice(0, MAX_LIBRARIES)) {
    const clean = library(lib, origin);
    if (clean) {
      libraries.push(clean);
    }
  }
  const webs: IWebResult[] = raw.webs.slice(0, MAX_WEBS).map((w: IWebResult) => ({
    title: text(w && w.title),
    url: safeHref(text(w && w.url), origin) || '',
    error: w && w.error ? text(w.error) : undefined
  }));
  const largest: ILargeFile[] = [];
  for (const f of Array.isArray(raw.largestOldFiles) ? raw.largestOldFiles.slice(0, MAX_LARGE_FILES) : []) {
    const clean = largeFile(f);
    if (clean) {
      largest.push(clean);
    }
  }
  const options = raw.options;
  return {
    scope: raw.scope === 'currentWeb' ? 'currentWeb' : 'siteCollection',
    rootUrl,
    rootTitle: text(raw.rootTitle),
    siteStorageBytes: raw.siteStorageBytes === undefined ? undefined : count(raw.siteStorageBytes),
    webs,
    libraries,
    largestOldFiles: largest,
    scanStartedAt: raw.scanStartedAt,
    scanCompletedAt: raw.scanCompletedAt,
    scannedBy: text(raw.scannedBy),
    options: options
      ? {
          includeHidden: options.includeHidden === true,
          excludeSystemLibraries: options.excludeSystemLibraries === true,
          excludedLibraries: Array.isArray(options.excludedLibraries) ? options.excludedLibraries.map((n) => text(n)).slice(0, 100) : []
        }
      : undefined
  };
}
