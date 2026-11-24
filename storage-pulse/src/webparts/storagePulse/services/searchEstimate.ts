import { IFileTypeStat, ILargeFile, ILibraryResult, ISearchEstimate } from '../models/IScanResult';
import { AGE_BANDS, ageInMonths, emptyFileTypeStat, emptyHistogram, MIN_THRESHOLD_MONTHS } from './activity';
import { ScanCancelledError, ScanPausedError } from './errors';
import { GetJson } from './folderWalk';

/** Size ranges asked of search, in bytes: half-decade steps from 1 KB to 100 GB. */
export const SIZE_BOUNDARIES: number[] = [];
for (let e = 3; e <= 11; e += 0.5) {
  SIZE_BOUNDARIES.push(Math.round(Math.pow(10, e)));
}
/** Types whose average size is asked for separately; the rest use their age band's average. */
const TYPES_SIZED = 10;
const MAX_TYPES = 300;
const LARGEST_KEPT = 200;

interface ICell {
  Key: string;
  Value: string | null;
}
interface IRefinerEntry {
  RefinementCount?: number | string;
  RefinementName?: string;
  RefinementToken?: string;
  RefinementValue?: string;
}
interface IRefiner {
  Name: string;
  Entries?: IRefinerEntry[] | { results?: IRefinerEntry[] };
}
interface ISearchResponse {
  PrimaryQueryResult?: {
    RelevantResults?: {
      TotalRows?: number | string;
      Table?: { Rows?: { Cells?: ICell[] | { results?: ICell[] } }[] | { results?: { Cells?: ICell[] | { results?: ICell[] } }[] } };
    };
    RefinementResults?: { Refiners?: IRefiner[] | { results?: IRefiner[] } } | null;
  };
}

interface IQueryResult {
  total: number;
  types: Map<string, number>;
  sizes: { from: number; to: number; count: number }[];
  rows: Record<string, string>[];
}

function list<T>(value: T[] | { results?: T[] } | undefined | null): T[] {
  if (!value) {
    return [];
  }
  return Array.isArray(value) ? value : value.results || [];
}

function toNumber(value: number | string | undefined | null): number {
  const parsed = typeof value === 'number' ? value : parseFloat(String(value === undefined || value === null ? '' : value));
  return isNaN(parsed) ? 0 : parsed;
}

/** "YYYY-MM-DDTHH:MM:SSZ", the form KQL accepts in property restrictions. */
function kqlDate(d: Date): string {
  return d.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/**
 * The first moment of the day after the date `months` before now: a file
 * modified before it is at least `months` old, counted the way the scan
 * counts age (whole calendar months, by day).
 */
function ageBoundary(now: Date, months: number): Date {
  const d = new Date(now.getTime());
  d.setMonth(d.getMonth() - months);
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 1);
  return d;
}

/** Parses a discretized refiner token such as "range(1024, 3162)" or "range(min, 1000)". */
export function parseRange(entry: IRefinerEntry): { from: number; to: number } | undefined {
  for (const text of [entry.RefinementToken, entry.RefinementValue, entry.RefinementName]) {
    const m = /range\(\s*([^,\s)]+)\s*,\s*([^,\s)]+)\s*\)/i.exec(text || '');
    if (m) {
      const from = m[1].toLowerCase() === 'min' ? 0 : parseFloat(m[1]);
      const to = m[2].toLowerCase() === 'max' ? Infinity : parseFloat(m[2]);
      if (!isNaN(from) && !isNaN(to)) {
        return { from, to };
      }
    }
  }
  return undefined;
}

/** A representative file size for a size range: the geometric middle, with sensible ends. */
function typicalSize(from: number, to: number, largeMean: number | undefined): number {
  if (to === Infinity) {
    return largeMean !== undefined ? largeMean : from * 1.5;
  }
  if (from <= 0) {
    return Math.min(400, to / 2);
  }
  return Math.sqrt(from * to);
}

export function normaliseType(value: string | undefined): string {
  const ext = (value || '').trim().toLowerCase().replace(/^\./, '');
  return !ext || ext.length > 10 || /\s/.test(ext) ? '(none)' : ext;
}

/**
 * Instant figures for one library from SharePoint search, in about twenty
 * requests whatever the library's size: search counts files per age band and
 * per type, and the largest old files come back with their exact sizes. Bytes
 * per band are estimated from search's size ranges, then scaled so the
 * average file size matches SharePoint's storage metrics. Search only sees
 * current versions of files that are in the index, and only those the person
 * may open.
 */
export async function estimateFromSearch(get: GetJson, library: ILibraryResult, now: Date): Promise<ISearchEstimate> {
  const at = now.toISOString();
  const search = async (kql: string, refiners: string, extra = ''): Promise<IQueryResult> => {
    const url =
      `${library.webUrl}/_api/search/query?querytext='${encodeURIComponent(kql.replace(/'/g, "''"))}'` +
      `&trimduplicates=false&enablequeryrules=false` +
      (refiners ? `&refiners='${encodeURIComponent(refiners)}'` : '') +
      (extra || '&rowlimit=0');
    const json = await get<ISearchResponse & { d?: { query?: ISearchResponse } }>(url);
    const body: ISearchResponse = json.d && json.d.query ? json.d.query : json;
    const primary = body.PrimaryQueryResult || {};
    const relevant = primary.RelevantResults || {};
    const result: IQueryResult = { total: toNumber(relevant.TotalRows), types: new Map(), sizes: [], rows: [] };
    for (const refiner of list(primary.RefinementResults ? primary.RefinementResults.Refiners : undefined)) {
      for (const entry of list(refiner.Entries)) {
        const count = toNumber(entry.RefinementCount);
        if (refiner.Name.toLowerCase() === 'filetype') {
          const ext = normaliseType(entry.RefinementValue || entry.RefinementName);
          result.types.set(ext, (result.types.get(ext) || 0) + count);
        } else if (refiner.Name.toLowerCase() === 'size') {
          const range = parseRange(entry);
          if (range) {
            result.sizes.push({ ...range, count });
          }
        }
      }
    }
    for (const row of list(relevant.Table ? relevant.Table.Rows : undefined)) {
      const cells: Record<string, string> = {};
      list(row.Cells).forEach((c) => (cells[c.Key] = c.Value || ''));
      result.rows.push(cells);
    }
    return result;
  };

  const sizeRefiner = `Size(discretize=manual/${SIZE_BOUNDARIES.join('/')})`;
  const typeRefiner = `FileType(filter=${MAX_TYPES}/0/*)`;
  let scope = `ListId:${library.id} IsDocument:1`;
  const libraryPath = decodeURIComponent(new URL(library.url).pathname);

  try {
    // Largest old files first: exact sizes, and they show whether the library is in the index at all.
    const cutoff = ageBoundary(now, MIN_THRESHOLD_MONTHS);
    const largestQuery = (s: string): Promise<IQueryResult> =>
      search(
        `${s} LastModifiedTime<${kqlDate(cutoff)}`,
        '',
        `&rowlimit=${LARGEST_KEPT}&sortlist='${encodeURIComponent('Size:descending')}'` +
          `&selectproperties='${encodeURIComponent('Path,Size,LastModifiedTime,Filename')}'`
      );
    let largestResult = await largestQuery(scope);
    const everything = await search(scope, '');
    let anyFiles = everything.total;
    if (anyFiles === 0) {
      // Some tenants do not match ListId; fall back to the library's address.
      const byPath = `path:"${new URL(library.url).origin}${libraryPath}" IsDocument:1`;
      const check = await search(byPath, '');
      if (check.total > 0) {
        scope = byPath;
        anyFiles = check.total;
        largestResult = await largestQuery(scope);
      }
    }
    if (anyFiles === 0) {
      return {
        at,
        files: 0,
        bytes: 0,
        histogram: emptyHistogram(),
        fileTypes: [],
        largest: [],
        error: 'SharePoint search found no files in this library (it may be excluded from search, or not indexed yet)'
      };
    }

    const largest: ILargeFile[] = [];
    for (const row of largestResult.rows) {
      let path = '';
      try {
        path = decodeURIComponent(new URL(row.Path).pathname);
      } catch {
        continue;
      }
      const modified = new Date(row.LastModifiedTime);
      if (path.indexOf(libraryPath + '/') !== 0 || isNaN(modified.getTime())) {
        continue;
      }
      largest.push({
        name: row.Filename || path.substring(path.lastIndexOf('/') + 1),
        serverRelativeUrl: path,
        bytes: toNumber(row.Size),
        modified: modified.toISOString(),
        ageMonths: ageInMonths(modified, now),
        libraryTitle: library.title,
        webTitle: library.webTitle
      });
    }
    // Files above the top size range are sized by the ones search listed exactly.
    const huge = largest.filter((f) => f.bytes >= SIZE_BOUNDARIES[SIZE_BOUNDARIES.length - 1]);
    const largeMean = huge.length ? huge.reduce((s, f) => s + f.bytes, 0) / huge.length : undefined;
    const bytesOf = (sizes: IQueryResult['sizes']): { bytes: number; files: number } => {
      let bytes = 0;
      let files = 0;
      for (const s of sizes) {
        bytes += s.count * typicalSize(s.from, s.to, largeMean);
        files += s.count;
      }
      return { bytes, files };
    };

    // One query per age band: files, size ranges and types of files of that age.
    const bands: IQueryResult[] = [];
    for (const band of AGE_BANDS) {
      const restrictions: string[] = [];
      if (band.to !== undefined) {
        restrictions.push(`LastModifiedTime>=${kqlDate(ageBoundary(now, band.to))}`);
      }
      if (band.from > 0) {
        restrictions.push(`LastModifiedTime<${kqlDate(ageBoundary(now, band.from))}`);
      }
      bands.push(await search(`${scope} ${restrictions.join(' ')}`.trim(), `${typeRefiner},${sizeRefiner}`));
    }

    // Average size of the most common types, so a type of big files is not sized like a type of small ones.
    const typeTotals = new Map<string, number>();
    bands.forEach((b) => b.types.forEach((n, ext) => typeTotals.set(ext, (typeTotals.get(ext) || 0) + n)));
    const allTypes: string[] = [];
    typeTotals.forEach((_, ext) => allTypes.push(ext));
    const commonTypes = allTypes
      .filter((ext) => ext !== '(none)')
      .sort((a, b) => (typeTotals.get(b) || 0) - (typeTotals.get(a) || 0))
      .slice(0, TYPES_SIZED);
    const typeAverage = new Map<string, number>();
    for (const ext of commonTypes) {
      const r = await search(`${scope} FileType:"${ext}"`, sizeRefiner);
      const est = bytesOf(r.sizes);
      if (est.files > 0) {
        typeAverage.set(ext, est.bytes / est.files);
      }
    }

    const histogram = emptyHistogram();
    const types = new Map<string, IFileTypeStat>();
    let files = 0;
    let bytes = 0;
    bands.forEach((b, i) => {
      const est = bytesOf(b.sizes);
      const bandFiles = Math.max(b.total, est.files);
      // No size ranges came back: size the band like the library's average file, if known.
      const average =
        est.files > 0
          ? est.bytes / est.files
          : library.metrics && library.metrics.fileCount
          ? library.metrics.fileStreamSize / library.metrics.fileCount
          : 0;
      const bandBytes = est.files > 0 ? est.bytes * (bandFiles / est.files) : bandFiles * average;
      histogram.counts[AGE_BANDS[i].from] += bandFiles;
      histogram.bytes[AGE_BANDS[i].from] += bandBytes;
      files += bandFiles;
      bytes += bandBytes;

      // Types in this band, sized by their own average where known, then scaled to the band's bytes.
      const parts: { ext: string; count: number; weight: number }[] = [];
      let typed = 0;
      b.types.forEach((count, ext) => {
        parts.push({ ext, count, weight: count * (typeAverage.get(ext) || average) });
        typed += count;
      });
      if (bandFiles - typed > 0) {
        parts.push({ ext: '(other)', count: bandFiles - typed, weight: (bandFiles - typed) * average });
      }
      const weights = parts.reduce((s, p) => s + p.weight, 0);
      for (const p of parts) {
        let stat = types.get(p.ext);
        if (!stat) {
          stat = emptyFileTypeStat(p.ext);
          types.set(p.ext, stat);
        }
        stat.counts[i] += p.count;
        stat.bytes[i] += weights > 0 ? bandBytes * (p.weight / weights) : 0;
      }
    });

    // Scale bytes so the average file matches SharePoint's own figures for the library.
    const m = library.metrics;
    let indexCoverage: number | undefined;
    if (m && m.fileCount && m.fileCount > 0) {
      indexCoverage = files / m.fileCount;
      const target = Math.min(m.fileStreamSize, (m.fileStreamSize / m.fileCount) * files);
      if (bytes > 0 && target > 0) {
        const factor = target / bytes;
        histogram.bytes = histogram.bytes.map((v) => v * factor);
        types.forEach((t) => (t.bytes = t.bytes.map((v) => v * factor)));
        bytes = target;
      }
    }
    histogram.bytes = histogram.bytes.map(Math.round);
    const fileTypes: IFileTypeStat[] = [];
    types.forEach((t) => {
      t.bytes = t.bytes.map(Math.round);
      fileTypes.push(t);
    });
    fileTypes.sort((a, b) => b.bytes.reduce((x, y) => x + y, 0) - a.bytes.reduce((x, y) => x + y, 0));

    return {
      at,
      files,
      bytes: Math.round(bytes),
      histogram,
      fileTypes: fileTypes.slice(0, MAX_TYPES),
      largest,
      indexCoverage
    };
  } catch (err) {
    if (err instanceof ScanCancelledError || err instanceof ScanPausedError) {
      throw err;
    }
    return {
      at,
      files: 0,
      bytes: 0,
      histogram: emptyHistogram(),
      fileTypes: [],
      largest: [],
      error: `SharePoint search could not be used (${err instanceof Error ? err.message : String(err)})`
    };
  }
}
