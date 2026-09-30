import { ILibraryResult, IScanResult, IWebResult } from '../models/IScanResult';

/** Libraries at or above either figure are flagged as large on the map. */
export const LARGE_LIBRARY_FILES = 250000;
export const LARGE_LIBRARY_BYTES = 200 * 1024 * 1024 * 1024;

export function libraryKey(lib: ILibraryResult): string {
  return `${lib.webUrl}|${lib.id}`;
}

/** Size shown on the map: SharePoint's own figure (with version history) when known, otherwise what the scan read. */
export function librarySize(lib: ILibraryResult): number {
  return lib.metrics ? lib.metrics.totalSize : lib.bytes;
}

export function libraryFiles(lib: ILibraryResult): number {
  if (lib.metrics && lib.metrics.fileCount !== undefined) {
    return lib.metrics.fileCount;
  }
  return lib.unscanned ? lib.itemCount : lib.files;
}

export function isLarge(lib: ILibraryResult): boolean {
  return libraryFiles(lib) >= LARGE_LIBRARY_FILES || librarySize(lib) >= LARGE_LIBRARY_BYTES;
}

export interface ICoverage {
  scanned: number;
  total: number;
  scannedBytes: number;
  totalBytes: number;
}

/** How much of the site has been read: libraries, and share of the storage SharePoint reports for them. */
export function coverage(result: IScanResult): ICoverage {
  const c: ICoverage = { scanned: 0, total: result.libraries.length, scannedBytes: 0, totalBytes: 0 };
  for (const lib of result.libraries) {
    const size = librarySize(lib);
    c.totalBytes += size;
    if (!lib.unscanned && !lib.pending) {
      c.scanned++;
      c.scannedBytes += size;
    }
  }
  return c;
}

export function isPartialCoverage(result: IScanResult): boolean {
  return result.libraries.some((l) => l.unscanned);
}

export interface IMapWeb {
  web: IWebResult;
  depth: number;
  /** Libraries of this site only. */
  own: ILibraryResult[];
  /** Libraries of this site and every subsite below it. */
  all: ILibraryResult[];
}

function pathOf(url: string): string {
  try {
    return decodeURIComponent(new URL(url).pathname).replace(/\/+$/, '').toLowerCase();
  } catch {
    return url.replace(/\/+$/, '').toLowerCase();
  }
}

/** Sites in discovery order (parents before their subsites) with the libraries of each. */
export function buildMap(result: IScanResult): IMapWeb[] {
  const rootDepth = pathOf(result.rootUrl).split('/').length;
  const nodes: IMapWeb[] = result.webs.map((web) => ({
    web,
    depth: Math.max(0, pathOf(web.url).split('/').length - rootDepth),
    own: result.libraries.filter((l) => l.webUrl === web.url),
    all: []
  }));
  for (const node of nodes) {
    const prefix = pathOf(node.web.url) + '/';
    node.all = result.libraries.filter((l) => l.webUrl === node.web.url || pathOf(l.webUrl).indexOf(prefix) === 0);
  }
  return nodes;
}

export function selectionTotals(result: IScanResult, keys: Set<string>): { count: number; bytes: number; files: number; large: number } {
  const totals = { count: 0, bytes: 0, files: 0, large: 0 };
  for (const lib of result.libraries) {
    if (keys.has(libraryKey(lib))) {
      totals.count++;
      totals.bytes += librarySize(lib);
      totals.files += libraryFiles(lib);
      if (isLarge(lib)) {
        totals.large++;
      }
    }
  }
  return totals;
}
