const KB = 1024;
const MB = KB * 1024;
const GB = MB * 1024;

/**
 * Band edges for the search "Size" refiner. Search returns one bin below the
 * first edge, one between each pair and one above the last. Roughly
 * log-spaced so no band spans more than ~5x, which keeps the per-band
 * estimate (the geometric middle of the band) within a sensible error.
 */
export const SIZE_THRESHOLDS: number[] = [
  10 * KB, 50 * KB, 100 * KB, 250 * KB, 500 * KB,
  1 * MB, 2.5 * MB, 5 * MB, 10 * MB, 25 * MB, 50 * MB, 100 * MB, 250 * MB, 500 * MB,
  1 * GB, 2.5 * GB, 5 * GB, 10 * GB
].map((n) => Math.round(n));

export interface ISizeBand {
  lo: number;
  hi?: number;
}

export interface ISizeBandEntry {
  RefinementName?: string;
  RefinementToken?: string;
  RefinementCount?: string | number;
}

function parseBound(raw: string, isLower: boolean): number | undefined {
  const text = raw.trim().toLowerCase();
  if (text.indexOf('min') >= 0) {
    return isLower ? 0 : undefined;
  }
  if (text.indexOf('max') >= 0) {
    return undefined;
  }
  // Take the last number rather than stripping letters: "decimal(10240)"
  // contains an "e", and in "int64(1048576)" the first number is the 64.
  const pattern = /-?\d+(?:\.\d+)?(?:e[+-]?\d+)?/gi;
  let last: string | undefined;
  let match = pattern.exec(text);
  while (match) {
    last = match[0];
    match = pattern.exec(text);
  }
  const n = last !== undefined ? parseFloat(last) : NaN;
  return isFinite(n) ? n : undefined;
}

/**
 * Reads a band's range from its refinement token, e.g. "range(min, 10240)",
 * "range(10240, 51200)", "range(10737418240, max)" - tolerating wrappers
 * such as decimal(...) around the numbers. Returns undefined when the token
 * is not a range, so the caller can count it as unparsed rather than guess.
 */
export function parseBand(token: string | undefined): ISizeBand | undefined {
  if (!token) {
    return undefined;
  }
  const match = /range\(\s*([^,]+),\s*(.+)\)\s*$/i.exec(token.trim());
  if (!match) {
    return undefined;
  }
  const lo = parseBound(match[1], true);
  const hi = parseBound(match[2], false);
  if (lo === undefined) {
    return undefined;
  }
  if (hi !== undefined && hi <= lo) {
    return undefined;
  }
  return { lo, hi };
}

/** A typical file size for a band: its geometric middle, half its top for the lowest band, 1.5x its floor for the open top band. */
export function representativeSize(band: ISizeBand): number {
  if (band.hi === undefined) {
    return band.lo * 1.5;
  }
  if (band.lo <= 0) {
    return band.hi / 2;
  }
  return Math.sqrt(band.lo * band.hi);
}

export interface ISizeEstimate {
  bytes: number;
  files: number;
  unparsedFiles: number;
}

export function estimateFromBands(entries: ISizeBandEntry[]): ISizeEstimate {
  let bytes = 0;
  let files = 0;
  let unparsedFiles = 0;
  entries.forEach((entry) => {
    const count = parseInt(String(entry.RefinementCount ?? '0').replace(/[^\d]/g, ''), 10) || 0;
    if (count <= 0) {
      return;
    }
    const band = parseBand(entry.RefinementToken) || parseBand(entry.RefinementName);
    if (!band) {
      unparsedFiles += count;
      return;
    }
    bytes += count * representativeSize(band);
    files += count;
  });
  return { bytes: Math.round(bytes), files, unparsedFiles };
}
