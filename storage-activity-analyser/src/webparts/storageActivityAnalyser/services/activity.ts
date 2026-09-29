import { AGE_BUCKET_COUNT, IAgeHistogram, IFileTypeStat, ILibraryResult, MAX_AGE_MONTHS } from '../models/IScanResult';

/** Inactivity thresholds offered in the property pane and on the dashboard. */
export const THRESHOLD_OPTIONS: number[] = [3, 6, 12, 24, 36, 60];
export const DEFAULT_THRESHOLD_MONTHS = 12;
export const MIN_THRESHOLD_MONTHS = THRESHOLD_OPTIONS[0];

export function thresholdLabel(months: number): string {
  if (months % 12 === 0) {
    const years = months / 12;
    return years === 1 ? '1 year' : `${years} years`;
  }
  return `${months} months`;
}

export function emptyHistogram(): IAgeHistogram {
  const counts: number[] = [];
  const bytes: number[] = [];
  for (let i = 0; i < AGE_BUCKET_COUNT; i++) {
    counts.push(0);
    bytes.push(0);
  }
  return { counts, bytes };
}

/**
 * Whole calendar months between `modified` and `now`: a file changed on
 * 30 Sep 2025 is 11 months old on 29 Sep 2026 and 12 months old on 30 Sep
 * 2026. Future dates (clock skew) count as 0.
 */
export function ageInMonths(modified: Date, now: Date): number {
  let months = (now.getFullYear() - modified.getFullYear()) * 12 + (now.getMonth() - modified.getMonth());
  if (now.getDate() < modified.getDate()) {
    months--;
  }
  return Math.max(0, Math.min(MAX_AGE_MONTHS, months));
}

export function addToHistogram(target: IAgeHistogram, source: IAgeHistogram): void {
  for (let i = 0; i < AGE_BUCKET_COUNT; i++) {
    target.counts[i] += source.counts[i] || 0;
    target.bytes[i] += source.bytes[i] || 0;
  }
}

export interface IActivitySplit {
  activeFiles: number;
  activeBytes: number;
  inactiveFiles: number;
  inactiveBytes: number;
  totalFiles: number;
  totalBytes: number;
}

/** A file is inactive when it has not been modified for `thresholdMonths` or more. */
export function splitByThreshold(histogram: IAgeHistogram, thresholdMonths: number): IActivitySplit {
  const split: IActivitySplit = {
    activeFiles: 0,
    activeBytes: 0,
    inactiveFiles: 0,
    inactiveBytes: 0,
    totalFiles: 0,
    totalBytes: 0
  };
  for (let i = 0; i < AGE_BUCKET_COUNT; i++) {
    const count = histogram.counts[i] || 0;
    const bytes = histogram.bytes[i] || 0;
    if (i >= thresholdMonths) {
      split.inactiveFiles += count;
      split.inactiveBytes += bytes;
    } else {
      split.activeFiles += count;
      split.activeBytes += bytes;
    }
  }
  split.totalFiles = split.activeFiles + split.inactiveFiles;
  split.totalBytes = split.activeBytes + split.inactiveBytes;
  return split;
}

/** Age in months of the most recently modified file, or undefined for an empty histogram. */
export function newestFileAge(histogram: IAgeHistogram): number | undefined {
  for (let i = 0; i < AGE_BUCKET_COUNT; i++) {
    if (histogram.counts[i] > 0) {
      return i;
    }
  }
  return undefined;
}

export function totalHistogram(libraries: ILibraryResult[]): IAgeHistogram {
  const total = emptyHistogram();
  for (const library of libraries) {
    addToHistogram(total, library.histogram);
  }
  return total;
}

export interface IAgeBand {
  /** Inclusive lower bound in months. */
  from: number;
  /** Exclusive upper bound in months; undefined for the open-ended last band. */
  to?: number;
  label: string;
}

/**
 * Bands for the age chart. Every threshold option is a band boundary, so
 * each band is either wholly active or wholly inactive.
 */
export const AGE_BANDS: IAgeBand[] = [
  { from: 0, to: 3, label: '< 3 mo' },
  { from: 3, to: 6, label: '3–6 mo' },
  { from: 6, to: 12, label: '6–12 mo' },
  { from: 12, to: 24, label: '1–2 yr' },
  { from: 24, to: 36, label: '2–3 yr' },
  { from: 36, to: 60, label: '3–5 yr' },
  { from: 60, to: 120, label: '5–10 yr' },
  { from: 120, label: '10+ yr' }
];

export function bandTotals(histogram: IAgeHistogram, band: IAgeBand): { files: number; bytes: number } {
  let files = 0;
  let bytes = 0;
  const end = band.to === undefined ? AGE_BUCKET_COUNT : Math.min(band.to, AGE_BUCKET_COUNT);
  for (let i = band.from; i < end; i++) {
    files += histogram.counts[i] || 0;
    bytes += histogram.bytes[i] || 0;
  }
  return { files, bytes };
}

export function describeAge(months: number): string {
  if (months <= 0) {
    return 'this month';
  }
  if (months >= MAX_AGE_MONTHS) {
    return '10+ years ago';
  }
  if (months < 12) {
    return months === 1 ? '1 month ago' : `${months} months ago`;
  }
  const years = Math.floor(months / 12);
  const rest = months % 12;
  const yearText = years === 1 ? '1 year' : `${years} years`;
  return rest === 0 ? `${yearText} ago` : `${yearText} ${rest} mo ago`;
}

/** Index into AGE_BANDS for a file of the given age in months. */
export function bandIndex(ageMonths: number): number {
  for (let i = AGE_BANDS.length - 1; i > 0; i--) {
    if (ageMonths >= AGE_BANDS[i].from) {
      return i;
    }
  }
  return 0;
}

export function emptyFileTypeStat(extension: string): IFileTypeStat {
  return { extension, counts: AGE_BANDS.map(() => 0), bytes: AGE_BANDS.map(() => 0) };
}

/**
 * The extension of a file name, lower-case and without the dot. Names with
 * no dot, or with something after the last dot that does not look like an
 * extension (too long, spaces), count as "(none)".
 */
export function extensionOf(fileName: string): string {
  const dot = fileName.lastIndexOf('.');
  if (dot <= 0 || dot === fileName.length - 1) {
    return '(none)';
  }
  const ext = fileName.substring(dot + 1).toLowerCase();
  return ext.length > 10 || /\s/.test(ext) ? '(none)' : ext;
}

export interface IFileTypeSplit extends IActivitySplit {
  extension: string;
}

export function splitFileType(stat: IFileTypeStat, thresholdMonths: number): IFileTypeSplit {
  const split: IFileTypeSplit = {
    extension: stat.extension,
    activeFiles: 0,
    activeBytes: 0,
    inactiveFiles: 0,
    inactiveBytes: 0,
    totalFiles: 0,
    totalBytes: 0
  };
  AGE_BANDS.forEach((band, i) => {
    if (band.from >= thresholdMonths) {
      split.inactiveFiles += stat.counts[i] || 0;
      split.inactiveBytes += stat.bytes[i] || 0;
    } else {
      split.activeFiles += stat.counts[i] || 0;
      split.activeBytes += stat.bytes[i] || 0;
    }
  });
  split.totalFiles = split.activeFiles + split.inactiveFiles;
  split.totalBytes = split.activeBytes + split.inactiveBytes;
  return split;
}

/** Adds up the file types of all libraries; undefined when no library has file-type data (older saved scans). */
export function totalFileTypes(libraries: ILibraryResult[]): IFileTypeStat[] | undefined {
  const map: { [ext: string]: IFileTypeStat } = {};
  let any = false;
  for (const library of libraries) {
    if (!library.fileTypes) {
      continue;
    }
    any = true;
    for (const stat of library.fileTypes) {
      const total = map[stat.extension] || (map[stat.extension] = emptyFileTypeStat(stat.extension));
      for (let i = 0; i < AGE_BANDS.length; i++) {
        total.counts[i] += stat.counts[i] || 0;
        total.bytes[i] += stat.bytes[i] || 0;
      }
    }
  }
  if (!any) {
    return undefined;
  }
  return Object.keys(map).map((k) => map[k]);
}
