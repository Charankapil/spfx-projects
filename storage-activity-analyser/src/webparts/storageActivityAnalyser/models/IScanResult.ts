export type ScanScope = 'siteCollection' | 'currentWeb';

/**
 * Files are counted into one bucket per whole month of age (time since last
 * modified, measured from when the scan started). Bucket 0 holds files
 * changed in the last month; the final bucket holds everything ten years or
 * older. Keeping month-level detail means the inactivity threshold can be
 * changed on the dashboard without scanning again.
 */
export const MAX_AGE_MONTHS = 120;
export const AGE_BUCKET_COUNT = MAX_AGE_MONTHS + 1;

export interface IAgeHistogram {
  /** Files per month of age; index = age in whole months, capped at MAX_AGE_MONTHS. */
  counts: number[];
  /** Bytes of the current file version per month of age. */
  bytes: number[];
}

/**
 * SharePoint's own storage figures for a library's root folder. TotalSize
 * includes version history; TotalFileStreamSize is current versions only.
 */
export interface ILibraryStorageMetrics {
  totalSize: number;
  fileStreamSize: number;
}

export interface ILibraryResult {
  id: string;
  title: string;
  webTitle: string;
  webUrl: string;
  url: string;
  /** Items according to the list itself (files and folders). */
  itemCount: number;
  files: number;
  bytes: number;
  histogram: IAgeHistogram;
  metrics?: ILibraryStorageMetrics;
  /** Set when the library could not be read, or only part of it could. */
  error?: string;
  partial?: boolean;
}

export interface IWebResult {
  title: string;
  url: string;
  error?: string;
}

export interface ILargeFile {
  name: string;
  serverRelativeUrl: string;
  bytes: number;
  modified: string;
  ageMonths: number;
  libraryTitle: string;
  webTitle: string;
}

export interface IScanResult {
  scope: ScanScope;
  rootUrl: string;
  rootTitle: string;
  /** Site collection storage used, as reported by SharePoint (includes versions and recycle bins). */
  siteStorageBytes?: number;
  webs: IWebResult[];
  libraries: ILibraryResult[];
  /**
   * The largest files at least MIN_THRESHOLD_MONTHS old, biggest first, so
   * the list stays correct for every threshold the dashboard offers.
   */
  largestOldFiles: ILargeFile[];
  /** Ages are measured from this moment (ISO string). */
  scanStartedAt: string;
  scanCompletedAt: string;
  scannedBy: string;
}

export interface IScanProgress {
  phase: 'idle' | 'starting' | 'discovering' | 'reading-files' | 'completed';
  currentItem: string;
  websFound: number;
  librariesFound: number;
  librariesDone: number;
  /** Sum of the libraries' ItemCount, used to estimate overall progress. */
  itemsExpected: number;
  filesRead: number;
  bytesRead: number;
}
