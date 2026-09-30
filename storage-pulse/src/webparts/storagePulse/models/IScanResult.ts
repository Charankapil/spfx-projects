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
  /** Files in the library, per SharePoint (recursive). Missing on scans saved before v2.1. */
  fileCount?: number;
  /** Most recent change anywhere in the library, per SharePoint (ISO). */
  lastModified?: string;
}

/**
 * Files of one extension, split into the dashboard's age bands (see
 * AGE_BANDS in services/activity.ts). Every inactivity threshold is a band
 * boundary, so the active / inactive split per type is exact for any
 * threshold without keeping month-level detail per type.
 */
export interface IFileTypeStat {
  /** Lower-case extension without the dot, "(none)" or "(other)". */
  extension: string;
  counts: number[];
  bytes: number[];
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
  /** Size by file type; missing on scans saved before v1.1. */
  fileTypes?: IFileTypeStat[];
  metrics?: ILibraryStorageMetrics;
  /** Set when the library could not be read, or only part of it could. */
  error?: string;
  partial?: boolean;
  /**
   * Items the library reports (ItemCount) that the scan did not see, usually
   * because they have permissions that exclude the person scanning.
   */
  unreadItems?: number;
  /**
   * Quick scan: the library had no changes for at least `dormantMonths`, so it
   * was measured as a whole from SharePoint's storage metrics instead of file
   * by file. Its files are all counted at that age (a lower bound), and it has
   * no file-type breakdown or largest-file entries.
   */
  measuredAsWhole?: { lastChange: string; dormantMonths: number };
  /**
   * Only in a checkpoint of an unfinished scan: this library has not been read
   * yet (or was being read when the checkpoint was taken), so it starts again
   * on resume.
   */
  pending?: boolean;
  /** The list's last change by a person (ISO). Kept only while pending or unscanned, for the quick-scan check when it is read. */
  lastUserChange?: string;
  /**
   * Found on the site map but not read yet (v2.3): files, bytes and histogram
   * are empty, and `metrics` holds SharePoint's own size and file count so the
   * map can show them. Left out of the dashboard figures.
   */
  unscanned?: boolean;
  /** When this library was last read (ISO). Missing on scans saved before v2.3. */
  scannedAt?: string;
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
  /** Which libraries were included; missing on scans saved before v2. */
  options?: { includeHidden: boolean; excludeSystemLibraries: boolean; excludedLibraries: string[] };
  /**
   * Quick scans measure libraries unchanged for at least this many months as
   * a whole. Splits at this threshold or lower are exact; longer ones are
   * approximate for those libraries. Missing on detailed scans.
   */
  quickAfterMonths?: number;
  /**
   * Set when SharePoint rejected the fast item query (HTTP 406) for some
   * libraries and a more conservative form was used for them.
   */
  paging?: { level: number; libraries: number };
  /**
   * Only in a checkpoint of an unfinished scan. Libraries still `pending`
   * are read again on resume; ages stay measured from scanStartedAt.
   */
  partial?: { librariesDone: number; librariesTotal: number };
}

export interface IScanProgress {
  phase: 'idle' | 'starting' | 'discovering' | 'reading-files' | 'completed';
  currentItem: string;
  websFound: number;
  librariesFound: number;
  librariesDone: number;
  /** Sum of the libraries' ItemCount (files and folders), used to estimate overall progress. */
  itemsExpected: number;
  /** Items (files and folders) read so far, comparable with itemsExpected. */
  itemsRead: number;
  filesRead: number;
  bytesRead: number;
  /** Present once SharePoint has throttled the scan. */
  throttle?: {
    /** Epoch ms until which requests are on hold; in the past when not on hold. */
    pausedUntil: number;
    throttledCount: number;
    /** Requests allowed at once now, and at most. */
    concurrency: number;
    maxConcurrency: number;
    waitedMs: number;
  };
}
