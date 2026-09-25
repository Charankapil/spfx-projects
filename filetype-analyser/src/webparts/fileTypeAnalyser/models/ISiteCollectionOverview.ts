import { IFileTypeStat } from './IFileTypeStat';
import { IStorageInfo } from './IStorageInfo';
import { IWebNode } from './IWebNode';

/** Headline numbers of an earlier scan, kept so the dashboard can show what changed. */
export interface IScanSummary {
  scanCompletedAt: Date;
  storageBytes?: number;
  totalFiles: number;
  totalLibraries: number;
  totalWebs: number;
  distinctTypes: number;
}

export interface ISiteCollectionOverview {
  siteUrl: string;
  siteTitle: string;
  storage: IStorageInfo;
  rootWeb: IWebNode;
  totalFileTypeStats: IFileTypeStat[];
  totalFiles: number;
  totalLibraries: number;
  totalWebs: number;
  /** Sum of library sizes; only meaningful when sizesAvailable is true. */
  totalLibraryBytes: number;
  sizesAvailable: boolean;
  /** True when at least one file type got a storage estimate (see IFileTypeStat.estimatedBytes). */
  typeSizesEstimated?: boolean;
  totalEstimatedBytes?: number;
  /** File types that got no estimate (beyond the per-scan cap, or the query failed). */
  unmeasuredTypes?: number;
  scanStartedAt: Date;
  scanCompletedAt?: Date;
  scannedBy?: string;
  previous?: IScanSummary;
}
