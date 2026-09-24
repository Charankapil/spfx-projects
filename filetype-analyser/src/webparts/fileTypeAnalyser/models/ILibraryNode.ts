import { IFileTypeStat } from './IFileTypeStat';

export interface ILibraryNode {
  id: string;
  title: string;
  webUrl: string;
  serverRelativeUrl: string;
  absoluteUrl: string;
  itemCount: number;
  totalFiles: number;
  /** Library storage (TotalSize from StorageMetrics, bytes); undefined when unavailable. */
  sizeBytes?: number;
  fileTypes: IFileTypeStat[];
  scanned: boolean;
  error?: string;
}
