import { IFileTypeStat } from './IFileTypeStat';
import { IStorageInfo } from './IStorageInfo';
import { IWebNode } from './IWebNode';

export interface ISiteCollectionOverview {
  siteUrl: string;
  siteTitle: string;
  storage: IStorageInfo;
  rootWeb: IWebNode;
  totalFileTypeStats: IFileTypeStat[];
  totalFiles: number;
  totalLibraries: number;
  totalWebs: number;
  scanStartedAt: Date;
  scanCompletedAt?: Date;
}
