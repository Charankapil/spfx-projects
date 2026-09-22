import { IFileTypeStat } from './IFileTypeStat';

export interface ILibraryNode {
  id: string;
  title: string;
  serverRelativeUrl: string;
  absoluteUrl: string;
  itemCount: number;
  totalFiles: number;
  fileTypes: IFileTypeStat[];
  scanned: boolean;
  error?: string;
}
