import { IFileTypeStat } from './IFileTypeStat';

export interface ILibraryNode {
  id: string;
  title: string;
  serverRelativeUrl: string;
  absoluteUrl: string;
  itemCount: number;
  totalFiles: number;
  fileTypes: IFileTypeStat[];
  /** Search only returned this library's 10 most common types; rarer ones are missing. */
  typeListCapped?: boolean;
  scanned: boolean;
  error?: string;
}
