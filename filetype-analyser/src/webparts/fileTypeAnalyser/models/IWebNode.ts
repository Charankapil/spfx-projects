import { ILibraryNode } from './ILibraryNode';

export interface IWebNode {
  id: string;
  title: string;
  url: string;
  serverRelativeUrl: string;
  libraries: ILibraryNode[];
  webs: IWebNode[];
  error?: string;
}
