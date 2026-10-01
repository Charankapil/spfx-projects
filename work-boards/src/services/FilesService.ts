import { SpClient } from './SpClient';

export interface ILibrary {
  title: string;
  /** Server-relative URL of the library's root folder. */
  url: string;
}

export interface IFolderEntry {
  name: string;
  /** Server-relative URL. */
  url: string;
  isFolder: boolean;
  modified?: string;
  size?: number;
}

/** Browse document libraries on the current site, to link files to items. Read only. */
export class FilesService {
  private readonly sp: SpClient;

  constructor(sp: SpClient) {
    this.sp = sp;
  }

  public origin(): string {
    return this.sp.webUrl.replace(/^(https?:\/\/[^/]+).*$/, '$1');
  }

  public async libraries(): Promise<ILibrary[]> {
    const res = await this.sp.get<{ value: { Title: string; RootFolder: { ServerRelativeUrl: string } }[] }>(
      "web/lists?$select=Title,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=BaseTemplate eq 101 and Hidden eq false"
    );
    return res.value
      .filter(l => l.Title !== 'Form Templates' && l.Title !== 'Style Library' && l.Title !== 'Site Assets')
      .map(l => ({ title: l.Title, url: l.RootFolder.ServerRelativeUrl }))
      .sort((a, b) => a.title.localeCompare(b.title));
  }

  /** Sub-folders then files of a folder, by name. */
  public async folder(serverRelativeUrl: string): Promise<IFolderEntry[]> {
    const path = `web/GetFolderByServerRelativePath(decodedurl='${encodeURIComponent(serverRelativeUrl.replace(/'/g, "''"))}')`;
    const [folders, files] = await Promise.all([
      this.sp.get<{ value: { Name: string; ServerRelativeUrl: string }[] }>(`${path}/Folders?$select=Name,ServerRelativeUrl`),
      this.sp.get<{ value: { Name: string; ServerRelativeUrl: string; TimeLastModified: string; Length: string }[] }>(
        `${path}/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Length`
      )
    ]);
    const byName = (a: IFolderEntry, b: IFolderEntry): number => a.name.localeCompare(b.name);
    return folders.value
      .filter(f => f.Name !== 'Forms')
      .map(f => ({ name: f.Name, url: f.ServerRelativeUrl, isFolder: true }))
      .sort(byName)
      .concat(files.value.map(f => ({ name: f.Name, url: f.ServerRelativeUrl, isFolder: false, modified: f.TimeLastModified, size: Number(f.Length) })).sort(byName));
  }
}
