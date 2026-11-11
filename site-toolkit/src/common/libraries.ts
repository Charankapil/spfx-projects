import { cached, cacheKey } from './cache';
import { SpClient } from './SpClient';

export interface ILibrary {
  id: string;
  title: string;
  /** Server-relative URL of the library's root folder. */
  path: string;
  itemCount: number;
}

interface ILibraryRow {
  Id: string;
  Title: string;
  ItemCount: number;
  RootFolder: { ServerRelativeUrl: string };
}

/** The site's visible document libraries: one request, reused for 15 minutes. */
export function getLibraries(sp: SpClient, scope: string): Promise<ILibrary[]> {
  return cached('session', cacheKey(scope, 'libraries'), 15, async () => {
    const json = await sp.getJson<{ value: ILibraryRow[] }>(
      `${sp.webUrl}/_api/web/lists?$filter=BaseTemplate eq 101 and Hidden eq false` +
        `&$select=Id,Title,ItemCount,RootFolder/ServerRelativeUrl&$expand=RootFolder`
    );
    return (json.value || [])
      .map((row) => ({ id: row.Id, title: row.Title, path: row.RootFolder.ServerRelativeUrl, itemCount: row.ItemCount }))
      .sort((a, b) => a.title.localeCompare(b.title));
  });
}
