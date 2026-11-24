import { IFolderStat, ILibraryResult } from '../models/IScanResult';
import { HttpError, ScanCancelledError, ScanPausedError } from './errors';

/** Folders deeper than this are not listed. */
export const MAX_FOLDER_DEPTH = 4;
/** Folders kept per library, biggest first. */
export const MAX_FOLDERS_KEPT = 300;
/** A folder is only opened when it holds at least this share of the library's storage. */
const MIN_SHARE_TO_OPEN = 0.02;

interface IFolderRow {
  Name?: string;
  ServerRelativeUrl?: string;
  ItemCount?: number | string;
  StorageMetrics?: IMetricsRow;
}

interface IMetricsRow {
  TotalSize?: number | string;
  TotalFileStreamSize?: number | string;
  TotalFileCount?: number | string;
  LastModified?: string;
}

export type GetJson = <T>(url: string) => Promise<T>;

function toNumber(value: number | string | undefined | null): number {
  const parsed = typeof value === 'number' ? value : parseFloat(String(value === undefined || value === null ? '' : value));
  return isNaN(parsed) ? 0 : parsed;
}

function rows<T>(json: { value?: T[]; d?: { results?: T[] } }): T[] {
  return json.value || (json.d && json.d.results) || [];
}

/** Server-relative path for GetFolderByServerRelativePath: quotes doubled, then URL-encoded (slashes kept). */
export function quotePath(serverRelativePath: string): string {
  return encodeURIComponent(serverRelativePath.replace(/'/g, "''")).replace(/%2F/g, '/');
}

function libraryRoot(library: ILibraryResult): string {
  return decodeURIComponent(new URL(library.url).pathname).replace(/\/+$/, '');
}

function rethrowIfStopped(err: unknown): void {
  if (err instanceof ScanCancelledError || err instanceof ScanPausedError) {
    throw err;
  }
}

/**
 * The biggest folders of a library with SharePoint's own storage figures,
 * without reading any files. Folders are opened biggest first, so the
 * request budget goes where the storage is; folders holding little of the
 * library are listed but not opened. Works the same for flat libraries (no
 * folders to list), a few huge folders, or deep trees.
 */
export async function walkFolders(
  get: GetJson,
  library: ILibraryResult,
  maxRequests: number
): Promise<{ folders: IFolderStat[]; note?: string }> {
  const root = libraryRoot(library);
  const base = `${library.webUrl}/_api/web`;
  const libraryTotal = library.metrics ? library.metrics.totalSize : 0;
  let withMetrics = true;
  let requests = 0;
  let budgetSpent = false;

  const metricsOf = (m: IMetricsRow | undefined): Pick<IFolderStat, 'files' | 'bytes' | 'totalSize' | 'lastModified'> => ({
    files: toNumber(m && m.TotalFileCount),
    bytes: toNumber(m && m.TotalFileStreamSize),
    totalSize: toNumber(m && m.TotalSize),
    lastModified: m && m.LastModified && !isNaN(new Date(m.LastModified).getTime()) ? m.LastModified : undefined
  });

  const listChildren = async (path: string): Promise<IFolderRow[]> => {
    const folder =
      path === root ? `${base}/lists(guid'${library.id}')/RootFolder` : `${base}/GetFolderByServerRelativePath(decodedurl='${quotePath(path)}')`;
    if (withMetrics) {
      try {
        requests++;
        return rows<IFolderRow>(
          await get(`${folder}/Folders?$select=Name,ServerRelativeUrl,ItemCount,StorageMetrics&$expand=StorageMetrics`)
        );
      } catch (err) {
        rethrowIfStopped(err);
        // Some tenants do not expand StorageMetrics on a folder collection: ask for each folder's figures instead.
        if (!(err instanceof HttpError) || err.status !== 400) {
          throw err;
        }
        withMetrics = false;
      }
    }
    requests++;
    const children = rows<IFolderRow>(await get(`${folder}/Folders?$select=Name,ServerRelativeUrl,ItemCount`));
    const measured: IFolderRow[] = [];
    for (const child of children) {
      if (requests >= maxRequests) {
        budgetSpent = true;
        break;
      }
      if (!child.ServerRelativeUrl || (path === root && (child.Name || '').toLowerCase() === 'forms')) {
        continue;
      }
      try {
        requests++;
        child.StorageMetrics = await get<IMetricsRow>(
          `${base}/GetFolderByServerRelativePath(decodedurl='${quotePath(child.ServerRelativeUrl)}')/StorageMetrics`
        );
        measured.push(child);
      } catch (err) {
        rethrowIfStopped(err);
      }
    }
    // A folder without figures would show as empty: only folders that were measured are kept.
    return measured;
  };

  const found: IFolderStat[] = [];
  // Folders waiting to be opened, biggest first.
  const queue: { path: string; depth: number; stat?: IFolderStat }[] = [{ path: root, depth: 0 }];
  let note: string | undefined;

  while (queue.length > 0) {
    if (requests >= maxRequests) {
      budgetSpent = true;
      break;
    }
    queue.sort((a, b) => (b.stat ? b.stat.totalSize : Infinity) - (a.stat ? a.stat.totalSize : Infinity));
    const next = queue.shift() as { path: string; depth: number; stat?: IFolderStat };
    let children: IFolderRow[];
    try {
      children = await listChildren(next.path);
    } catch (err) {
      rethrowIfStopped(err);
      const reason = err instanceof Error ? err.message : String(err);
      if (!next.stat) {
        return { folders: [], note: `Folders could not be listed (${reason})` };
      }
      next.stat.childrenUnavailable = true;
      continue;
    }
    const total = libraryTotal || children.reduce((sum, c) => sum + toNumber(c.StorageMetrics && c.StorageMetrics.TotalSize), 0);
    for (const child of children) {
      const url = child.ServerRelativeUrl || '';
      // "Forms" in the library root holds the library's own views and templates, not people's files.
      if (!url || url.indexOf(root + '/') !== 0 || (next.depth === 0 && (child.Name || '').toLowerCase() === 'forms')) {
        continue;
      }
      const stat: IFolderStat = { path: url.substring(root.length + 1), depth: next.depth + 1, ...metricsOf(child.StorageMetrics) };
      found.push(stat);
      if (stat.depth < MAX_FOLDER_DEPTH && toNumber(child.ItemCount) > 0 && stat.totalSize >= total * MIN_SHARE_TO_OPEN) {
        queue.push({ path: url, depth: stat.depth, stat });
      }
    }
  }

  found.sort((a, b) => b.totalSize - a.totalSize);
  if (found.length === 0) {
    note = 'No folders: every file is directly in the library';
  } else if (found.length > MAX_FOLDERS_KEPT) {
    note = `Showing the ${MAX_FOLDERS_KEPT} biggest of ${found.length} folders`;
  } else if (budgetSpent) {
    note = 'Smaller folders were not opened, to keep the number of requests to SharePoint low';
  }
  return { folders: found.slice(0, MAX_FOLDERS_KEPT), note };
}
