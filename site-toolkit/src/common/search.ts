import { CancelToken } from './RequestQueue';
import { SpClient } from './SpClient';

/**
 * OData 3 ("verbose") wraps every collection in a { results: [] } envelope
 * while nometadata returns a bare array. Both shapes are accepted.
 */
type ODataCollection<T> = T[] | { results?: T[] } | undefined | null;

function toArray<T>(collection: ODataCollection<T>): T[] {
  if (!collection) {
    return [];
  }
  if (Array.isArray(collection)) {
    return collection;
  }
  return collection.results || [];
}

interface ICell {
  Key: string;
  Value: string | null;
}

interface IRow {
  Cells: ODataCollection<ICell>;
}

interface IRefinerEntry {
  RefinementName: string;
  RefinementCount: string;
  RefinementToken?: string;
}

interface IRefiner {
  Name: string;
  Entries: ODataCollection<IRefinerEntry>;
}

interface IPrimaryQueryResult {
  RelevantResults?: { TotalRows?: number | string; Table?: { Rows?: ODataCollection<IRow> } };
  RefinementResults?: { Refiners?: ODataCollection<IRefiner> };
}

interface ISearchResponse {
  PrimaryQueryResult?: IPrimaryQueryResult;
  d?: {
    query?: { PrimaryQueryResult?: IPrimaryQueryResult };
    PrimaryQueryResult?: IPrimaryQueryResult;
  };
}

export interface ISearchQuery {
  kql: string;
  select: string[];
  /** 1-500. */
  rowLimit: number;
  startRow?: number;
  /** e.g. "Size:descending" or "[DocId]:ascending". */
  sort?: string;
  /** Comma-separated refiner names, e.g. "FileType,DisplayAuthor". */
  refiners?: string;
  token?: CancelToken;
}

export interface IRefinerValue {
  name: string;
  count: number;
}

export type SearchRow = Record<string, string>;

export interface ISearchResult {
  rows: SearchRow[];
  total: number;
  refiners: Record<string, IRefinerValue[]>;
}

/** Search caps a page at 500 rows. */
export const SEARCH_PAGE_SIZE = 500;

function param(name: string, value: string): string {
  return `&${name}=${encodeURIComponent(`'${value.replace(/'/g, "''")}'`)}`;
}

/** Wraps a URL for a KQL `Path:"..."` restriction covering everything below it. */
export function underPath(url: string): string {
  return `Path:"${url.replace(/"/g, '')}/*"`;
}

/** Wraps free text as a KQL phrase. */
export function kqlPhrase(text: string): string {
  return `"${text.replace(/"/g, '')}"`;
}

/**
 * One call to the SharePoint search index. Results are security-trimmed to
 * the signed-in user by SharePoint itself.
 */
export async function search(sp: SpClient, query: ISearchQuery): Promise<ISearchResult> {
  const rowLimit = Math.max(0, Math.min(query.rowLimit, SEARCH_PAGE_SIZE));
  let url =
    `${sp.webUrl}/_api/search/query?querytext=${encodeURIComponent(`'${query.kql.replace(/'/g, "''")}'`)}` +
    `&rowlimit=${rowLimit}` +
    `&startrow=${query.startRow || 0}` +
    `&trimduplicates=false` +
    `&clienttype='ContentSearchRegular'`;
  if (query.select.length > 0) {
    url += param('selectproperties', query.select.join(','));
  }
  if (query.sort) {
    url += param('sortlist', query.sort);
  }
  if (query.refiners) {
    url += param('refiners', query.refiners);
  }

  const json = await sp.getJson<ISearchResponse>(url, { search: true, token: query.token });
  const primary =
    json.PrimaryQueryResult ||
    (json.d && json.d.query && json.d.query.PrimaryQueryResult) ||
    (json.d && json.d.PrimaryQueryResult);

  const relevant = primary && primary.RelevantResults;
  const totalRaw = relevant ? relevant.TotalRows : 0;
  const total = typeof totalRaw === 'number' ? totalRaw : parseInt(String(totalRaw), 10) || 0;

  const rows: SearchRow[] = toArray(relevant && relevant.Table ? relevant.Table.Rows : undefined).map((row) => {
    const out: SearchRow = {};
    for (const cell of toArray(row.Cells)) {
      out[cell.Key] = cell.Value === null || cell.Value === undefined ? '' : String(cell.Value);
    }
    return out;
  });

  const refiners: Record<string, IRefinerValue[]> = {};
  for (const refiner of toArray(primary && primary.RefinementResults ? primary.RefinementResults.Refiners : undefined)) {
    refiners[refiner.Name] = toArray(refiner.Entries)
      .map((entry) => ({
        name: entry.RefinementName || '(unknown)',
        count: parseInt(String(entry.RefinementCount || '0').replace(/[^\d]/g, ''), 10) || 0
      }))
      .sort((a, b) => b.count - a.count);
  }

  return { rows, total, refiners };
}

/** Last path segment of a document URL, i.e. the file name. */
export function fileNameOf(path: string): string {
  const clean = path.split('?')[0];
  return clean.substring(clean.lastIndexOf('/') + 1);
}

/** Parent folder URL of a document URL. */
export function folderOf(path: string): string {
  const clean = path.split('?')[0];
  return clean.substring(0, clean.lastIndexOf('/'));
}
