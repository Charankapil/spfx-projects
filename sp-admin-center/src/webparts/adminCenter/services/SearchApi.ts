import { entityOf, HttpError, odataString, SPClient, toArray, ODataCollection, trimSlash } from '../core/SPClient';
import { IFileRow, IFileTypeCount, ISiteRow } from '../models';

/**
 * Everything here is answered by SharePoint's search index: one request returns
 * an aggregate (counts, top-N), so the cost never grows with the number of
 * files. That is what keeps the Storage and Sites views safe on tenants with
 * millions of items - we never enumerate lists or libraries to get a number.
 */

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
}
interface IRefiner {
  Name: string;
  Entries: ODataCollection<IRefinerEntry>;
}
interface IPrimary {
  RelevantResults?: { TotalRows?: number; Table?: { Rows?: ODataCollection<IRow> } };
  RefinementResults?: { Refiners?: ODataCollection<IRefiner> };
}
interface ISearchResponse {
  PrimaryQueryResult?: IPrimary;
  query?: { PrimaryQueryResult?: IPrimary };
  d?: { query?: { PrimaryQueryResult?: IPrimary }; PrimaryQueryResult?: IPrimary };
}

export interface ISearchResult {
  rows: Array<{ [key: string]: string }>;
  total: number;
  refiners: { [name: string]: Array<{ name: string; count: number }> };
}

export interface ISearchOptions {
  query: string;
  select?: string[];
  rowLimit?: number;
  startRow?: number;
  sort?: string;
  refiners?: string;
}

/** Quote a value for KQL: "..." with inner quotes removed. */
export function kqlQuote(value: string): string {
  return '"' + value.replace(/"/g, '') + '"';
}

export function parseSearch(json: ISearchResponse): ISearchResult {
  const primary =
    (json.d && (json.d.query ? json.d.query.PrimaryQueryResult : json.d.PrimaryQueryResult)) ||
    (json.query ? json.query.PrimaryQueryResult : undefined) ||
    json.PrimaryQueryResult;
  const result: ISearchResult = { rows: [], total: 0, refiners: {} };
  if (!primary) {
    return result;
  }
  const relevant = primary.RelevantResults;
  if (relevant) {
    result.total = Number(relevant.TotalRows) || 0;
    toArray(relevant.Table && relevant.Table.Rows).forEach((row) => {
      const record: { [key: string]: string } = {};
      toArray(row.Cells).forEach((cell) => {
        record[cell.Key] = cell.Value === null || cell.Value === undefined ? '' : String(cell.Value);
      });
      result.rows.push(record);
    });
  }
  const refiners = primary.RefinementResults && primary.RefinementResults.Refiners;
  toArray(refiners).forEach((r) => {
    result.refiners[r.Name] = toArray(r.Entries).map((e) => ({
      name: e.RefinementName,
      count: Number(e.RefinementCount) || 0
    }));
  });
  return result;
}

export class SearchApi {
  constructor(private client: SPClient) {}

  public async run(webUrl: string, o: ISearchOptions): Promise<ISearchResult> {
    const parts = [
      `querytext='${encodeURIComponent(odataString(o.query))}'`,
      `rowlimit=${o.rowLimit === undefined ? 1 : o.rowLimit}`,
      'trimduplicates=false'
    ];
    if (o.select && o.select.length) {
      parts.push(`selectproperties='${encodeURIComponent(o.select.join(','))}'`);
    }
    if (o.startRow) {
      parts.push(`startrow=${o.startRow}`);
    }
    if (o.sort) {
      parts.push(`sortlist='${encodeURIComponent(o.sort)}'`);
    }
    if (o.refiners) {
      parts.push(`refiners='${encodeURIComponent(o.refiners)}'`);
    }
    const url = `${trimSlash(webUrl)}/_api/search/query?${parts.join('&')}`;
    const json = await this.client.get<ISearchResponse>(url, { odata3: true });
    return parseSearch(entityOf<ISearchResponse>(json));
  }

  // ---- Sites inventory ------------------------------------------------------

  /**
   * Every site collection the signed-in user can see (search is security trimmed).
   *
   * Pages of 500 are read in DocId order with "IndexDocId>last" in the query, which is
   * Microsoft's recommended way to walk large result sets: unlike startrow paging it
   * has no 50,000-row ceiling and stays fast at the end of the set. One request per
   * 500 sites (17,000 sites = 35 requests), paced by SPClient. If DocId is not
   * returned (older farms), it falls back to startrow paging.
   */
  public async listSites(homeWebUrl: string, maxSites = 100000, onProgress?: (loaded: number) => void): Promise<{ sites: ISiteRow[]; truncated: boolean }> {
    const origin = (/^(https:\/\/[^/]+)/i.exec(homeWebUrl) || [''])[1].toLowerCase();
    const sites: ISiteRow[] = [];
    const seen: { [url: string]: boolean } = {};
    const select = ['Title', 'Path', 'WebTemplate', 'Created', 'LastModifiedTime', 'GroupId', 'DocId'];
    let lastDocId = 0;
    let useDocId = true;
    let startRow = 0;
    for (let guard = 0; guard < 400; guard++) {
      const res = await this.run(homeWebUrl, useDocId
        ? { query: `contentclass:STS_Site IndexDocId>${lastDocId}`, select, rowLimit: 500, sort: '[DocId]:ascending' }
        : { query: 'contentclass:STS_Site', select, rowLimit: 500, startRow, sort: 'LastModifiedTime:descending' });
      let maxDoc = lastDocId;
      res.rows.forEach((r) => {
        const doc = Number(r.DocId);
        if (isFinite(doc) && doc > maxDoc) {
          maxDoc = doc;
        }
        const url = trimSlash(r.Path || '');
        if (!url || seen[url.toLowerCase()]) {
          return;
        }
        seen[url.toLowerCase()] = true;
        sites.push({
          title: r.Title || url,
          url,
          template: r.WebTemplate || '',
          created: toDate(r.Created),
          lastModified: toDate(r.LastModifiedTime),
          groupConnected: !!r.GroupId && r.GroupId !== '00000000-0000-0000-0000-000000000000',
          manageable: url.toLowerCase().indexOf(origin) === 0
        });
      });
      if (onProgress) {
        onProgress(sites.length);
      }
      if (res.rows.length < 500) {
        return { sites, truncated: false };
      }
      if (sites.length >= maxSites) {
        return { sites, truncated: true };
      }
      if (useDocId && maxDoc === lastDocId) {
        // DocId not returned: switch to startrow paging (capped by search at 50,000 rows).
        useDocId = false;
        startRow = 0;
        sites.length = 0;
        Object.keys(seen).forEach((k) => delete seen[k]);
        continue;
      }
      lastDocId = maxDoc;
      startRow += 500;
      if (!useDocId && startRow >= 50000) {
        return { sites, truncated: true };
      }
    }
    return { sites, truncated: true };
  }

  // ---- Storage insights -----------------------------------------------------

  private docScope(siteUrl: string): string {
    return `IsDocument:1 Path:${kqlQuote(trimSlash(siteUrl) + '/*')}`;
  }

  /** File counts by extension for a whole site collection - one request. */
  public async fileTypes(siteUrl: string): Promise<{ types: IFileTypeCount[]; total: number; topTenOnly: boolean }> {
    const scope = this.docScope(siteUrl);
    let topTenOnly = false;
    let res: ISearchResult;
    try {
      res = await this.run(siteUrl, { query: scope, rowLimit: 1, select: ['Title'], refiners: 'FileType(filter=500/0/*)' });
    } catch (e) {
      if (!(e instanceof HttpError) || e.status === 401 || e.status === 403) {
        throw e;
      }
      // Some tenants reject the refiner options: fall back to the default top 10.
      topTenOnly = true;
      res = await this.run(siteUrl, { query: scope, rowLimit: 1, select: ['Title'], refiners: 'FileType' });
    }
    const entries = res.refiners.FileType || [];
    return {
      types: entries.map((e) => ({ type: e.name, count: e.count })).sort((a, b) => b.count - a.count),
      total: res.total,
      topTenOnly
    };
  }

  /** Largest files in a site collection (top N by size) - one request. */
  public async largestFiles(siteUrl: string, minBytes: number, top: number): Promise<IFileRow[]> {
    const res = await this.run(siteUrl, {
      query: `${this.docScope(siteUrl)} Size>=${Math.floor(minBytes)}`,
      select: ['Title', 'Path', 'Size', 'FileExtension', 'LastModifiedTime', 'Author'],
      rowLimit: top,
      sort: 'Size:descending'
    });
    return res.rows.map((r) => ({
      title: r.Title || lastSegment(r.Path),
      path: r.Path,
      size: Number(r.Size) || 0,
      extension: (r.FileExtension || '').toLowerCase(),
      modified: toDate(r.LastModifiedTime),
      author: r.Author || ''
    }));
  }

  /** How many indexed files have not been modified since each cut-off date. */
  public async staleCounts(siteUrl: string, years: number[]): Promise<Array<{ years: number; count: number }>> {
    const out: Array<{ years: number; count: number }> = [];
    for (const y of years) {
      const cutoff = new Date();
      cutoff.setFullYear(cutoff.getFullYear() - y);
      const res = await this.run(siteUrl, {
        query: `${this.docScope(siteUrl)} LastModifiedTime<${cutoff.toISOString().substring(0, 10)}`,
        select: ['Title'],
        rowLimit: 1
      });
      out.push({ years: y, count: res.total });
    }
    return out;
  }
}

export function toDate(value: string | undefined | null): Date | undefined {
  if (!value) {
    return undefined;
  }
  const d = new Date(value);
  return isNaN(d.getTime()) ? undefined : d;
}

function lastSegment(path: string): string {
  const parts = (path || '').split('/');
  return decodeURIComponent(parts[parts.length - 1] || path || '');
}
