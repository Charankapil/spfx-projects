import { entityOf, HttpError, odataString, ODataCollection, SPClient, toArray, trimSlash } from '../core/SPClient';
import { ISite } from './AccessScan';

/**
 * The sites a report covers: every site collection the signed-in user can see (search is
 * security trimmed), minus OneDrive personal sites and anything on another host, optionally
 * narrowed by a text filter. The flow lists them the same way.
 */

interface ICell {
  Key: string;
  Value: string | null;
}
interface IPrimary {
  RelevantResults?: { Table?: { Rows?: ODataCollection<{ Cells: ODataCollection<ICell> }> } };
}
interface ISearchResponse {
  PrimaryQueryResult?: IPrimary;
  query?: { PrimaryQueryResult?: IPrimary };
  d?: { query?: { PrimaryQueryResult?: IPrimary }; PrimaryQueryResult?: IPrimary };
}

export const PAGE = 500;

export function searchRows(json: ISearchResponse): Array<{ [key: string]: string }> {
  const primary = (json.d && (json.d.query ? json.d.query.PrimaryQueryResult : json.d.PrimaryQueryResult)) || (json.query ? json.query.PrimaryQueryResult : undefined) || json.PrimaryQueryResult;
  const table = primary && primary.RelevantResults && primary.RelevantResults.Table;
  return toArray(table && table.Rows).map((row) => {
    const record: { [key: string]: string } = {};
    toArray(row.Cells).forEach((cell) => {
      record[cell.Key] = cell.Value === null || cell.Value === undefined ? '' : String(cell.Value);
    });
    return record;
  });
}

/** OneDrive (personal) site: a /personal/... address right under the host, or the SPSPERS template. */
export function isOneDriveSite(url: string, template?: string): boolean {
  return /^https:\/\/[^/]+\/personal\//i.test(url + '/') || /^SPSPERS/i.test(template || '');
}

export function originOf(url: string): string {
  return (/^(https:\/\/[^/]+)/i.exec(url) || [''])[1];
}

/** Is this search hit a site the report should cover? */
export function inScope(url: string, title: string, template: string, origin: string, filter: string): boolean {
  if (!url || (url + '/').toLowerCase().indexOf(origin.toLowerCase() + '/') !== 0 || isOneDriveSite(url, template)) {
    return false;
  }
  return (title + ' ' + url).toLowerCase().indexOf(filter.trim().toLowerCase()) >= 0;
}

export class SiteLister {
  constructor(private client: SPClient) {}

  private async page(homeWebUrl: string, query: string, sort: string, startRow: number): Promise<Array<{ [key: string]: string }>> {
    const parts = [
      `querytext='${encodeURIComponent(odataString(query))}'`,
      `selectproperties='${encodeURIComponent('Title,Path,WebTemplate,DocId')}'`,
      `rowlimit=${PAGE}`,
      'trimduplicates=false',
      `sortlist='${encodeURIComponent(sort)}'`
    ];
    if (startRow) {
      parts.push(`startrow=${startRow}`);
    }
    const json = await this.client.get<ISearchResponse>(`${trimSlash(homeWebUrl)}/_api/search/query?${parts.join('&')}`, { odata3: true, cacheTtlMs: 0 });
    return searchRows(entityOf<ISearchResponse>(json));
  }

  /**
   * Pages of 500 in DocId order with "IndexDocId>last" in the query - Microsoft's recommended
   * way to walk a large result set (no 50,000-row ceiling). If the tenant's search refuses
   * that or returns no DocId, it falls back to startrow paging.
   */
  public async list(homeWebUrl: string, filter: string, onProgress?: (loaded: number) => void): Promise<{ sites: ISite[]; truncated: boolean }> {
    const origin = originOf(homeWebUrl);
    const scope = 'contentclass:STS_Site';
    let sites: ISite[] = [];
    let seen: { [url: string]: boolean } = {};
    let lastDocId = 0;
    let useDocId = true;
    let startRow = 0;
    const restart = (): void => {
      useDocId = false;
      startRow = 0;
      sites = [];
      seen = {};
    };
    for (let guard = 0; guard < 400; guard++) {
      let rows: Array<{ [key: string]: string }>;
      try {
        rows = useDocId ? await this.page(homeWebUrl, `${scope} IndexDocId>${lastDocId}`, '[DocId]:ascending', 0) : await this.page(homeWebUrl, scope, 'LastModifiedTime:descending', startRow);
      } catch (e) {
        if (useDocId && e instanceof HttpError && e.status === 400) {
          restart();
          continue;
        }
        throw e;
      }
      let maxDoc = lastDocId;
      rows.forEach((r) => {
        const doc = Number(r.DocId);
        if (isFinite(doc) && doc > maxDoc) {
          maxDoc = doc;
        }
        const url = trimSlash(r.Path || '');
        const title = r.Title || url;
        if (seen[url.toLowerCase()] || !inScope(url, title, r.WebTemplate || '', origin, filter)) {
          return;
        }
        seen[url.toLowerCase()] = true;
        sites.push({ url, title });
      });
      if (onProgress) {
        onProgress(sites.length);
      }
      if (rows.length < PAGE) {
        return { sites, truncated: false };
      }
      if (useDocId && maxDoc === lastDocId) {
        restart();
        continue;
      }
      lastDocId = maxDoc;
      startRow += PAGE;
      if (!useDocId && startRow >= 50000) {
        return { sites, truncated: true };
      }
    }
    return { sites, truncated: true };
  }
}
