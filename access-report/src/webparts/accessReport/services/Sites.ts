import { entityOf, odataString, ODataCollection, SPClient, toArray, trimSlash } from '../core/SPClient';
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
export const MAX_ROWS = 50000;

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

  private async page(homeWebUrl: string, startRow: number): Promise<Array<{ [key: string]: string }>> {
    const parts = [
      `querytext='${encodeURIComponent(odataString('contentclass:STS_Site'))}'`,
      `selectproperties='${encodeURIComponent('Title,Path,WebTemplate')}'`,
      `rowlimit=${PAGE}`,
      `startrow=${startRow}`,
      'trimduplicates=false',
      `sortlist='${encodeURIComponent('[DocId]:ascending')}'`
    ];
    const json = await this.client.get<ISearchResponse>(`${trimSlash(homeWebUrl)}/_api/search/query?${parts.join('&')}`, { odata3: true, cacheTtlMs: 0 });
    return searchRows(entityOf<ISearchResponse>(json));
  }

  /**
   * Pages of 500 of the plain site query, in a fixed (DocId) order, by startrow. Some tenants
   * refuse the "IndexDocId>n" paging trick ("We didn't understand your search terms"), so it
   * is not used; startrow serves up to 50,000 rows, beyond which the list is marked truncated.
   */
  public async list(homeWebUrl: string, filter: string, onProgress?: (loaded: number) => void): Promise<{ sites: ISite[]; truncated: boolean }> {
    const origin = originOf(homeWebUrl);
    const sites: ISite[] = [];
    const seen: { [url: string]: boolean } = {};
    for (let startRow = 0; startRow < MAX_ROWS; startRow += PAGE) {
      const rows = await this.page(homeWebUrl, startRow);
      rows.forEach((r) => {
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
    }
    return { sites, truncated: true };
  }
}
