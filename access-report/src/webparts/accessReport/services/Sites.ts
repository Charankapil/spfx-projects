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
export const MAX_ROWS = 50000;
/** Same list, same order as SITE_ORDERS in flow/definition.ts (the tests check). */
export const SITE_ORDERS = ['Created:ascending', 'LastModifiedTime:ascending', ''];

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

  private async page(homeWebUrl: string, order: string, startRow: number, rowLimit = PAGE, select = 'Title,Path,WebTemplate'): Promise<Array<{ [key: string]: string }>> {
    const parts = [`querytext='${encodeURIComponent(odataString('contentclass:STS_Site'))}'`, `selectproperties='${encodeURIComponent(select)}'`, `rowlimit=${rowLimit}`];
    if (startRow) {
      parts.push(`startrow=${startRow}`);
    }
    parts.push('trimduplicates=false');
    if (order) {
      parts.push(`sortlist='${encodeURIComponent(order)}'`);
    }
    const json = await this.client.get<ISearchResponse>(`${trimSlash(homeWebUrl)}/_api/search/query?${parts.join('&')}`, { odata3: true, cacheTtlMs: 0 });
    return searchRows(entityOf<ISearchResponse>(json));
  }

  /** The first ordering of SITE_ORDERS this tenant's search accepts (one-row probes, as the flow does). */
  public async order(homeWebUrl: string): Promise<string> {
    const refusals: string[] = [];
    for (const o of SITE_ORDERS) {
      try {
        await this.page(homeWebUrl, o, 0, 1, 'Path');
        return o;
      } catch (e) {
        if (!(e instanceof HttpError) || e.status !== 400) {
          throw e;
        }
        refusals.push(`${o || 'no order'}: ${e.message}`);
      }
    }
    throw new Error(`SharePoint search refused every way of listing the sites. ${refusals.join('; ')}`);
  }

  /**
   * Pages of 500 of the plain site query by startrow (search serves up to 50,000 rows this
   * way; beyond that the list is marked truncated), in the first ordering the tenant accepts.
   */
  public async list(homeWebUrl: string, filter: string, onProgress?: (loaded: number) => void): Promise<{ sites: ISite[]; truncated: boolean }> {
    const origin = originOf(homeWebUrl);
    const order = await this.order(homeWebUrl);
    const sites: ISite[] = [];
    const seen: { [url: string]: boolean } = {};
    for (let startRow = 0; startRow < MAX_ROWS; startRow += PAGE) {
      const rows = await this.page(homeWebUrl, order, startRow);
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
