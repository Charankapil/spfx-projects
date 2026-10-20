import {
  ODataVersion,
  SPHttpClient,
  SPHttpClientConfiguration,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import { describeError } from './httpErrors';

/**
 * OData 3 ("verbose") wraps every collection in a { results: [] } envelope
 * while OData 4 / nometadata returns a bare array, and which one comes back
 * depends on the negotiated Accept header. Both shapes are accepted.
 */
export type ODataCollection<T> = T[] | { results?: T[] } | undefined;

export function toArray<T>(collection: ODataCollection<T>): T[] {
  if (!collection) {
    return [];
  }
  if (Array.isArray(collection)) {
    return collection;
  }
  return collection.results || [];
}

/**
 * SPHttpClient.configurations.v1 sends "OData-Version: 4.0", but the search
 * REST endpoint only speaks OData 3.0 and answers a v4 request with HTTP 500
 * once it has rows to serialize. Only search is downgraded.
 */
export const SEARCH_CONFIG: SPHttpClientConfiguration = SPHttpClient.configurations.v1.overrideWith({
  defaultODataVersion: ODataVersion.v3
});

const MAX_RETRIES = 3;

interface ISearchCell {
  Key: string;
  Value: string | null;
}

interface ISearchRow {
  Cells: ODataCollection<ISearchCell>;
}

interface IPrimaryQueryResult {
  RelevantResults?: {
    TotalRows?: number | string;
    RowCount?: number | string;
    Table?: { Rows?: ODataCollection<ISearchRow> };
  };
}

interface ISearchQueryResponse {
  PrimaryQueryResult?: IPrimaryQueryResult;
  d?: {
    query?: { PrimaryQueryResult?: IPrimaryQueryResult };
    PrimaryQueryResult?: IPrimaryQueryResult;
  };
}

export interface ISearchPage {
  rows: { [managedProperty: string]: string }[];
  totalRows: number;
}

export interface ISearchOptions {
  selectProperties: string[];
  rowLimit: number;
  startRow?: number;
  sortList?: string;
}

/**
 * Thin wrapper over the ambient SPHttpClient: the current user's SharePoint
 * session does the authentication, so nothing needs an AAD app registration,
 * and every result is already trimmed to what that user may see.
 */
export class SpRestClient {
  constructor(private httpClient: SPHttpClient, private logPrefix: string) {}

  private async delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  private retryDelayMs(response: SPHttpClientResponse, attempt: number): number {
    const header = response.headers.get('Retry-After');
    const seconds = header ? parseInt(header, 10) : 2 * (attempt + 1);
    return (isNaN(seconds) ? 2 : seconds) * 1000;
  }

  private async handle<T>(
    url: string,
    send: () => Promise<SPHttpClientResponse>,
    attempt: number
  ): Promise<T> {
    const response = await send();
    if ((response.status === 429 || response.status === 503) && attempt < MAX_RETRIES) {
      await this.delay(this.retryDelayMs(response, attempt));
      return this.handle<T>(url, send, attempt + 1);
    }
    if (!response.ok) {
      const detail = await describeError(response);
      console.error(`[${this.logPrefix}] ${response.status} from ${url}: ${detail}`);
      throw new HttpError(response.status, `${response.status}: ${detail}`);
    }
    // MERGE / DELETE and some actions answer 204 with no body.
    const text = await response.text();
    return (text ? JSON.parse(text) : undefined) as T;
  }

  public getJson<T>(url: string, configuration: SPHttpClientConfiguration = SPHttpClient.configurations.v1): Promise<T> {
    return this.handle<T>(url, () => this.httpClient.get(url, configuration), 0);
  }

  /** POSTs JSON without __metadata: SharePoint Online accepts OData 4 payloads untyped. */
  public postJson<T>(url: string, body?: unknown, headers: { [name: string]: string } = {}): Promise<T> {
    return this.handle<T>(
      url,
      () =>
        this.httpClient.post(url, SPHttpClient.configurations.v1, {
          headers: { 'Content-Type': 'application/json;charset=utf-8', ...headers },
          body: body === undefined ? undefined : JSON.stringify(body)
        }),
      0
    );
  }

  /** Uploads raw bytes (file content) to an endpoint such as Files/AddUsingPath. */
  public postBinary<T>(url: string, content: ArrayBuffer): Promise<T> {
    return this.handle<T>(
      url,
      () =>
        this.httpClient.post(url, SPHttpClient.configurations.v1, {
          headers: { 'Content-Type': 'application/octet-stream' },
          body: content
        }),
      0
    );
  }

  /**
   * One page of search results, flattened to { managedProperty: value } rows.
   * Uses GET /_api/search/query, which the File Type Analyser has proven
   * against large tenants.
   */
  public async search(siteUrl: string, kql: string, options: ISearchOptions): Promise<ISearchPage> {
    let url =
      `${siteUrl}/_api/search/query` +
      `?querytext=${encodeURIComponent(`'${kql.replace(/'/g, "''")}'`)}` +
      `&selectproperties=${encodeURIComponent(`'${options.selectProperties.join(',')}'`)}` +
      `&rowlimit=${options.rowLimit}` +
      `&startrow=${options.startRow || 0}` +
      `&trimduplicates=false` +
      `&clienttype='ContentSearchRegular'`;
    if (options.sortList) {
      url += `&sortlist=${encodeURIComponent(`'${options.sortList}'`)}`;
    }

    const json = await this.getJson<ISearchQueryResponse>(url, SEARCH_CONFIG);
    const primary = json.PrimaryQueryResult || json.d?.query?.PrimaryQueryResult || json.d?.PrimaryQueryResult;
    const relevant = primary?.RelevantResults;
    const rows = toArray(relevant?.Table?.Rows).map((row) => {
      const flat: { [key: string]: string } = {};
      toArray(row.Cells).forEach((cell) => {
        flat[cell.Key] = cell.Value === null || cell.Value === undefined ? '' : String(cell.Value);
      });
      return flat;
    });
    const totalRows = parseInt(String(relevant?.TotalRows ?? rows.length), 10) || 0;
    return { rows, totalRows };
  }
}

export class HttpError extends Error {
  constructor(public status: number, message: string) {
    // ES5 output breaks instanceof for Error subclasses, so callers check .status instead.
    super(message);
  }
}

/** Lower-cased, no trailing slash, absolute — so URLs from lists, search and pageContext compare equal. */
export function normalizeUrl(url: string, origin: string): string {
  let absolute = (url || '').trim();
  if (absolute.indexOf('/') === 0) {
    absolute = origin + absolute;
  }
  return absolute.replace(/[?#].*$/, '').replace(/\/+$/, '').toLowerCase();
}
