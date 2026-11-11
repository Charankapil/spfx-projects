import { ODataVersion, SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse } from '@microsoft/sp-http';

import { buildBatchBody, buildBatchGetBody, IBatchResult, IBatchUpdate, newBatchId, parseBatchResponse } from './batch';

import { CancelToken, enqueue, toolkitError } from './RequestQueue';

/**
 * The search REST endpoint only speaks OData 3.0 and answers a v4 request
 * with HTTP 500 once it has rows to serialize, so search is downgraded.
 */
const SEARCH_CONFIG: SPHttpClientConfiguration = SPHttpClient.configurations.v1.overrideWith({
  defaultODataVersion: ODataVersion.v3
});

export interface IRequestOptions {
  token?: CancelToken;
  /** Use the OData 3 configuration (search endpoint). */
  search?: boolean;
  headers?: Record<string, string>;
}

/** The minimum of an SPFx component context that the toolkit needs. */
export interface IToolkitContext {
  spHttpClient: SPHttpClient;
  pageContext: {
    web: { absoluteUrl: string; serverRelativeUrl: string };
    site: { absoluteUrl: string; serverRelativeUrl: string };
  };
}

/** SharePoint puts the useful part of a failure in the response body. */
export async function describeError(response: SPHttpClientResponse): Promise<string> {
  try {
    const body = await response.text();
    if (!body) {
      return response.statusText || 'no error details returned';
    }
    try {
      const parsed = JSON.parse(body);
      const message = parsed && (parsed.error || parsed['odata.error']);
      const text = message && message.message;
      if (typeof text === 'string') {
        return text;
      }
      if (text && typeof text.value === 'string') {
        return text.value;
      }
    } catch {
      // Body was not JSON - surface the raw text instead.
    }
    return body.substring(0, 300);
  } catch {
    return response.statusText || 'no error details returned';
  }
}

/** Escapes a value for use inside single quotes in a REST URL, e.g. getbytitle('...'). */
export function quote(value: string): string {
  return encodeURIComponent(value.replace(/'/g, "''"));
}

/**
 * All SharePoint traffic for a component. Every call is routed through the
 * page-wide request queue, runs as the signed-in user, and never leaves
 * SharePoint.
 */
export class SpClient {
  public readonly webUrl: string;
  public readonly siteUrl: string;
  public readonly webServerRelativeUrl: string;
  public readonly siteServerRelativeUrl: string;
  public readonly origin: string;

  constructor(private http: SPHttpClient, pageContext: IToolkitContext['pageContext']) {
    this.webUrl = pageContext.web.absoluteUrl.replace(/\/$/, '');
    this.siteUrl = pageContext.site.absoluteUrl.replace(/\/$/, '');
    this.webServerRelativeUrl = pageContext.web.serverRelativeUrl.replace(/\/$/, '');
    this.siteServerRelativeUrl = pageContext.site.serverRelativeUrl.replace(/\/$/, '');
    const match = /^(https?:\/\/[^/]+)/i.exec(this.webUrl);
    this.origin = match ? match[1] : '';
  }

  public static fromContext(context: IToolkitContext): SpClient {
    return new SpClient(context.spHttpClient, context.pageContext);
  }

  private async fail(response: SPHttpClientResponse): Promise<never> {
    const detail = await describeError(response);
    throw toolkitError(detail, 'http', response.status);
  }

  private async parse<T>(response: SPHttpClientResponse): Promise<T | undefined> {
    if (response.status === 204) {
      return undefined;
    }
    const text = await response.text();
    if (!text) {
      return undefined;
    }
    try {
      return JSON.parse(text) as T;
    } catch {
      return undefined;
    }
  }

  public async getJson<T>(url: string, options: IRequestOptions = {}): Promise<T> {
    const configuration = options.search ? SEARCH_CONFIG : SPHttpClient.configurations.v1;
    const response = await enqueue(
      () => this.http.get(url, configuration, options.headers ? { headers: options.headers } : undefined),
      options.token
    );
    if (!response.ok) {
      return this.fail(response);
    }
    const json = await this.parse<T>(response);
    if (json === undefined) {
      throw toolkitError('SharePoint returned an empty response.', 'http', response.status);
    }
    return json;
  }

  /** Raw text of a file or endpoint; undefined on 404. */
  public async getText(url: string, options: IRequestOptions = {}): Promise<string | undefined> {
    const response = await enqueue(() => this.http.get(url, SPHttpClient.configurations.v1), options.token);
    if (response.status === 404) {
      return undefined;
    }
    if (!response.ok) {
      return this.fail(response);
    }
    return response.text();
  }

  public async post<T>(url: string, body?: unknown, options: IRequestOptions = {}): Promise<T | undefined> {
    const response = await enqueue(
      () =>
        this.http.post(url, SPHttpClient.configurations.v1, {
          headers: options.headers,
          body: body === undefined ? undefined : typeof body === 'string' ? body : JSON.stringify(body)
        }),
      options.token
    );
    if (!response.ok) {
      return this.fail(response);
    }
    return this.parse<T>(response);
  }

  /** Updates an entity in place (list item, list, field). */
  public async merge(url: string, body: unknown, options: IRequestOptions = {}): Promise<void> {
    await this.post(url, body, {
      token: options.token,
      headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' }
    });
  }

  /** Permanently removes an entity that has no recycle bin, e.g. a custom action. */
  public async remove(url: string): Promise<void> {
    await this.post(url, undefined, { headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE' } });
  }

  /** One page of a collection plus the link to the next page, for "Load more". */
  public async getPage<T>(url: string, options: IRequestOptions = {}): Promise<{ items: T[]; next: string | undefined }> {
    const page = await this.getJson<Record<string, unknown>>(url, options);
    return {
      items: (page.value as T[]) || [],
      next: (page['@odata.nextLink'] || page['odata.nextLink']) as string | undefined
    };
  }

  /**
   * Applies many item updates in one HTTP call. If SharePoint throttles the
   * call, or any update inside it, the whole batch is sent again after the
   * wait; that is safe because writing the same values twice changes nothing.
   */
  public batchMerge(updates: IBatchUpdate[], token?: CancelToken): Promise<IBatchResult[]> {
    return this.sendBatch(updates.length, (id) => buildBatchBody(updates, id), token);
  }

  /**
   * Reads many addresses in one HTTP call (at most 100). Each result carries
   * its own status and body. Throttling is handled as for batchMerge.
   */
  public batchGet(urls: string[], token?: CancelToken): Promise<IBatchResult[]> {
    return this.sendBatch(urls.length, (id) => buildBatchGetBody(urls, id), token);
  }

  private async sendBatch(size: number, build: (batchId: string) => string, token?: CancelToken): Promise<IBatchResult[]> {
    if (size === 0) {
      return [];
    }
    let results: IBatchResult[] = [];
    const outer = await enqueue(
      async () => {
        const batchId = newBatchId();
        const response = await this.http.post(`${this.webUrl}/_api/$batch`, SPHttpClient.configurations.v1, {
          headers: { 'Content-Type': `multipart/mixed; boundary="batch_${batchId}"` },
          body: build(batchId)
        });
        if (!response.ok) {
          // Includes 429 and 503 on the batch itself: the queue waits and repeats it.
          return response;
        }
        results = parseBatchResponse(await response.text());
        const throttled = results.filter((r) => r.status === 429 || r.status === 503)[0];
        if (throttled) {
          const wait = throttled.retryAfter;
          return {
            ok: false,
            status: throttled.status,
            headers: { get: (name: string): string | null => (wait && /^retry-after$/i.test(name) ? String(wait) : null) }
          };
        }
        return { ok: true, status: 200, headers: { get: (): string | null => null } };
      },
      token,
      size
    );
    if (!outer.ok) {
      throw toolkitError(`The batch could not be sent (SharePoint answered ${outer.status}).`, 'http', outer.status);
    }
    if (results.length !== size) {
      throw toolkitError('SharePoint returned an incomplete batch answer, so the results cannot be trusted.', 'http');
    }
    return results;
  }

  /**
   * Reads a collection page by page, following SharePoint's next link.
   * Stops after `maxItems` so one call can never turn into an unbounded crawl.
   */
  public async getAll<T>(url: string, maxItems: number, options: IRequestOptions = {}): Promise<{ items: T[]; more: boolean }> {
    const items: T[] = [];
    let next: string | undefined = url;
    while (next && items.length < maxItems) {
      const page: Record<string, unknown> = await this.getJson<Record<string, unknown>>(next, options);
      const values = (page.value as T[]) || [];
      for (const value of values) {
        items.push(value);
      }
      next = (page['@odata.nextLink'] || page['odata.nextLink']) as string | undefined;
    }
    return { items: items.slice(0, maxItems), more: !!next || items.length > maxItems };
  }

  /**
   * Server-relative path of an absolute URL on this tenant. Search results
   * and list items already hold decoded paths; only links taken from page
   * HTML are percent-encoded and need `decode`.
   */
  public toServerRelative(absoluteOrRelativeUrl: string, decode = false): string {
    let path = absoluteOrRelativeUrl;
    if (this.origin && path.toLowerCase().indexOf(this.origin.toLowerCase()) === 0) {
      path = path.substring(this.origin.length);
    }
    if (!decode) {
      return path;
    }
    try {
      return decodeURIComponent(path);
    } catch {
      return path;
    }
  }

  /** `GetFileByServerRelativePath(decodedurl='...')` for a decoded server-relative path. */
  public fileApi(serverRelativePath: string): string {
    return `${this.webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${quote(serverRelativePath)}')`;
  }

  public folderApi(serverRelativePath: string): string {
    return `${this.webUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${quote(serverRelativePath)}')`;
  }
}
