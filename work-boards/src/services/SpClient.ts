import type { SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

/* eslint-disable @typescript-eslint/no-explicit-any */

export class SpError extends Error {
  public readonly status: number;
  public readonly code: string;
  /** The failing request, for example "GET /_api/web/lists(...)". Shown in setup errors to help diagnose problems. */
  public readonly request: string;

  constructor(message: string, status: number, code: string, request: string = '') {
    super(message);
    this.status = status;
    this.code = code;
    this.request = request;
    // Keep instanceof working when compiled to ES5.
    Object.setPrototypeOf(this, SpError.prototype);
  }

  public get isThrottled(): boolean {
    return this.status === 429 || this.status === 503;
  }

  public get isConflict(): boolean {
    return this.status === 412;
  }

  /** SharePoint answers a missing list, field or item with 404, but some endpoints use other statuses with a "does not exist" message. */
  public get isNotFound(): boolean {
    return this.status === 404 || /does not exist/i.test(this.message);
  }

  public get isAccessDenied(): boolean {
    return this.status === 403 || this.status === 401;
  }

  /** The 5,000-item list view threshold. */
  public get isThreshold(): boolean {
    return this.code.indexOf('SPQueryThrottledException') >= 0 || /threshold/i.test(this.message);
  }
}

const READ_HEADERS = { Accept: 'application/json;odata=minimalmetadata', 'odata-version': '' };
const WRITE_HEADERS = {
  Accept: 'application/json;odata=minimalmetadata',
  'Content-Type': 'application/json;odata=verbose',
  'odata-version': ''
};

const MAX_ATTEMPTS = 6;

function wait(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Thin wrapper over SPFx's SPHttpClient. Every call runs as the signed-in user
 * against the current site's /_api. No tokens, no app registration.
 * Retries throttled requests, honouring Retry-After.
 */
export class SpClient {
  public readonly webUrl: string;
  private readonly http: SPHttpClient;
  private readonly config: SPHttpClientConfiguration;

  /** `config` is SPHttpClient.configurations.v1, passed in so this module has no runtime SPFx import. */
  constructor(http: SPHttpClient, config: SPHttpClientConfiguration, webUrl: string) {
    this.http = http;
    this.config = config;
    this.webUrl = webUrl.replace(/\/$/, '');
  }

  public api(path: string): string {
    return this.webUrl + '/_api/' + path.replace(/^\//, '');
  }

  public async get<T = any>(path: string): Promise<T> {
    const res = await this.send('GET', this.api(path), { headers: READ_HEADERS });
    return (await res.json()) as T;
  }

  /** GET that follows odata.nextLink until all rows are read. */
  public async getAll<T = any>(path: string, maxRows: number = 100000): Promise<T[]> {
    let url: string | null = this.api(path);
    const out: T[] = [];
    while (url && out.length < maxRows) {
      const res = await this.send('GET', url, { headers: READ_HEADERS });
      const json: any = await res.json();
      (json.value || []).forEach((r: T) => out.push(r));
      url = json['odata.nextLink'] || json['@odata.nextLink'] || null;
    }
    return out;
  }

  public async post<T = any>(path: string, body?: any, extraHeaders?: { [k: string]: string }): Promise<T> {
    const res = await this.send('POST', this.api(path), {
      headers: { ...WRITE_HEADERS, ...(extraHeaders || {}) },
      body: body === undefined ? undefined : JSON.stringify(body)
    });
    return this.readOptionalJson<T>(res);
  }

  /** MERGE (partial update). Returns the new ETag when SharePoint sends one. */
  public async merge(path: string, body: any, etag: string = '*'): Promise<string> {
    const res = await this.send('POST', this.api(path), {
      headers: { ...WRITE_HEADERS, 'IF-MATCH': etag, 'X-HTTP-Method': 'MERGE' },
      body: JSON.stringify(body)
    });
    return res.headers.get('ETag') || '';
  }

  public async remove(path: string, etag: string = '*'): Promise<void> {
    await this.send('POST', this.api(path), {
      headers: { ...WRITE_HEADERS, 'IF-MATCH': etag, 'X-HTTP-Method': 'DELETE' }
    });
  }

  public async postBinary<T = any>(path: string, content: ArrayBuffer): Promise<T> {
    const res = await this.send('POST', this.api(path), {
      // A Content-Type is required: with no OData version set, SPHttpClient refuses any write that lacks one.
      headers: { Accept: 'application/json;odata=minimalmetadata', 'Content-Type': 'application/octet-stream', 'odata-version': '' },
      body: content
    });
    return this.readOptionalJson<T>(res);
  }

  private async readOptionalJson<T>(res: SPHttpClientResponse): Promise<T> {
    const text = await res.text();
    return (text ? JSON.parse(text) : undefined) as T;
  }

  private async send(method: 'GET' | 'POST', url: string, options: ISPHttpClientOptions): Promise<SPHttpClientResponse> {
    let attempt = 0;
    for (;;) {
      attempt++;
      const res: SPHttpClientResponse = method === 'GET'
        ? await this.http.get(url, this.config, options)
        : await this.http.post(url, this.config, options);
      if (res.ok) {
        return res;
      }
      if ((res.status === 429 || res.status === 503) && attempt < MAX_ATTEMPTS) {
        const retryAfter = parseInt(res.headers.get('Retry-After') || '', 10);
        const delay = !isNaN(retryAfter) ? retryAfter * 1000 : Math.min(30000, 500 * Math.pow(2, attempt));
        await wait(delay);
        continue;
      }
      throw await SpClient.toError(res, method, url, this.webUrl);
    }
  }

  private static async toError(res: SPHttpClientResponse, method: string, url: string, webUrl: string): Promise<SpError> {
    let message = res.statusText || 'Request failed';
    let code = '';
    try {
      const json: any = await res.json();
      const err = json['odata.error'] || json.error || {};
      code = err.code || '';
      message = (err.message && (err.message.value || err.message)) || message;
    } catch {
      // Body was not JSON.
    }
    return new SpError(String(message), res.status, String(code), `${method} ${url.replace(webUrl, '')}`);
  }
}

/** Quote a string for use inside an OData URL literal: O'Brien -> 'O''Brien'. */
export function odataString(value: string): string {
  return "'" + encodeURIComponent(value.replace(/'/g, "''")) + "'";
}

/** Run async tasks with a concurrency limit, keeping result order. */
export async function mapLimit<T, R>(items: T[], limit: number, fn: (item: T, index: number) => Promise<R>): Promise<R[]> {
  const results: R[] = new Array(items.length);
  let next = 0;
  const workers: Promise<void>[] = [];
  const worker = async (): Promise<void> => {
    while (next < items.length) {
      const index = next++;
      results[index] = await fn(items[index], index);
    }
  };
  for (let i = 0; i < Math.min(limit, items.length); i++) {
    workers.push(worker());
  }
  await Promise.all(workers);
  return results;
}
