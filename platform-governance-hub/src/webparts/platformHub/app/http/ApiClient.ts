/**
 * Thin, transport-agnostic HTTP client used by every service.
 *
 * - The SPFx web part plugs in a transport backed by AadHttpClient, so tokens
 *   for Graph, Power Platform and Dataverse are acquired by SharePoint.
 * - Handles throttling (429/503/504) using Retry-After with exponential
 *   backoff, follows OData paging links and classifies failures so the
 *   "Data sources" panel can explain exactly what is wrong.
 */

export interface IHttpResponseLike {
  status: number;
  ok: boolean;
  // eslint-disable-next-line @rushstack/no-new-null -- mirrors the DOM Headers API
  headers: { get(name: string): string | null };
  json(): Promise<unknown>;
  text(): Promise<string>;
}

export interface IHttpRequestInit {
  method: 'GET' | 'POST';
  headers?: Record<string, string>;
  body?: string;
}

export interface IHttpTransport {
  /** `resource` is the Entra ID resource (audience) the token is issued for. */
  fetch(resource: string, url: string, init: IHttpRequestInit): Promise<IHttpResponseLike>;
}

export type ApiErrorKind = 'forbidden' | 'blocked' | 'notFound' | 'throttled' | 'error';

export class ApiError extends Error {
  public readonly kind: ApiErrorKind;
  public readonly status: number;
  public readonly url: string;

  constructor(kind: ApiErrorKind, status: number, message: string, url: string) {
    super(message);
    this.kind = kind;
    this.status = status;
    this.url = url;
    // Restore prototype chain when targeting ES5.
    Object.setPrototypeOf(this, ApiError.prototype);
  }
}

export function isApiError(e: unknown): e is ApiError {
  return e instanceof ApiError;
}

/** Turns anything thrown into a short, user-facing explanation. */
export function describeError(e: unknown): string {
  if (isApiError(e)) {
    switch (e.kind) {
      case 'forbidden':
        return `Access denied (${e.status || 'consent'}). ${e.message}`;
      case 'blocked':
        return `The browser could not reach this API (network or CORS). ${e.message}`;
      case 'notFound':
        return `Not found (404). ${e.message}`;
      case 'throttled':
        return `Throttled by the service after several retries. ${e.message}`;
      default:
        return e.message;
    }
  }
  if (e instanceof Error) return e.message;
  return String(e);
}

const RETRY_STATUSES = [429, 502, 503, 504];
const MAX_RETRIES = 5;

function delay(ms: number): Promise<void> {
  return new Promise<void>((resolve) => setTimeout(resolve, ms));
}

function retryAfterMs(res: IHttpResponseLike, attempt: number): number {
  const header = res.headers.get('Retry-After');
  if (header) {
    const seconds = parseInt(header, 10);
    if (!isNaN(seconds)) return Math.min(seconds, 120) * 1000;
    const date = Date.parse(header);
    if (!isNaN(date)) return Math.max(0, Math.min(date - Date.now(), 120000));
  }
  // Exponential backoff with jitter: ~1s, 2s, 4s, 8s, 16s.
  return Math.pow(2, attempt) * 1000 + Math.floor(Math.random() * 400);
}

async function readErrorMessage(res: IHttpResponseLike): Promise<string> {
  try {
    const text = await res.text();
    try {
      const body = JSON.parse(text);
      const err = body && (body.error || body);
      if (err && typeof err === 'object') {
        const msg = err.message || (err.innererror && err.innererror.message) || err.code;
        if (msg) return String(msg).slice(0, 400);
      }
    } catch {
      /* not JSON */
    }
    return text ? text.slice(0, 400) : `HTTP ${res.status}`;
  } catch {
    return `HTTP ${res.status}`;
  }
}

function classifyThrown(e: unknown, url: string): ApiError {
  if (isApiError(e)) return e;
  const message = e instanceof Error ? e.message : String(e);
  // Token acquisition failures surface as AADSTS errors from the SPFx token provider.
  if (/AADSTS|consent|interaction_required|access_denied|not have permission|token/i.test(message)) {
    return new ApiError(
      'forbidden',
      0,
      `${message.slice(0, 300)} — approve the pending API access requests in SharePoint admin center > Advanced > API access.`,
      url
    );
  }
  // fetch() rejects with TypeError on network failure or CORS refusal.
  return new ApiError('blocked', 0, message.slice(0, 300), url);
}

export interface IPagedOptions {
  /** Stop after this many items (protects the browser on huge tenants). */
  maxItems?: number;
  headers?: Record<string, string>;
  /** Called after each page with the running total. */
  onPage?: (count: number) => void;
}

export class ApiClient {
  private readonly _transport: IHttpTransport;

  constructor(transport: IHttpTransport) {
    this._transport = transport;
  }

  public async request(resource: string, url: string, init: IHttpRequestInit): Promise<IHttpResponseLike> {
    for (let attempt = 0; ; attempt++) {
      let res: IHttpResponseLike;
      try {
        res = await this._transport.fetch(resource, url, init);
      } catch (e) {
        throw classifyThrown(e, url);
      }
      if (res.ok) return res;
      if (RETRY_STATUSES.indexOf(res.status) >= 0 && attempt < MAX_RETRIES) {
        await delay(retryAfterMs(res, attempt));
        continue;
      }
      const message = await readErrorMessage(res);
      if (res.status === 401 || res.status === 403) throw new ApiError('forbidden', res.status, message, url);
      if (res.status === 404) throw new ApiError('notFound', res.status, message, url);
      if (RETRY_STATUSES.indexOf(res.status) >= 0) throw new ApiError('throttled', res.status, message, url);
      throw new ApiError('error', res.status, message, url);
    }
  }

  public async getJson<T>(resource: string, url: string, headers?: Record<string, string>): Promise<T> {
    const res = await this.request(resource, url, {
      method: 'GET',
      headers: { Accept: 'application/json', ...(headers || {}) }
    });
    return (await res.json()) as T;
  }

  public async postJson<T>(resource: string, url: string, body: unknown, headers?: Record<string, string>): Promise<T> {
    const res = await this.request(resource, url, {
      method: 'POST',
      headers: { Accept: 'application/json', 'Content-Type': 'application/json', ...(headers || {}) },
      body: JSON.stringify(body)
    });
    const text = await res.text();
    return (text ? JSON.parse(text) : {}) as T;
  }

  /**
   * Follows `@odata.nextLink` (Graph, Dataverse) or `nextLink` (Power Platform)
   * until exhausted or `maxItems` is reached.
   */
  public async getAll<T>(resource: string, url: string, options: IPagedOptions = {}): Promise<T[]> {
    const items: T[] = [];
    let next: string | undefined = url;
    const seen: Record<string, boolean> = {};
    while (next) {
      if (seen[next]) break; // defensive: never loop on a repeated link
      seen[next] = true;
      const page: { value?: T[]; nextLink?: string; '@odata.nextLink'?: string } = await this.getJson(resource, next, options.headers);
      const value = page && Array.isArray(page.value) ? page.value : [];
      for (const v of value) items.push(v);
      if (options.onPage) options.onPage(items.length);
      if (options.maxItems && items.length >= options.maxItems) {
        items.length = options.maxItems;
        break;
      }
      next = page ? page['@odata.nextLink'] || page.nextLink : undefined;
    }
    return items;
  }
}

/** Runs async work over `items` with at most `limit` in flight. */
export async function mapWithConcurrency<T, R>(
  items: T[],
  limit: number,
  worker: (item: T, index: number) => Promise<R>
): Promise<R[]> {
  const results: R[] = new Array(items.length);
  let cursor = 0;
  const runners: Promise<void>[] = [];
  const count = Math.max(1, Math.min(limit, items.length));
  for (let r = 0; r < count; r++) {
    runners.push(
      (async () => {
        while (cursor < items.length) {
          const index = cursor++;
          results[index] = await worker(items[index], index);
        }
      })()
    );
  }
  await Promise.all(runners);
  return results;
}
