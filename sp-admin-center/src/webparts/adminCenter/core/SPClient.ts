import { ISPHttpClientOptions, ODataVersion, SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * The single door to SharePoint for the whole solution.
 *
 * Everything runs as the signed-in user through SPFx's SPHttpClient (cookie /
 * digest auth of the current page): no Azure AD app, no client id, no Graph.
 *
 * Throttling is treated as a design constraint, not an afterthought:
 *  - at most MAX_CONCURRENCY requests in flight, and request starts are spaced
 *    MIN_GAP_MS apart (~6 req/s peak), far below SharePoint's limits;
 *  - a 429 / 503 sets a *global* pause (Retry-After honoured, capped), so every
 *    queued request waits instead of piling on;
 *  - RateLimit-* response headers (when SharePoint sends them) slow the queue
 *    down before a 429 ever happens;
 *  - GET responses are cached and identical in-flight GETs are shared, so
 *    switching tabs never repeats a call;
 *  - nothing here polls: every request is caused by a user action.
 */

export const MAX_CONCURRENCY = 2;
export const MIN_GAP_MS = 150;
const MAX_RETRIES = 4;
const MAX_WAIT_MS = 120000;
const DEFAULT_CACHE_TTL_MS = 5 * 60 * 1000;

/** Search only speaks OData 3.0; the _api/web and _api/site calls are fine on v4. */
const V3_CONFIG: SPHttpClientConfiguration = SPHttpClient.configurations.v1.overrideWith({
  defaultODataVersion: ODataVersion.v3
});

export class CancelledError extends Error {
  constructor() {
    super('Cancelled');
    // Compiled to ES5, subclassing Error loses the prototype chain, which would break instanceof.
    Object.setPrototypeOf(this, CancelledError.prototype);
    this.name = 'CancelledError';
  }
}

export class HttpError extends Error {
  constructor(public status: number, message: string, public url: string) {
    super(message);
    Object.setPrototypeOf(this, HttpError.prototype);
    this.name = 'HttpError';
  }
}

export interface IClientStats {
  requests: number;
  retries: number;
  throttled: number;
  cacheHits: number;
  /** Epoch ms until which the queue is paused because of throttling (0 = running). */
  pausedUntil: number;
  queued: number;
}

export interface IRequestOptions {
  /** Use OData 3.0 (required for search and for change-log calls). */
  odata3?: boolean;
  /** Cache the response for this long (GET only). 0 disables. Default 5 min. */
  cacheTtlMs?: number;
  /** Extra request headers. */
  headers?: { [name: string]: string };
  /** HTTP verb override sent as X-HTTP-Method (MERGE / DELETE / PATCH). */
  method?: 'MERGE' | 'DELETE' | 'PATCH' | 'PUT';
  /** Send the body as-is (a string such as file content) instead of JSON-encoding it. */
  raw?: boolean;
}

interface ICacheEntry {
  at: number;
  value: unknown;
}

export class SPClient {
  private active = 0;
  private queue: Array<{ generation: number; run: () => void; reject: (e: Error) => void }> = [];
  private lastStart = 0;
  private pausedUntil = 0;
  private timer: ReturnType<typeof setTimeout> | undefined;
  private generation = 0;
  private cache: { [key: string]: ICacheEntry } = {};
  private inflight: { [key: string]: Promise<unknown> | undefined } = {};
  private listeners: Array<() => void> = [];

  public stats: IClientStats = { requests: 0, retries: 0, throttled: 0, cacheHits: 0, pausedUntil: 0, queued: 0 };

  constructor(private http: SPHttpClient, private homeOrigin: string) {}

  public subscribe(listener: () => void): () => void {
    this.listeners.push(listener);
    return () => {
      this.listeners = this.listeners.filter((l) => l !== listener);
    };
  }

  private notify(): void {
    this.stats.pausedUntil = this.pausedUntil > Date.now() ? this.pausedUntil : 0;
    this.stats.queued = this.queue.length + this.active;
    this.listeners.forEach((l) => l());
  }

  /** Drops queued (not yet started) requests, e.g. when the target site changes. */
  public cancelPending(): void {
    this.generation++;
    const dropped = this.queue;
    this.queue = [];
    dropped.forEach((j) => j.reject(new CancelledError()));
    this.notify();
  }

  public clearCache(): void {
    this.cache = {};
  }

  /** Only ever talk to the tenant the page is on. */
  public assertSameOrigin(url: string): void {
    const m = /^(https:\/\/[^/?#]+)/i.exec(url);
    if (!m || m[1].toLowerCase() !== this.homeOrigin.toLowerCase()) {
      throw new Error('Blocked request to a different origin: ' + url);
    }
  }

  // ---- scheduling ---------------------------------------------------------

  private enqueue<T>(task: () => Promise<T>): Promise<T> {
    return new Promise<T>((resolve, reject) => {
      const generation = this.generation;
      this.queue.push({
        generation,
        reject,
        run: () => {
          const done = (): void => {
            this.active--;
            this.notify();
            this.pump();
          };
          task().then(
            (v) => {
              done();
              resolve(v);
            },
            (e) => {
              done();
              reject(e);
            }
          );
        }
      });
      this.notify();
      this.pump();
    });
  }

  private pump(): void {
    if (this.timer) {
      return;
    }
    while (this.active < MAX_CONCURRENCY && this.queue.length > 0) {
      const now = Date.now();
      const wait = Math.max(this.pausedUntil - now, this.lastStart + MIN_GAP_MS - now, 0);
      if (wait > 0) {
        this.timer = setTimeout(() => {
          this.timer = undefined;
          this.pump();
        }, wait);
        return;
      }
      const job = this.queue.shift() as { generation: number; run: () => void };
      this.active++;
      this.lastStart = Date.now();
      job.run();
    }
  }

  private pauseFor(ms: number): void {
    const until = Date.now() + Math.min(ms, MAX_WAIT_MS);
    if (until > this.pausedUntil) {
      this.pausedUntil = until;
    }
    this.notify();
  }

  private sleep(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  private retryAfterMs(header: string | null, attempt: number): number {
    if (header) {
      const secs = parseInt(header, 10);
      if (!isNaN(secs)) {
        return Math.min(secs * 1000, MAX_WAIT_MS);
      }
      const date = Date.parse(header);
      if (!isNaN(date)) {
        return Math.min(Math.max(date - Date.now(), 1000), MAX_WAIT_MS);
      }
    }
    // No usable header: exponential backoff with jitter.
    return Math.min(2000 * Math.pow(2, attempt) + Math.floor(Math.random() * 1000), MAX_WAIT_MS);
  }

  private observeRateLimit(headers: Headers): void {
    const limit = parseInt(headers.get('RateLimit-Limit') || '', 10);
    const remaining = parseInt(headers.get('RateLimit-Remaining') || '', 10);
    const reset = parseInt(headers.get('RateLimit-Reset') || '', 10);
    if (!isNaN(limit) && !isNaN(remaining) && limit > 0 && remaining / limit < 0.1) {
      // Under 10% of the budget left: ease off until the window resets.
      this.pauseFor((isNaN(reset) ? 5 : Math.min(Math.max(reset, 1), 30)) * 1000);
    }
  }

  // ---- requests -----------------------------------------------------------

  private async execute(url: string, init: ISPHttpClientOptions, odata3: boolean, verb: string): Promise<SPHttpClientResponse> {
    let attempt = 0;
    for (;;) {
      this.stats.requests++;
      const config = odata3 ? V3_CONFIG : SPHttpClient.configurations.v1;
      const response =
        verb === 'GET'
          ? await this.http.get(url, config, init)
          : await this.http.post(url, config, init);
      this.observeRateLimit(response.headers);

      const transient = response.status === 429 || response.status === 503;
      const gateway = response.status === 502 || response.status === 504;
      if ((transient || gateway) && attempt < MAX_RETRIES) {
        const wait = this.retryAfterMs(response.headers.get('Retry-After'), attempt);
        this.stats.retries++;
        if (transient) {
          this.stats.throttled++;
          this.pauseFor(wait);
        }
        attempt++;
        await this.sleep(wait);
        continue;
      }
      return response;
    }
  }

  private async readError(response: SPHttpClientResponse, url: string): Promise<HttpError> {
    let detail = response.statusText || 'no details returned';
    try {
      const body = await response.text();
      if (body) {
        try {
          const parsed = JSON.parse(body);
          const m = parsed && (parsed.error ? parsed.error.message : parsed['odata.error'] && parsed['odata.error'].message);
          if (typeof m === 'string') {
            detail = m;
          } else if (m && typeof m.value === 'string') {
            detail = m.value;
          } else {
            detail = body.substring(0, 300);
          }
        } catch {
          detail = body.substring(0, 300);
        }
      }
    } catch {
      /* keep statusText */
    }
    // Long URLs swamp the UI; the console keeps the full detail for support.
    console.error(`[SharePoint Admin Center] ${response.status} ${url}: ${detail}`);
    return new HttpError(response.status, friendlyStatus(response.status, detail), url);
  }

  public async get<T>(url: string, options: IRequestOptions = {}): Promise<T> {
    this.assertSameOrigin(url);
    const ttl = options.cacheTtlMs === undefined ? DEFAULT_CACHE_TTL_MS : options.cacheTtlMs;
    const key = (options.odata3 ? 'v3|' : 'v4|') + url;

    if (ttl > 0) {
      const hit = this.cache[key];
      if (hit && Date.now() - hit.at < ttl) {
        this.stats.cacheHits++;
        this.notify();
        return hit.value as T;
      }
      const pending = this.inflight[key];
      if (pending) {
        return pending as Promise<T>;
      }
    }

    const promise = this.enqueue<T>(async () => {
      const init: ISPHttpClientOptions = options.headers ? { headers: options.headers } : {};
      const response = await this.execute(url, init, !!options.odata3, 'GET');
      if (!response.ok) {
        throw await this.readError(response, url);
      }
      return (await response.json()) as T;
    });

    if (ttl > 0) {
      this.inflight[key] = promise;
    }
    try {
      const value = await promise;
      if (ttl > 0) {
        this.cache[key] = { at: Date.now(), value };
      }
      return value;
    } finally {
      this.inflight[key] = undefined;
    }
  }

  /** POST (or MERGE / DELETE via X-HTTP-Method). Any write drops the read cache. */
  public async post<T = undefined>(url: string, body?: unknown, options: IRequestOptions = {}, webUrl?: string): Promise<T | undefined> {
    this.assertSameOrigin(url);
    const headers: { [name: string]: string } = { ...(options.headers || {}) };
    if (options.method) {
      headers['X-HTTP-Method'] = options.method;
      headers['IF-MATCH'] = '*';
    }
    const init: ISPHttpClientOptions = { headers, body: body === undefined ? undefined : options.raw ? String(body) : JSON.stringify(body) };
    if (webUrl) {
      this.assertSameOrigin(webUrl);
      (init as { webUrl?: string }).webUrl = webUrl;
    }
    const result = await this.enqueue<T | undefined>(async () => {
      const response = await this.execute(url, init, !!options.odata3, 'POST');
      if (!response.ok) {
        throw await this.readError(response, url);
      }
      if (response.status === 204) {
        return undefined;
      }
      const text = await response.text();
      return text ? (JSON.parse(text) as T) : undefined;
    });
    this.clearCache();
    return result;
  }
}

function friendlyStatus(status: number, detail: string): string {
  if (status === 401 || status === 403) {
    return `Access denied (${status}). You need higher permissions on this site for that. ${detail}`;
  }
  if (status === 404) {
    return `Not found (404). ${detail}`;
  }
  if (status === 429 || status === 503) {
    return 'SharePoint is throttling requests right now. Wait a minute and try again.';
  }
  return `${status}: ${detail}`;
}

// ---- response helpers (OData v3 verbose / v4 minimal / nometadata) ----------

export type ODataCollection<T> = T[] | { results?: T[]; value?: T[] } | undefined;

export function toArray<T>(c: ODataCollection<T>): T[] {
  if (!c) {
    return [];
  }
  if (Array.isArray(c)) {
    return c;
  }
  return c.results || c.value || [];
}

/** Rows of a collection response regardless of the OData flavour. */
export function rowsOf<T>(json: unknown): T[] {
  const j = json as { value?: T[]; d?: { results?: T[] } | T[]; results?: T[] } | T[] | undefined;
  if (!j) {
    return [];
  }
  if (Array.isArray(j)) {
    return j;
  }
  if (j.value) {
    return j.value;
  }
  if (j.d) {
    return Array.isArray(j.d) ? j.d : j.d.results || [];
  }
  return j.results || [];
}

/** The object of a single-entity response regardless of the OData flavour. */
export function entityOf<T>(json: unknown): T {
  const j = json as { d?: T };
  return (j && j.d ? j.d : json) as T;
}

export function nextLinkOf(json: unknown): string | undefined {
  const j = json as { [k: string]: unknown; d?: { __next?: string } };
  const link = j['odata.nextLink'] || j['@odata.nextLink'] || (j.d && j.d.__next);
  return typeof link === 'string' ? link : undefined;
}

/** Escapes a value for use inside an OData string literal ('...'). */
export function odataString(value: string): string {
  return value.replace(/'/g, "''");
}

export function trimSlash(url: string): string {
  return url.replace(/\/+$/, '');
}
