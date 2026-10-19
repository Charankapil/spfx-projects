import { HttpError } from './SPClient';

/**
 * Microsoft Graph access for the optional "Microsoft 365" features.
 *
 * Calls go through SPFx's AadHttpClient, so there is still no app registration
 * and no client id: tokens are issued to the tenant's built-in "SharePoint Online
 * Client Extensibility" principal, for the permissions a Global Administrator
 * approved in SharePoint admin center > Advanced > API access. Until that is
 * done every Graph feature reports "not approved" and the rest of the solution
 * keeps working unchanged.
 *
 * The same throttling rules as SPClient apply: 2 requests in flight, starts
 * spaced 150 ms apart, a global pause on 429/503 honouring Retry-After, GET
 * caching, and nothing runs unless the admin opens a Graph feature.
 */

export const GRAPH_ORIGIN = 'https://graph.microsoft.com';
const MAX_CONCURRENCY = 2;
const MIN_GAP_MS = 150;
const MAX_RETRIES = 4;
const MAX_WAIT_MS = 120000;
const DEFAULT_CACHE_TTL_MS = 5 * 60 * 1000;

export interface IGraphResponse {
  ok: boolean;
  status: number;
  statusText: string;
  headers: { get(name: string): string | null };
  text(): Promise<string>;
}

/** Sends one request with a Graph token. Supplied by the web part (AadHttpClient) or by tests. */
export type GraphFetcher = (url: string, method: string, headers: { [name: string]: string }, body?: string) => Promise<IGraphResponse>;

/** The permission is not approved (no token for it) or the signed-in user lacks the directory role. */
export class GraphAccessError extends HttpError {
  constructor(status: number, message: string, url: string, public notApproved: boolean) {
    super(status, message, url);
    Object.setPrototypeOf(this, GraphAccessError.prototype);
    this.name = 'GraphAccessError';
  }
}

export interface IGraphRequestOptions {
  cacheTtlMs?: number;
  headers?: { [name: string]: string };
}

/** Token errors (consent missing) come back as thrown errors, not HTTP responses. */
export function isConsentError(message: string): boolean {
  return /AADSTS65001|AADSTS90094|AADSTS500011|invalid_resource|consent|not been granted|has not been authorized|invalid_grant|interaction_required/i.test(message || '');
}

export class GraphClient {
  private active = 0;
  private queue: Array<() => void> = [];
  private lastStart = 0;
  private pausedUntil = 0;
  private timer: ReturnType<typeof setTimeout> | undefined;
  private cache: { [key: string]: { at: number; value: unknown } } = {};
  private inflight: { [key: string]: Promise<unknown> | undefined } = {};
  private fetcher: Promise<GraphFetcher> | undefined;

  public stats = { requests: 0, retries: 0, throttled: 0 };

  /** `connect` is only called on the first Graph request, so pages that never use Graph never ask for a token. */
  constructor(private connect: () => Promise<GraphFetcher>, private tokenScopes?: () => Promise<string[]>) {}

  /**
   * The delegated Graph scopes in the current token ("scp" claim), i.e. exactly what has
   * been approved. Costs no Graph request. Empty when nothing is approved.
   */
  public async grantedScopes(): Promise<string[]> {
    if (!this.tokenScopes) {
      return [];
    }
    try {
      return await this.tokenScopes();
    } catch (e) {
      if (isConsentError((e as Error).message)) {
        return [];
      }
      throw e;
    }
  }

  public clearCache(): void {
    this.cache = {};
  }

  private getFetcher(): Promise<GraphFetcher> {
    if (!this.fetcher) {
      this.fetcher = this.connect();
      // A failed connect is retried next time instead of being remembered.
      this.fetcher.catch(() => {
        this.fetcher = undefined;
      });
    }
    return this.fetcher;
  }

  /** Accepts "/v1.0/..." paths or absolute Graph URLs (e.g. @odata.nextLink); nothing else. */
  public resolve(pathOrUrl: string): string {
    const url = /^https:\/\//i.test(pathOrUrl) ? pathOrUrl : GRAPH_ORIGIN + (pathOrUrl.charAt(0) === '/' ? '' : '/') + pathOrUrl;
    if (url.toLowerCase().indexOf(GRAPH_ORIGIN + '/') !== 0) {
      throw new Error('Blocked request outside Microsoft Graph: ' + url);
    }
    return url;
  }

  private enqueue<T>(task: () => Promise<T>): Promise<T> {
    return new Promise<T>((resolve, reject) => {
      this.queue.push(() => {
        task().then(
          (v) => {
            this.active--;
            this.pump();
            resolve(v);
          },
          (e) => {
            this.active--;
            this.pump();
            reject(e);
          }
        );
      });
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
      const job = this.queue.shift() as () => void;
      this.active++;
      this.lastStart = Date.now();
      job();
    }
  }

  private retryAfterMs(header: string | null, attempt: number): number {
    const secs = parseInt(header || '', 10);
    if (!isNaN(secs)) {
      return Math.min(Math.max(secs, 1) * 1000, MAX_WAIT_MS);
    }
    return Math.min(2000 * Math.pow(2, attempt) + Math.floor(Math.random() * 1000), MAX_WAIT_MS);
  }

  private async execute(url: string, method: string, headers: { [name: string]: string }, body?: string): Promise<IGraphResponse> {
    let fetcher: GraphFetcher;
    try {
      fetcher = await this.getFetcher();
    } catch (e) {
      throw new GraphAccessError(401, notApprovedMessage((e as Error).message), url, true);
    }
    for (let attempt = 0; ; attempt++) {
      this.stats.requests++;
      let response: IGraphResponse;
      try {
        response = await fetcher(url, method, headers, body);
      } catch (e) {
        const message = (e as Error).message || String(e);
        if (isConsentError(message)) {
          throw new GraphAccessError(401, notApprovedMessage(message), url, true);
        }
        throw e;
      }
      const transient = response.status === 429 || response.status === 503 || response.status === 504;
      if (transient && attempt < MAX_RETRIES) {
        const wait = this.retryAfterMs(response.headers.get('Retry-After'), attempt);
        this.stats.retries++;
        if (response.status !== 504) {
          this.stats.throttled++;
          this.pausedUntil = Math.max(this.pausedUntil, Date.now() + wait);
        }
        await new Promise((r) => setTimeout(r, wait));
        continue;
      }
      return response;
    }
  }

  private async readError(response: IGraphResponse, url: string): Promise<HttpError> {
    let code = '';
    let detail = response.statusText || 'no details returned';
    try {
      const text = await response.text();
      const parsed = JSON.parse(text);
      if (parsed && parsed.error) {
        code = String(parsed.error.code || '');
        detail = String(parsed.error.message || detail);
      }
    } catch {
      /* keep statusText */
    }
    console.error(`[SharePoint Admin Center] Graph ${response.status} ${url}: ${code} ${detail}`);
    if (response.status === 401) {
      return new GraphAccessError(401, notApprovedMessage(detail), url, true);
    }
    if (response.status === 403) {
      return new GraphAccessError(403, `Microsoft 365 refused this (${code || 403}). Either the API permission it needs is not approved yet (see Settings > API access) or your admin role does not allow it. ${detail}`, url, false);
    }
    if (response.status === 429 || response.status === 503) {
      return new HttpError(response.status, 'Microsoft Graph is throttling requests right now. Wait a minute and try again.', url);
    }
    return new HttpError(response.status, `${response.status}${code ? ' ' + code : ''}: ${detail}`, url);
  }

  public async get<T>(path: string, options: IGraphRequestOptions = {}): Promise<T> {
    const text = await this.getText(path, options);
    return (text ? JSON.parse(text) : {}) as T;
  }

  public async getText(path: string, options: IGraphRequestOptions = {}): Promise<string> {
    const url = this.resolve(path);
    const ttl = options.cacheTtlMs === undefined ? DEFAULT_CACHE_TTL_MS : options.cacheTtlMs;
    const key = url + '|' + JSON.stringify(options.headers || {});
    if (ttl > 0) {
      const hit = this.cache[key];
      if (hit && Date.now() - hit.at < ttl) {
        return hit.value as string;
      }
      const pending = this.inflight[key];
      if (pending) {
        return pending as Promise<string>;
      }
    }
    const promise = this.enqueue<string>(async () => {
      const response = await this.execute(url, 'GET', { Accept: 'application/json', ...(options.headers || {}) });
      if (!response.ok) {
        throw await this.readError(response, url);
      }
      return response.text();
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

  /** Follows @odata.nextLink up to maxItems; reports progress after each page. */
  public async getAll<T>(path: string, maxItems: number, onProgress?: (n: number) => void, options: IGraphRequestOptions = {}): Promise<{ items: T[]; truncated: boolean }> {
    const items: T[] = [];
    let next: string | undefined = path;
    for (let guard = 0; next && guard < 1000; guard++) {
      const page: { value?: T[]; '@odata.nextLink'?: string } = await this.get(next, { ...options, cacheTtlMs: 0 });
      (page.value || []).forEach((v) => items.push(v));
      next = page['@odata.nextLink'];
      if (onProgress) {
        onProgress(items.length);
      }
      if (items.length >= maxItems) {
        return { items, truncated: !!next };
      }
    }
    return { items, truncated: !!next };
  }

  /** POST / PATCH / DELETE with a JSON body. Any write drops the read cache. */
  public async send<T = undefined>(method: 'POST' | 'PATCH' | 'DELETE', path: string, body?: unknown): Promise<T | undefined> {
    const url = this.resolve(path);
    const result = await this.enqueue<T | undefined>(async () => {
      const response = await this.execute(url, method, { Accept: 'application/json', 'Content-Type': 'application/json' }, body === undefined ? undefined : JSON.stringify(body));
      if (!response.ok) {
        throw await this.readError(response, url);
      }
      const text = response.status === 204 ? '' : await response.text();
      return text ? (JSON.parse(text) as T) : undefined;
    });
    this.clearCache();
    return result;
  }
}

function notApprovedMessage(detail: string): string {
  return 'This feature needs Microsoft Graph API access that has not been approved for this tenant yet. ' +
    'A Global Administrator approves it in SharePoint admin center > Advanced > API access (see Settings > API access). ' +
    (detail ? `Details: ${detail.substring(0, 200)}` : '');
}

/** Reads the "scp" claim of a JWT access token (no validation: display only). */
export function scopesFromToken(token: string): string[] {
  const part = (token || '').split('.')[1];
  if (!part) {
    return [];
  }
  try {
    const b64 = part.replace(/-/g, '+').replace(/_/g, '/');
    const padded = b64 + '==='.substring(0, (4 - (b64.length % 4)) % 4);
    const raw = atob(padded);
    const claims = JSON.parse(raw) as { scp?: string };
    return (claims.scp || '').split(' ').filter((x) => !!x);
  } catch {
    return [];
  }
}
