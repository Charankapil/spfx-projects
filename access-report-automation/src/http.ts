import { Logger, silentLogger } from './log.js';

/**
 * One HTTP layer for everything (sign-in, Microsoft Graph, SharePoint):
 * a ceiling on requests in flight, a time limit per request, and patient retries.
 *
 * Throttling is handled the way Microsoft asks: on 429 / 503 the Retry-After header is
 * honoured, and every other request waits too (one shared pause), because a throttled
 * tenant stays throttled longer when the caller keeps knocking.
 */

export class HttpError extends Error {
  constructor(public status: number, message: string, public url: string, public body = '') {
    super(message);
    this.name = 'HttpError';
  }
}

export interface HttpRequest {
  method?: string;
  headers?: Record<string, string>;
  body?: string | Uint8Array;
  /** What to do with the response body. Default json. */
  as?: 'json' | 'bytes' | 'none';
  timeoutMs?: number;
  /**
   * true (default): repeat on any passing fault. 'throttled': repeat only when the server
   * refused the request outright (429 / 503), for calls that must not run twice.
   */
  retry?: boolean | 'throttled';
}

export interface HttpResponse<T = unknown> {
  status: number;
  headers: Headers;
  data: T;
}

export interface HttpOptions {
  concurrency?: number;
  maxRetries?: number;
  timeoutMs?: number;
  baseDelayMs?: number;
  maxDelayMs?: number;
  userAgent?: string;
  log?: Logger;
  fetchImpl?: typeof fetch;
}

export interface HttpStats {
  requests: number;
  retries: number;
  throttled: number;
}

const sleep = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));

/** Retry-After is seconds or an HTTP date. */
export function parseRetryAfter(value: string | null, now = Date.now()): number | undefined {
  if (!value) {
    return undefined;
  }
  const v = value.trim();
  if (/^\d+(\.\d+)?$/.test(v)) {
    return Math.round(parseFloat(v) * 1000);
  }
  const t = Date.parse(v);
  return isNaN(t) ? undefined : Math.max(0, t - now);
}

/** Pulls the human message out of the error shapes SharePoint, Graph and Entra ID use. */
export function errorMessageFrom(text: string, status: number): string {
  const fallback = text ? text.replace(/\s+/g, ' ').substring(0, 300) : `HTTP ${status}`;
  try {
    const j = JSON.parse(text) as Record<string, unknown>;
    const e = (j['odata.error'] || j.error) as Record<string, unknown> | string | undefined;
    if (typeof j.error_description === 'string') {
      return j.error_description.split(/\r?\n/)[0].substring(0, 400);
    }
    if (typeof e === 'string') {
      return e;
    }
    if (e && typeof e === 'object') {
      const m = e.message as { value?: string } | string | undefined;
      const msg = typeof m === 'string' ? m : m && m.value;
      if (msg) {
        return String(msg).substring(0, 400);
      }
      if (typeof e.code === 'string') {
        return e.code;
      }
    }
  } catch {
    /* not JSON */
  }
  return fallback;
}

export class Http {
  public stats: HttpStats = { requests: 0, retries: 0, throttled: 0 };
  private active = 0;
  private waiting: Array<() => void> = [];
  private pausedUntil = 0;
  private readonly concurrency: number;
  private readonly maxRetries: number;
  private readonly timeoutMs: number;
  private readonly baseDelayMs: number;
  private readonly maxDelayMs: number;
  private readonly userAgent: string;
  private readonly log: Logger;
  private readonly fetchImpl: typeof fetch;

  constructor(o: HttpOptions = {}) {
    this.concurrency = Math.max(1, o.concurrency ?? 2);
    this.maxRetries = o.maxRetries ?? 7;
    this.timeoutMs = o.timeoutMs ?? 120000;
    this.baseDelayMs = o.baseDelayMs ?? 1000;
    this.maxDelayMs = o.maxDelayMs ?? 300000;
    this.userAgent = o.userAgent ?? 'NONISV|SPFxProjects|AccessReportAutomation/1.0';
    this.log = o.log ?? silentLogger;
    this.fetchImpl = o.fetchImpl ?? fetch;
  }

  private async acquire(): Promise<void> {
    if (this.active < this.concurrency) {
      this.active++;
      return;
    }
    await new Promise<void>((resolve) => this.waiting.push(resolve));
  }

  private release(): void {
    const next = this.waiting.shift();
    if (next) {
      next();
    } else {
      this.active--;
    }
  }

  public async request<T = unknown>(url: string, req: HttpRequest = {}): Promise<HttpResponse<T>> {
    await this.acquire();
    try {
      return await this.attempt<T>(url, req);
    } finally {
      this.release();
    }
  }

  private async attempt<T>(url: string, req: HttpRequest): Promise<HttpResponse<T>> {
    const as = req.as ?? 'json';
    const mayRetry = req.retry === undefined || req.retry === true;
    const mayRetryThrottled = mayRetry || req.retry === 'throttled';
    let lastError: unknown;
    for (let tryNo = 0; ; tryNo++) {
      const wait = this.pausedUntil - Date.now();
      if (wait > 0) {
        await sleep(wait);
      }
      this.stats.requests++;
      let res: Response;
      try {
        res = await this.fetchImpl(url, {
          method: req.method ?? 'GET',
          headers: { 'User-Agent': this.userAgent, ...(req.headers ?? {}) },
          body: req.body as RequestInit['body'],
          signal: AbortSignal.timeout(req.timeoutMs ?? this.timeoutMs),
          redirect: 'manual'
        });
      } catch (e) {
        // Connection refused / reset / DNS / time limit: worth repeating.
        lastError = new HttpError(0, `Network error: ${(e as Error).message || String(e)}`, url);
        if (!mayRetry || tryNo >= this.maxRetries) {
          throw lastError;
        }
        this.stats.retries++;
        await sleep(this.backoff(tryNo));
        continue;
      }

      if (res.ok) {
        try {
          if (as === 'none') {
            await res.arrayBuffer();
            return { status: res.status, headers: res.headers, data: undefined as T };
          }
          if (as === 'bytes') {
            return { status: res.status, headers: res.headers, data: new Uint8Array(await res.arrayBuffer()) as T };
          }
          const text = await res.text();
          return { status: res.status, headers: res.headers, data: (text ? JSON.parse(text) : undefined) as T };
        } catch (e) {
          // The body was cut off or is not what was promised.
          lastError = new HttpError(0, `Unreadable response: ${(e as Error).message}`, url);
          if (!mayRetry || tryNo >= this.maxRetries) {
            throw lastError;
          }
          this.stats.retries++;
          await sleep(this.backoff(tryNo));
          continue;
        }
      }

      // Redirects are never followed: a bearer token must not travel to another address.
      const body = await res.text().catch(() => '');
      const status = res.status;
      const throttled = status === 429 || status === 503;
      // 500 is retried a little: SharePoint answers 500 for some passing faults, but also for real errors.
      const transient = throttled || status === 502 || status === 504 || status === 408 || (status === 500 && tryNo < 2);
      lastError = new HttpError(status, errorMessageFrom(body, status), url, body.substring(0, 2000));
      if (!transient || !(throttled ? mayRetryThrottled : mayRetry) || tryNo >= this.maxRetries) {
        throw lastError;
      }
      this.stats.retries++;
      let delay = this.backoff(tryNo);
      if (throttled) {
        this.stats.throttled++;
        const asked = parseRetryAfter(res.headers.get('retry-after'));
        delay = Math.min(this.maxDelayMs, Math.max(asked ?? 0, delay));
        // Everyone waits, not only this request.
        this.pausedUntil = Math.max(this.pausedUntil, Date.now() + delay);
        this.log.info(`Throttled (HTTP ${status}); pausing all requests for ${Math.ceil(delay / 1000)} s.`);
      }
      await sleep(delay);
    }
  }

  private backoff(tryNo: number): number {
    const exp = Math.min(this.maxDelayMs, this.baseDelayMs * Math.pow(2, tryNo));
    return Math.round(exp / 2 + Math.random() * (exp / 2));
  }
}
