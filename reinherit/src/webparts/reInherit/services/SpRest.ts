import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { buildBatchBody, IBatchRequest, IBatchResponse, parseBatchResponse } from './batch';
import { describeError } from './httpErrors';

/**
 * All calls go through the ambient SPHttpClient, which reuses the signed-in
 * user's SharePoint session and fetches / refreshes the form digest for
 * POSTs by itself. No AAD app registration, client id or secret is needed.
 */

const MAX_THROTTLE_RETRIES = 10;
const MAX_NETWORK_RETRIES = 3;

/**
 * Every request starts at least `baseGap` ms after the previous one, across
 * all parallel readers. When SharePoint throttles, the gap doubles (up to
 * MAX_GAP) and *every* request waits out the cool-down, not just the one that
 * was refused; after a run of successes the gap shrinks back.
 */
export const GENTLE_GAP_MS = 1000;
export const STANDARD_GAP_MS = 150;
const MAX_GAP_MS = 10000;
const RECOVERY_STREAK = 25;

/**
 * Browser (cookie-authenticated) traffic that SharePoint throttles hard is not
 * always answered with 429: the request is redirected to
 * /_layouts/15/Throttle.htm ("Something's not right"), which arrives as a 200
 * HTML page. That means "stop for a while", so the wait starts long and grows.
 */
const THROTTLE_PAGE_WAITS_S = [30, 60, 120, 240, 300];

/**
 * TypeScript compiles `class X extends Error` to ES5 for SPFx, and there the
 * object Error() returns is not an X - `instanceof X` is always false unless
 * the prototype is put back by hand. Every error class here does that.
 */
function fixPrototype(error: Error, proto: object, name: string): void {
  (Object as unknown as { setPrototypeOf: (o: object, p: object) => void }).setPrototypeOf(error, proto);
  error.name = name;
}

export class CancelledError extends Error {
  constructor(message = 'Cancelled by user.') {
    super(message);
    fixPrototype(this, CancelledError.prototype, 'CancelledError');
  }
}

export class HttpError extends Error {
  public status: number;
  public detail: string;
  constructor(status: number, detail: string) {
    super(`${status}: ${detail}`);
    fixPrototype(this, HttpError.prototype, 'HttpError');
    this.status = status;
    this.detail = detail;
  }
}

export function isCancelled(err: unknown): boolean {
  return err instanceof CancelledError || (!!err && (err as Error).name === 'CancelledError');
}

export function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export function newId(): string {
  return `${Date.now().toString(36)}${Math.random().toString(36).substring(2, 10)}`;
}

/** Escapes a server-relative path for use inside a '...' REST function argument. */
export function quoteForUrl(serverRelativeUrl: string): string {
  return encodeURIComponent(serverRelativeUrl.replace(/'/g, "''")).replace(/%2F/g, '/');
}

/** Retry-After in seconds, or a back-off that grows with each attempt. */
function retryAfterSeconds(response: SPHttpClientResponse, attempt: number): number {
  const header = response.headers.get('Retry-After');
  const seconds = header ? parseInt(header, 10) : NaN;
  return isNaN(seconds) ? Math.min(60, 2 * Math.pow(2, attempt)) : seconds;
}

function isThrottlePage(response: SPHttpClientResponse): boolean {
  const url = ((response as unknown as { url?: string }).url || '').toLowerCase();
  return url.indexOf('/_layouts/15/throttle.htm') >= 0;
}

/** A JSON or multipart API answered with an HTML page: SharePoint's throttle / error page. */
function looksLikeHtml(text: string): boolean {
  const start = text.replace(/^\s+/, '').substring(0, 15).toLowerCase();
  return start.indexOf('<!doctype') === 0 || start.indexOf('<html') === 0;
}

export class SpRest {
  private cancelled = false;
  private baseGapMs = GENTLE_GAP_MS;
  private gapMs = GENTLE_GAP_MS;
  private nextSlot = 0;
  private cooldownUntil = 0;
  private successStreak = 0;
  private throttlePageHits = 0;

  /** Told whenever all requests are paused for throttling, so the UI can say so. */
  public onThrottle: ((waitSeconds: number, reason: string) => void) | undefined;

  constructor(private context: WebPartContext) {}

  public cancel(): void {
    this.cancelled = true;
  }

  public reset(): void {
    this.cancelled = false;
  }

  public get isCancelled(): boolean {
    return this.cancelled;
  }

  /** Minimum pause between requests: GENTLE_GAP_MS or STANDARD_GAP_MS. */
  public setPace(gapMs: number): void {
    this.baseGapMs = gapMs;
    this.gapMs = gapMs;
  }

  public throwIfCancelled(): void {
    if (this.cancelled) {
      throw new CancelledError();
    }
  }

  /** A wait that ends early with CancelledError, so Stop doesn't sit out a long back-off. */
  public async sleep(ms: number): Promise<void> {
    const until = Date.now() + ms;
    for (;;) {
      this.throwIfCancelled();
      const left = until - Date.now();
      if (left <= 0) {
        return;
      }
      await delay(Math.min(left, 250));
    }
  }

  /** Called when SharePoint throttled a request, or one inside a $batch. */
  public noteThrottled(waitSeconds: number, reason = 'SharePoint asked to slow down'): void {
    const until = Date.now() + waitSeconds * 1000;
    if (until > this.cooldownUntil) {
      this.cooldownUntil = until;
      if (this.onThrottle) {
        this.onThrottle(waitSeconds, reason);
      }
    }
    this.gapMs = Math.min(MAX_GAP_MS, Math.max(this.gapMs * 2, 500));
    this.successStreak = 0;
  }

  private noteThrottlePage(): void {
    const wait = THROTTLE_PAGE_WAITS_S[Math.min(this.throttlePageHits, THROTTLE_PAGE_WAITS_S.length - 1)];
    this.throttlePageHits++;
    this.noteThrottled(wait, 'SharePoint is throttling this account');
  }

  private noteSuccess(): void {
    if (++this.successStreak >= RECOVERY_STREAK) {
      this.successStreak = 0;
      this.throttlePageHits = Math.max(0, this.throttlePageHits - 1);
      if (this.gapMs > this.baseGapMs) {
        this.gapMs = Math.max(this.baseGapMs, Math.floor(this.gapMs / 2));
      }
    }
  }

  /** Waits for this request's turn: after the pacing gap and any throttling cool-down. */
  private async pace(): Promise<void> {
    for (;;) {
      this.throwIfCancelled();
      const now = Date.now();
      const earliest = Math.max(this.nextSlot, this.cooldownUntil);
      if (earliest <= now) {
        this.nextSlot = now + this.gapMs;
        return;
      }
      await this.sleep(earliest - now);
    }
  }

  /**
   * Sends a request and returns its body text, waiting out throttling (429,
   * 503 with Retry-After, or the Throttle.htm page) and retrying dropped
   * connections and gateway errors. A throttled request is retried as it is;
   * it is never split into more requests.
   */
  private async fetchText(
    doFetch: () => Promise<SPHttpClientResponse>,
    url: string,
    allowNotFound = false
  ): Promise<{ status: number; text: string }> {
    let throttled = 0;
    let transient = 0;
    for (;;) {
      await this.pace();
      let response: SPHttpClientResponse;
      try {
        response = await doFetch();
      } catch (err) {
        if (++transient > MAX_NETWORK_RETRIES) {
          throw err;
        }
        await this.sleep(2000 * transient);
        continue;
      }
      if (response.status === 429 || response.status === 503) {
        if (++throttled > MAX_THROTTLE_RETRIES) {
          throw new HttpError(response.status, 'SharePoint throttled this request too many times.');
        }
        this.noteThrottled(retryAfterSeconds(response, throttled));
        continue;
      }
      // Gateway errors are usually momentary; other 5xx (timeouts, list view threshold) are not retried here.
      if ((response.status === 502 || response.status === 504) && transient < MAX_NETWORK_RETRIES) {
        transient++;
        await this.sleep(2000 * transient);
        continue;
      }
      if (allowNotFound && response.status === 404) {
        return { status: 404, text: '' };
      }
      if (!response.ok) {
        const detail = await describeError(response);
        console.error(`[ReInherit] ${response.status} from ${url}: ${detail}`);
        throw new HttpError(response.status, detail);
      }
      const text = await response.text();
      if (isThrottlePage(response) || looksLikeHtml(text)) {
        if (++throttled > MAX_THROTTLE_RETRIES) {
          throw new HttpError(429, 'SharePoint kept throttling this account (Throttle.htm). Try again later or outside business hours.');
        }
        console.warn(`[ReInherit] Throttle page instead of data from ${url}; backing off.`);
        this.noteThrottlePage();
        continue;
      }
      this.noteSuccess();
      return { status: response.status, text };
    }
  }

  public async getJson<T>(url: string): Promise<T> {
    const { text } = await this.fetchText(
      () => this.context.spHttpClient.get(url, SPHttpClient.configurations.v1),
      url
    );
    return (text ? JSON.parse(text) : {}) as T;
  }

  /** The response body, or undefined for a 404. */
  public async getText(url: string): Promise<string | undefined> {
    const { status, text } = await this.fetchText(
      () => this.context.spHttpClient.get(url, SPHttpClient.configurations.v1),
      url,
      true
    );
    return status === 404 ? undefined : text;
  }

  public async post(url: string, body?: string): Promise<number> {
    const { status } = await this.fetchText(
      () => this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, body === undefined ? {} : { body }),
      url
    );
    return status;
  }

  /**
   * Runs up to ~100 requests in one round trip through the web's $batch
   * endpoint. Throws if the batch as a whole fails or its response can't be
   * matched up with the requests, so callers can fall back to single calls.
   */
  public async batch(webUrl: string, requests: IBatchRequest[]): Promise<IBatchResponse[]> {
    const batchId = newId();
    const url = `${webUrl}/_api/$batch`;
    const { text } = await this.fetchText(
      () =>
        this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
          headers: {
            'Content-Type': `multipart/mixed; boundary="batch_${batchId}"`,
            Accept: 'application/json'
          },
          body: buildBatchBody(batchId, requests)
        }),
      url
    );
    const responses = parseBatchResponse(text);
    if (responses.length !== requests.length) {
      throw new Error(`Batch returned ${responses.length} responses for ${requests.length} requests.`);
    }
    return responses;
  }
}
