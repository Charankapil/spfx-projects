import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { buildBatchBody, IBatchRequest, IBatchResponse, parseBatchResponse } from './batch';
import { describeError } from './httpErrors';

/**
 * All calls go through the ambient SPHttpClient, which reuses the signed-in
 * user's SharePoint session and fetches / refreshes the form digest for
 * POSTs by itself. No AAD app registration, client id or secret is needed.
 */

const MAX_THROTTLE_RETRIES = 8;
const MAX_NETWORK_RETRIES = 3;

/**
 * Every request starts at least GAP ms after the previous one, across all the
 * parallel readers. When SharePoint throttles, the gap doubles (up to MAX_GAP)
 * and *every* request waits out the Retry-After period, not just the one that
 * was refused; after a run of successes the gap shrinks back. Without this the
 * parallel readers keep hammering while one of them waits, which is what
 * turns a throttled library into tenant-wide throttling.
 */
const BASE_GAP_MS = 150;
const MAX_GAP_MS = 5000;
const RECOVERY_STREAK = 25;

export class CancelledError extends Error {}

export class HttpError extends Error {
  constructor(public status: number, public detail: string) {
    super(`${status}: ${detail}`);
  }
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
function retryDelayMs(response: SPHttpClientResponse, attempt: number): number {
  const header = response.headers.get('Retry-After');
  const seconds = header ? parseInt(header, 10) : NaN;
  return (isNaN(seconds) ? Math.min(60, 2 * Math.pow(2, attempt)) : seconds) * 1000;
}

export class SpRest {
  private cancelled = false;
  private gapMs = BASE_GAP_MS;
  private nextSlot = 0;
  private cooldownUntil = 0;
  private successStreak = 0;

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

  public throwIfCancelled(): void {
    if (this.cancelled) {
      throw new CancelledError('Cancelled by user.');
    }
  }

  /** A wait that ends early with CancelledError, so Stop doesn't sit out a long Retry-After. */
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
  public noteThrottled(retryAfterSeconds: number): void {
    this.cooldownUntil = Math.max(this.cooldownUntil, Date.now() + retryAfterSeconds * 1000);
    this.gapMs = Math.min(MAX_GAP_MS, Math.max(this.gapMs * 2, 500));
    this.successStreak = 0;
  }

  private noteSuccess(): void {
    if (this.gapMs > BASE_GAP_MS && ++this.successStreak >= RECOVERY_STREAK) {
      this.gapMs = Math.max(BASE_GAP_MS, Math.floor(this.gapMs / 2));
      this.successStreak = 0;
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
   * Sends a request, waiting out 429/503 throttling (honouring Retry-After as
   * SharePoint asks) and retrying dropped connections - a run over a million
   * items is long enough that both will happen.
   */
  private async send(
    doFetch: () => Promise<SPHttpClientResponse>,
    url: string,
    allowNotFound = false
  ): Promise<SPHttpClientResponse> {
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
        this.noteThrottled(retryDelayMs(response, throttled) / 1000);
        continue;
      }
      // Gateway errors are usually momentary; other 5xx (timeouts, list view threshold) are not retried.
      if ((response.status === 502 || response.status === 504) && transient < MAX_NETWORK_RETRIES) {
        transient++;
        await this.sleep(2000 * transient);
        continue;
      }
      if (!response.ok && !(allowNotFound && response.status === 404)) {
        const detail = await describeError(response);
        console.error(`[ReInherit] ${response.status} from ${url}: ${detail}`);
        throw new HttpError(response.status, detail);
      }
      this.noteSuccess();
      return response;
    }
  }

  public async getJson<T>(url: string): Promise<T> {
    const response = await this.send(
      () => this.context.spHttpClient.get(url, SPHttpClient.configurations.v1),
      url
    );
    return (await response.json()) as T;
  }

  /** The response body, or undefined for a 404. */
  public async getText(url: string): Promise<string | undefined> {
    const response = await this.send(
      () => this.context.spHttpClient.get(url, SPHttpClient.configurations.v1),
      url,
      true
    );
    return response.status === 404 ? undefined : response.text();
  }

  public async post(url: string, body?: string): Promise<SPHttpClientResponse> {
    return this.send(
      () => this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, body === undefined ? {} : { body }),
      url
    );
  }

  /**
   * Runs up to ~100 requests in one round trip through the web's $batch
   * endpoint. Throws if the batch as a whole fails or its response can't be
   * matched up with the requests, so callers can fall back to single calls.
   */
  public async batch(webUrl: string, requests: IBatchRequest[]): Promise<IBatchResponse[]> {
    const batchId = newId();
    const url = `${webUrl}/_api/$batch`;
    const response = await this.send(
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
    const responses = parseBatchResponse(await response.text());
    if (responses.length !== requests.length) {
      throw new Error(`Batch returned ${responses.length} responses for ${requests.length} requests.`);
    }
    return responses;
  }
}
