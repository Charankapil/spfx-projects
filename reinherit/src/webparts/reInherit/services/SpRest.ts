import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { buildBatchBody, IBatchRequest, IBatchResponse, parseBatchResponse } from './batch';
import { describeError } from './httpErrors';

/**
 * All calls go through the ambient SPHttpClient, which reuses the signed-in
 * user's SharePoint session and fetches / refreshes the form digest for
 * POSTs by itself. No AAD app registration, client id or secret is needed.
 */

const MAX_THROTTLE_RETRIES = 6;
const MAX_NETWORK_RETRIES = 3;

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

  /**
   * Sends a request, waiting out 429/503 throttling (honouring Retry-After as
   * SharePoint asks) and retrying dropped connections - a run over a million
   * items is long enough that both will happen.
   */
  private async send(
    doFetch: () => Promise<SPHttpClientResponse>,
    url: string
  ): Promise<SPHttpClientResponse> {
    let throttled = 0;
    let networkFailures = 0;
    for (;;) {
      this.throwIfCancelled();
      let response: SPHttpClientResponse;
      try {
        response = await doFetch();
      } catch (err) {
        if (++networkFailures > MAX_NETWORK_RETRIES) {
          throw err;
        }
        await delay(2000 * networkFailures);
        continue;
      }
      if (response.status === 429 || response.status === 503) {
        if (++throttled > MAX_THROTTLE_RETRIES) {
          throw new HttpError(response.status, 'SharePoint throttled this request too many times.');
        }
        await delay(retryDelayMs(response, throttled));
        continue;
      }
      if (!response.ok) {
        const detail = await describeError(response);
        console.error(`[ReInherit] ${response.status} from ${url}: ${detail}`);
        throw new HttpError(response.status, detail);
      }
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

  public async getText(url: string): Promise<string | undefined> {
    this.throwIfCancelled();
    const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (response.status === 404) {
      return undefined;
    }
    if (!response.ok) {
      throw new HttpError(response.status, await describeError(response));
    }
    return response.text();
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
