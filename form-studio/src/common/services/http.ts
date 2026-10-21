import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { PageContext } from '@microsoft/sp-page-context';

/**
 * The parts of an SPFx context the services use. Web parts and the form
 * customizer both satisfy it, so the same services run in both.
 */
export interface ISpContext {
  spHttpClient: SPHttpClient;
  pageContext: PageContext;
}

/** Headers for JSON light writes: no __metadata needed in the body. */
export const NO_METADATA = {
  Accept: 'application/json;odata=nometadata',
  'Content-Type': 'application/json;odata=nometadata',
  'odata-version': ''
};

/**
 * SharePoint puts the useful part of a failure in the response body, so read
 * it rather than reporting a bare status code.
 */
export async function describeError(response: SPHttpClientResponse): Promise<string> {
  try {
    const body = await response.text();
    if (!body) {
      return response.statusText || 'no error details returned';
    }
    try {
      const parsed = JSON.parse(body);
      const message = parsed?.error?.message ?? parsed?.['odata.error']?.message;
      if (typeof message === 'string') {
        return message;
      }
      if (message && typeof message.value === 'string') {
        return message.value;
      }
    } catch {
      // Body was not JSON - fall through and surface the raw text instead.
    }
    return body.substring(0, 300);
  } catch {
    return response.statusText || 'no error details returned';
  }
}

export class HttpError extends Error {
  constructor(public status: number, message: string) {
    super(message);
    this.name = 'HttpError';
  }
}

export async function ensureOk(response: SPHttpClientResponse): Promise<SPHttpClientResponse> {
  if (!response.ok) {
    throw new HttpError(response.status, `${response.status}: ${await describeError(response)}`);
  }
  return response;
}

/** Quotes a value for use inside '...' in a REST URL. */
export function quoteForUrl(value: string): string {
  return encodeURIComponent(value.replace(/'/g, "''")).replace(/%2F/g, '/');
}

/** Quotes a value for use inside '...' in an OData $filter. */
export function odataString(value: string): string {
  return value.replace(/'/g, "''");
}

export async function getJson<T>(ctx: ISpContext, url: string): Promise<T> {
  const response = await ensureOk(await ctx.spHttpClient.get(url, SPHttpClient.configurations.v1, { headers: { Accept: NO_METADATA.Accept, 'odata-version': '' } }));
  return (await response.json()) as T;
}

export async function postJson<T>(ctx: ISpContext, url: string, body: unknown, headers?: { [name: string]: string }): Promise<T | undefined> {
  const options: ISPHttpClientOptions = {
    headers: { ...NO_METADATA, ...(headers || {}) },
    body: body === undefined ? undefined : JSON.stringify(body)
  };
  const response = await ensureOk(await ctx.spHttpClient.post(url, SPHttpClient.configurations.v1, options));
  if (response.status === 204) {
    return undefined;
  }
  const text = await response.text();
  return text ? (JSON.parse(text) as T) : undefined;
}

export function guidOnly(value: string): string {
  return value.replace(/[{}]/g, '');
}

export function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
