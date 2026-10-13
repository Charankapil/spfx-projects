import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { describeError } from './httpErrors';

/**
 * Thin wrapper over the ambient SPHttpClient: the signed-in user's own
 * SharePoint session, no app registration. Adds throttling retries
 * (honouring Retry-After) and turns failures into readable errors.
 */

const MAX_RETRIES = 4;

export class SpRequestError extends Error {
  constructor(message: string, public status: number) {
    super(message);
  }
}

function delay(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/** Escape a value for an OData string literal or an @alias parameter. */
export function odataString(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

export class SpRest {
  constructor(private client: SPHttpClient, public webUrl: string) {}

  public async fetch(url: string, method: 'GET' | 'POST', options: ISPHttpClientOptions = {}): Promise<SPHttpClientResponse> {
    const response = await this.send(url, method, options, false);
    if (!response) {
      throw new SpRequestError('404: not found', 404);
    }
    return response;
  }

  private async send(
    url: string,
    method: 'GET' | 'POST',
    options: ISPHttpClientOptions,
    allow404: boolean
  ): Promise<SPHttpClientResponse | undefined> {
    const absolute = url.indexOf('http') === 0 ? url : `${this.webUrl}${url}`;
    for (let attempt = 0; ; attempt++) {
      const response =
        method === 'GET'
          ? await this.client.get(absolute, SPHttpClient.configurations.v1, options)
          : await this.client.post(absolute, SPHttpClient.configurations.v1, options);
      if (response.ok) {
        return response;
      }
      if (allow404 && response.status === 404) {
        return undefined;
      }
      if ((response.status === 429 || response.status === 503) && attempt < MAX_RETRIES) {
        const retryAfter = parseInt(response.headers.get('Retry-After') || '', 10);
        await delay((isNaN(retryAfter) ? 2 ** attempt * 2 : retryAfter) * 1000);
        continue;
      }
      throw new SpRequestError(`${response.status}: ${await describeError(response)}`, response.status);
    }
  }

  public async getJson<T>(url: string): Promise<T> {
    const response = await this.fetch(url, 'GET');
    return (await response.json()) as T;
  }

  public async tryGetJson<T>(url: string): Promise<T | undefined> {
    const response = await this.send(url, 'GET', {}, true);
    return response ? ((await response.json()) as T) : undefined;
  }

  public async postJson<T>(url: string, body?: unknown, headers: Record<string, string> = {}): Promise<T> {
    const response = await this.fetch(url, 'POST', {
      headers: { 'Content-Type': 'application/json;odata=nometadata', ...headers },
      body: body === undefined ? undefined : JSON.stringify(body)
    });
    const text = await response.text();
    return (text ? JSON.parse(text) : {}) as T;
  }

  /** MERGE update of an existing entity (list item, list). */
  public async merge(url: string, body: unknown): Promise<void> {
    await this.fetch(url, 'POST', {
      headers: {
        'Content-Type': 'application/json;odata=nometadata',
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*'
      },
      body: JSON.stringify(body)
    });
  }

  public async getBinary(url: string): Promise<ArrayBuffer> {
    const response = await this.fetch(url, 'GET');
    return response.arrayBuffer();
  }

  /** Read every page of a collection (follows odata.nextLink) up to a limit. */
  public async getAll<T>(url: string, limit = 5000): Promise<T[]> {
    const items: T[] = [];
    let next: string | undefined = url;
    while (next && items.length < limit) {
      const page: { value: T[]; 'odata.nextLink'?: string; '@odata.nextLink'?: string } = await this.getJson(next);
      items.push(...(page.value || []));
      next = page['odata.nextLink'] || page['@odata.nextLink'];
    }
    return items;
  }
}
