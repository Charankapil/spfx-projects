import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/** A loosely typed SharePoint REST entity (list item, list, field...). */
export type SPEntity = Record<string, unknown>;

/** A page of results returned by a collection endpoint. */
export interface ISPCollection<T = SPEntity> {
  value: T[];
  nextLink?: string;
}

export type FieldKind = 'Text' | 'Note' | 'Number' | 'DateTime' | 'Boolean';

export interface IFieldDefinition {
  /** Internal (and static) name. Must be a valid XML name without spaces. */
  name: string;
  displayName: string;
  kind: FieldKind;
  indexed?: boolean;
  /** Add the field to the list's default view (default true). */
  showInView?: boolean;
}

export class SharePointError extends Error {
  public readonly status: number;
  public readonly retryAfter?: number;

  constructor(message: string, status: number, retryAfter?: number) {
    super(message);
    // Required for `instanceof` to work when TypeScript down-levels classes to ES5.
    Object.setPrototypeOf(this, SharePointError.prototype);
    this.name = 'SharePointError';
    this.status = status;
    this.retryAfter = retryAfter;
  }
}

/** SP.AddFieldOptions: AddFieldInternalNameHint (8) | AddFieldToDefaultView (16). */
const ADD_FIELD_INTERNAL_NAME_HINT = 8;
const ADD_FIELD_TO_DEFAULT_VIEW = 16;
const MAX_ATTEMPTS = 6;

/**
 * Thin SharePoint REST wrapper using SPHttpClient and the signed-in user's session.
 *
 * Uses the OData v4 "nometadata" format. (The 0.1 code sent `odata=verbose` through
 * SPHttpClient.configurations.v1, which also sends `OData-Version: 4.0`; SharePoint rejects
 * that combination. It also needed an extra request per item to discover the entity type.)
 */
export class SharePointService {
  private readonly webUrl: string;

  constructor(private readonly client: SPHttpClient, webUrl: string) {
    this.webUrl = webUrl.replace(/\/$/, '');
  }

  public static odataString(value: string): string {
    return String(value).replace(/'/g, "''");
  }

  public listPath(title: string): string {
    return `web/lists/getbytitle('${encodeURIComponent(SharePointService.odataString(title))}')`;
  }

  // ---------------------------------------------------------------------------
  // Core HTTP
  // ---------------------------------------------------------------------------

  public get<T = SPEntity>(path: string): Promise<T> {
    return this.request<T>('GET', this.url(path));
  }

  /** GET a collection, normalising the next-page link. */
  public async getCollection<T = SPEntity>(pathOrAbsoluteUrl: string): Promise<ISPCollection<T>> {
    const url = /^https?:\/\//i.test(pathOrAbsoluteUrl) ? pathOrAbsoluteUrl : this.url(pathOrAbsoluteUrl);
    if (url.indexOf(this.webUrl.split('/').slice(0, 3).join('/') + '/') !== 0) {
      throw new Error('Refusing to follow a continuation link to another host.');
    }
    const payload = await this.request<SPEntity>('GET', url);
    return {
      value: (payload.value as T[]) || [],
      nextLink: (payload['odata.nextLink'] || payload['@odata.nextLink']) as string | undefined
    };
  }

  /** Non-idempotent POST (create). Never retried automatically to avoid duplicate items. */
  public post<T = SPEntity>(path: string, body: unknown): Promise<T> {
    return this.request<T>('POST', this.url(path), body, undefined, false);
  }

  /** Idempotent MERGE update (retried on throttling). */
  public async merge(path: string, body: unknown): Promise<void> {
    await this.request('POST', this.url(path), body, { 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' }, true);
  }

  /** Idempotent DELETE (retried on throttling; a 404 after a retry is treated as success). */
  public async delete(path: string): Promise<void> {
    await this.request('POST', this.url(path), undefined, { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE' }, true, true);
  }

  // ---------------------------------------------------------------------------
  // Lists, fields and items
  // ---------------------------------------------------------------------------

  public async getListTitles(prefix: string): Promise<string[]> {
    const r = await this.getCollection<{ Title: string }>(
      `web/lists?$select=Title&$filter=${encodeURIComponent(`startswith(Title,'${SharePointService.odataString(prefix)}')`)}&$top=5000`
    );
    return r.value.map(l => l.Title);
  }

  public async ensureList(title: string, description: string): Promise<void> {
    try {
      await this.get(`${this.listPath(title)}?$select=Id`);
      return;
    } catch (e) {
      if (!SharePointService.isNotFound(e)) throw e;
    }
    await this.post('web/lists', {
      Title: title,
      Description: description,
      BaseTemplate: 100,
      AllowContentTypes: true,
      ContentTypesEnabled: false
    });
  }

  /**
   * Creates any missing fields on a list using CreateFieldAsXml so the internal name is
   * exactly the one requested. (POSTing an SP.Field with `StaticName` - as 0.1 did - derives the
   * internal name from the *display* name, e.g. "Workspace_x0020_ID", which breaks every
   * subsequent query and item write that uses `WorkspaceId`.)
   */
  public async ensureFields(listTitle: string, fields: IFieldDefinition[]): Promise<void> {
    const existing = await this.getCollection<{ InternalName: string }>(
      `${this.listPath(listTitle)}/fields?$select=InternalName&$top=5000`
    );
    const names = new Set(existing.value.map(f => f.InternalName));
    for (const field of fields) {
      if (names.has(field.name)) continue;
      const options = ADD_FIELD_INTERNAL_NAME_HINT | (field.showInView === false ? 0 : ADD_FIELD_TO_DEFAULT_VIEW);
      await this.post(`${this.listPath(listTitle)}/fields/createfieldasxml`, {
        parameters: { SchemaXml: SharePointService.fieldXml(field), Options: options }
      });
    }
  }

  public async getListItemCount(title: string): Promise<number> {
    const r = await this.get<{ ItemCount?: number }>(`${this.listPath(title)}?$select=ItemCount`);
    return r.ItemCount || 0;
  }

  public createItem<T = SPEntity>(listTitle: string, fields: SPEntity): Promise<T & { Id: number }> {
    return this.post<T & { Id: number }>(`${this.listPath(listTitle)}/items`, fields);
  }

  public updateItem(listTitle: string, id: number, fields: SPEntity): Promise<void> {
    return this.merge(`${this.listPath(listTitle)}/items(${id})`, fields);
  }

  public deleteItem(listTitle: string, id: number): Promise<void> {
    return this.delete(`${this.listPath(listTitle)}/items(${id})`);
  }

  public static isNotFound(e: unknown): boolean {
    return e instanceof SharePointError && e.status === 404;
  }

  public static errorMessage(e: unknown, fallback: string = 'An unexpected error occurred.'): string {
    return e instanceof Error && e.message ? e.message : fallback;
  }

  // ---------------------------------------------------------------------------
  // Internals
  // ---------------------------------------------------------------------------

  private url(path: string): string {
    return `${this.webUrl}/_api/${path}`;
  }

  private static fieldXml(f: IFieldDefinition): string {
    const attr = (v: string): string =>
      v.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const common = `Name="${attr(f.name)}" StaticName="${attr(f.name)}" DisplayName="${attr(f.displayName)}" Required="FALSE"` +
      (f.indexed && f.kind !== 'Note' ? ' Indexed="TRUE"' : '');
    switch (f.kind) {
      case 'Text': return `<Field Type="Text" ${common} MaxLength="255" />`;
      case 'Note': return `<Field Type="Note" ${common} NumLines="6" RichText="FALSE" />`;
      case 'Number': return `<Field Type="Number" ${common} />`;
      case 'DateTime': return `<Field Type="DateTime" ${common} Format="DateTime" />`;
      case 'Boolean': return `<Field Type="Boolean" ${common}><Default>0</Default></Field>`;
      default: throw new Error(`Unsupported field kind for ${f.name}`);
    }
  }

  private async request<T>(
    method: 'GET' | 'POST',
    url: string,
    body?: unknown,
    extraHeaders?: Record<string, string>,
    retryable: boolean = method === 'GET',
    notFoundAfterRetryIsSuccess: boolean = false
  ): Promise<T> {
    const headers: Record<string, string> = {
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=nometadata',
      'odata-version': '',
      ...(extraHeaders || {})
    };

    for (let attempt = 0; ; attempt++) {
      try {
        const response = method === 'GET'
          ? await this.client.get(url, SPHttpClient.configurations.v1, { headers })
          : await this.client.post(url, SPHttpClient.configurations.v1, {
              headers,
              body: body === undefined ? undefined : JSON.stringify(body)
            });
        return await SharePointService.parse<T>(response);
      } catch (e) {
        const status = e instanceof SharePointError ? e.status : 0;
        if (attempt > 0 && notFoundAfterRetryIsSuccess && status === 404) return {} as T;
        if (!retryable || (status !== 429 && status !== 503) || attempt >= MAX_ATTEMPTS - 1) throw e;
        const retryAfter = (e as SharePointError).retryAfter;
        const delay = retryAfter && retryAfter > 0
          ? Math.min(120000, retryAfter * 1000)
          : Math.min(60000, 1000 * Math.pow(2, attempt));
        await new Promise<void>(resolve => setTimeout(resolve, delay));
      }
    }
  }

  private static async parse<T>(response: SPHttpClientResponse): Promise<T> {
    if (!response.ok) {
      const text = await response.text();
      let detail = text || response.statusText;
      try {
        const json = JSON.parse(text);
        const err = json['odata.error'] || json.error;
        if (err && err.message) detail = typeof err.message === 'string' ? err.message : err.message.value;
      } catch {
        // Non-JSON body; keep the raw text.
      }
      const retry = Number(response.headers.get('Retry-After'));
      throw new SharePointError(
        `SharePoint request failed (${response.status}): ${detail}`,
        response.status,
        Number.isFinite(retry) ? retry : undefined
      );
    }
    if (response.status === 204) return {} as T;
    const text = await response.text();
    return (text ? JSON.parse(text) : {}) as T;
  }
}
