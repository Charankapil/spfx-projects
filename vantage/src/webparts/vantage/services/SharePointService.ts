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
  /** Enforce unique values (implies indexed). Only applied when the field is created. */
  unique?: boolean;
  /** Add the field to the list's default view (default true). */
  showInView?: boolean;
}

export interface IAttachment {
  FileName: string;
  ServerRelativeUrl: string;
}

export interface IPersonSuggestion {
  /** Claims login, e.g. i:0#.f|membership|user@contoso.com */
  key: string;
  displayName: string;
  email?: string;
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

/** Built-in role definition IDs. */
export const ROLE = { fullControl: 1073741829, contribute: 1073741827, read: 1073741826 };

/**
 * Thin SharePoint REST wrapper using SPHttpClient and the signed-in user's session.
 *
 * All JSON requests use the OData "nometadata" format, so request bodies must NOT contain
 * `__metadata` (SharePoint rejects it in this format) and no entity-type lookups are needed.
 */
export class SharePointService {
  private readonly webUrl: string;
  private readonly origin: string;

  constructor(private readonly client: SPHttpClient, webUrl: string) {
    this.webUrl = webUrl.replace(/\/$/, '');
    this.origin = this.webUrl.split('/').slice(0, 3).join('/');
  }

  public static odataString(value: string): string {
    return String(value).replace(/'/g, "''");
  }

  /** Build an encoded `$filter=` query parameter (always encode: claims logins contain `#` and `|`). */
  public static filter(expression: string): string {
    return `$filter=${encodeURIComponent(expression)}`;
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
    if (url.indexOf(this.origin + '/') !== 0) {
      throw new Error('Refusing to follow a continuation link to another host.');
    }
    const payload = await this.request<SPEntity>('GET', url);
    return {
      value: (payload.value as T[]) || [],
      nextLink: (payload['odata.nextLink'] || payload['@odata.nextLink']) as string | undefined
    };
  }

  /** Follow continuation links up to `maxPages` and return every row. */
  public async getAll<T = SPEntity>(path: string, maxPages: number = 20): Promise<T[]> {
    const out: T[] = [];
    let next: string | undefined = path;
    for (let page = 0; next && page < maxPages; page++) {
      const r: ISPCollection<T> = await this.getCollection<T>(next);
      out.push(...r.value);
      next = r.nextLink;
    }
    return out;
  }

  /** Non-idempotent POST (create). Never retried automatically to avoid duplicate items. */
  public post<T = SPEntity>(path: string, body?: unknown): Promise<T> {
    return this.request<T>('POST', this.url(path), body === undefined ? {} : body, undefined, false);
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
    const rows = await this.getAll<{ Title: string }>(
      `web/lists?$select=Title&${SharePointService.filter(`startswith(Title,'${SharePointService.odataString(prefix)}')`)}&$top=500`
    );
    return rows.map(l => l.Title);
  }

  public async getSiteListCount(): Promise<number> {
    return (await this.getAll('web/lists?$select=Id&$top=500', 10)).length;
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
      ContentTypesEnabled: false,
      EnableAttachments: true
    });
  }

  /**
   * Creates any missing fields on a list using CreateFieldAsXml so the internal name is
   * exactly the one requested. (POSTing an SP.Field with `StaticName` derives the internal
   * name from the *display* name, e.g. "Workspace_x0020_ID", which breaks every query.)
   */
  public async ensureFields(listTitle: string, fields: IFieldDefinition[]): Promise<void> {
    const existing = await this.getAll<{ InternalName: string }>(
      `${this.listPath(listTitle)}/fields?$select=InternalName&$top=500`
    );
    const names = new Set(existing.map(f => f.InternalName));
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
    return this.post<T & { Id: number }>(`${this.listPath(listTitle)}/items`, SharePointService.clean(fields));
  }

  public updateItem(listTitle: string, id: number, fields: SPEntity): Promise<void> {
    return this.merge(`${this.listPath(listTitle)}/items(${id})`, SharePointService.clean(fields));
  }

  public deleteItem(listTitle: string, id: number): Promise<void> {
    return this.delete(`${this.listPath(listTitle)}/items(${id})`);
  }

  // ---------------------------------------------------------------------------
  // Attachments (binary)
  // ---------------------------------------------------------------------------

  public async listAttachments(listTitle: string, itemId: number): Promise<IAttachment[]> {
    const r = await this.getCollection<IAttachment>(`${this.listPath(listTitle)}/items(${itemId})/AttachmentFiles?$select=FileName,ServerRelativeUrl`);
    return r.value;
  }

  public async getAttachmentContent(file: IAttachment): Promise<ArrayBuffer> {
    const url = this.url(`web/getfilebyserverrelativeurl('${encodeURIComponent(SharePointService.odataString(file.ServerRelativeUrl))}')/$value`);
    for (let attempt = 0; ; attempt++) {
      const response = await this.client.get(url, SPHttpClient.configurations.v1, { headers: { Accept: 'application/octet-stream' } });
      if (response.ok) return response.arrayBuffer();
      if ((response.status === 429 || response.status === 503) && attempt < MAX_ATTEMPTS - 1) {
        await SharePointService.wait(Number(response.headers.get('Retry-After')), attempt);
        continue;
      }
      throw new SharePointError(`Attachment download failed (${response.status}) for ${file.FileName}`, response.status);
    }
  }

  public async addAttachment(listTitle: string, itemId: number, fileName: string, content: ArrayBuffer): Promise<void> {
    const url = this.url(`${this.listPath(listTitle)}/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(SharePointService.odataString(fileName))}')`);
    const response = await this.client.post(url, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/octet-stream' },
      body: content
    });
    await SharePointService.parse(response);
  }

  public deleteAttachment(listTitle: string, itemId: number, fileName: string): Promise<void> {
    return this.delete(`${this.listPath(listTitle)}/items(${itemId})/AttachmentFiles/getbyfilename('${encodeURIComponent(SharePointService.odataString(fileName))}')`);
  }

  // ---------------------------------------------------------------------------
  // People, groups and permissions
  // ---------------------------------------------------------------------------

  public async searchPeople(query: string, max: number = 8): Promise<IPersonSuggestion[]> {
    if (query.trim().length < 2) return [];
    const r = await this.post<{ value?: string }>(
      'SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser',
      {
        queryParams: {
          QueryString: query.trim(), MaximumEntitySuggestions: max, AllowEmailAddresses: true,
          AllowOnlyEmailAddresses: false, PrincipalSource: 15, PrincipalType: 1, AllUrlZones: false, AllowMultipleEntities: false
        }
      }
    );
    const entities: Array<{ Key: string; DisplayText: string; EntityData?: { Email?: string } }> = r.value ? JSON.parse(r.value) : [];
    return entities.map(e => ({ key: e.Key, displayName: e.DisplayText, email: e.EntityData && e.EntityData.Email }));
  }

  public ensureUser(loginName: string): Promise<{ Id: number; LoginName: string; Title: string }> {
    return this.post('web/ensureuser', { logonName: SharePointService.claims(loginName) });
  }

  public async ensureGroup(title: string, description: string): Promise<number> {
    try {
      return (await this.get<{ Id: number }>(`web/sitegroups/getbyname('${encodeURIComponent(SharePointService.odataString(title))}')?$select=Id`)).Id;
    } catch (e) {
      if (!SharePointService.isNotFound(e)) throw e;
    }
    const g = await this.post<{ Id: number }>('web/sitegroups', {
      Title: title, Description: description, AllowMembersEditMembership: false, AllowRequestToJoinLeave: false, OnlyAllowMembersViewMembership: false
    });
    return g.Id;
  }

  public async addUserToGroup(groupId: number, loginName: string): Promise<void> {
    const user = await this.ensureUser(loginName);
    await this.post(`web/sitegroups/getbyid(${groupId})/users`, { LoginName: user.LoginName });
  }

  public async removeUserFromGroup(groupId: number, userId: number): Promise<void> {
    await this.post(`web/sitegroups/getbyid(${groupId})/users/removebyid(${userId})`);
  }

  public async getGroupUsers(groupId: number): Promise<Array<{ Id: number; Title: string; LoginName: string; Email?: string }>> {
    return this.getAll(`web/sitegroups/getbyid(${groupId})/users?$select=Id,Title,LoginName,Email&$top=500`, 10);
  }

  public async getCurrentUserGroupTitles(): Promise<string[]> {
    return (await this.getAll<{ Title: string }>('web/currentuser/groups?$select=Title&$top=500', 10)).map(g => g.Title);
  }

  public async getAssociatedOwnerGroupId(): Promise<number> {
    return (await this.get<{ Id: number }>('web/associatedownergroup?$select=Id')).Id;
  }

  /** Break inheritance (without copying) on a list or list item and grant the given principals. */
  public async setUniquePermissions(securablePath: string, grants: Array<{ principalId: number; roleId: number }>): Promise<void> {
    const meta = await this.get<{ HasUniqueRoleAssignments?: boolean }>(`${securablePath}?$select=HasUniqueRoleAssignments`);
    if (!meta.HasUniqueRoleAssignments) {
      await this.post(`${securablePath}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`);
    }
    const assigned = await this.getAll<{ PrincipalId: number }>(`${securablePath}/roleassignments?$select=PrincipalId&$top=500`, 5);
    const existing = new Set(assigned.map(a => a.PrincipalId));
    for (const g of grants) {
      if (existing.has(g.principalId)) continue;
      await this.post(`${securablePath}/roleassignments/addroleassignment(principalid=${g.principalId},roledefid=${g.roleId})`);
    }
  }

  // ---------------------------------------------------------------------------
  // Helpers
  // ---------------------------------------------------------------------------

  public static claims(login: string): string {
    const l = login.trim();
    return l.indexOf('|') >= 0 ? l : `i:0#.f|membership|${l}`;
  }

  public static isNotFound(e: unknown): boolean {
    return e instanceof SharePointError && e.status === 404;
  }

  public static errorMessage(e: unknown, fallback: string = 'An unexpected error occurred.'): string {
    return e instanceof Error && e.message ? e.message : fallback;
  }

  private url(path: string): string {
    return `${this.webUrl}/_api/${path}`;
  }

  /** Drop `undefined` values so they are not serialised as missing properties. */
  private static clean(fields: SPEntity): SPEntity {
    const out: SPEntity = {};
    Object.keys(fields).forEach(k => { if (fields[k] !== undefined) out[k] = fields[k]; });
    return out;
  }

  private static fieldXml(f: IFieldDefinition): string {
    const attr = (v: string): string =>
      v.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const indexed = (f.indexed || f.unique) && f.kind !== 'Note';
    const common = `Name="${attr(f.name)}" StaticName="${attr(f.name)}" DisplayName="${attr(f.displayName)}" Required="FALSE"` +
      (indexed ? ' Indexed="TRUE"' : '') + (f.unique && f.kind !== 'Note' ? ' EnforceUniqueValues="TRUE"' : '');
    switch (f.kind) {
      case 'Text': return `<Field Type="Text" ${common} MaxLength="255" />`;
      case 'Note': return `<Field Type="Note" ${common} NumLines="6" RichText="FALSE" />`;
      case 'Number': return `<Field Type="Number" ${common} />`;
      case 'DateTime': return `<Field Type="DateTime" ${common} Format="DateTime" />`;
      case 'Boolean': return `<Field Type="Boolean" ${common}><Default>0</Default></Field>`;
      default: throw new Error(`Unsupported field kind for ${f.name}`);
    }
  }

  private static wait(retryAfterSeconds: number, attempt: number): Promise<void> {
    const delay = retryAfterSeconds > 0
      ? Math.min(120000, retryAfterSeconds * 1000)
      : Math.min(60000, 1000 * Math.pow(2, attempt));
    return new Promise<void>(resolve => setTimeout(resolve, delay));
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
        await SharePointService.wait((e as SharePointError).retryAfter || 0, attempt);
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
