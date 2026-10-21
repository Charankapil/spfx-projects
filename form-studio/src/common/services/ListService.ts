import { SPHttpClient } from '@microsoft/sp-http';

import { buildSelect, IItemPayload, readItem } from '../form/valueCodec';
import { ILookupValue, IUserValue, ITermValue, IFormValues } from '../form/formValues';
import { FieldKind, IFieldInfo } from '../model/IFieldInfo';
import {
  ensureOk,
  getJson,
  guidOnly,
  HttpError,
  ISpContext,
  odataString,
  postJson,
  quoteForUrl
} from './http';

/* eslint-disable @typescript-eslint/no-explicit-any */

export interface IListSummary {
  id: string;
  title: string;
  baseTemplate: number;
  itemCount: number;
  isLibrary: boolean;
}

export interface IListInfo extends IListSummary {
  enableAttachments: boolean;
  rootFolderUrl: string;
}

export interface IContentTypeInfo {
  id: string;
  name: string;
}

export interface IPersonSuggestion {
  key: string;
  title: string;
  email: string;
  description: string;
  isGroup: boolean;
  groupId?: number;
}

export interface ICurrentUser {
  id: number;
  name: string;
  email: string;
  loginName: string;
  isSiteAdmin: boolean;
  groups: string[];
}

export interface ILoadedItem {
  values: IFormValues;
  etag: string;
  contentTypeId: string;
}

const KIND_BY_TYPE: { [type: string]: FieldKind } = {
  Text: 'Text',
  Note: 'Note',
  Number: 'Number',
  Currency: 'Currency',
  Integer: 'Integer',
  Choice: 'Choice',
  MultiChoice: 'MultiChoice',
  Boolean: 'Boolean',
  DateTime: 'DateTime',
  Lookup: 'Lookup',
  LookupMulti: 'LookupMulti',
  User: 'User',
  UserMulti: 'UserMulti',
  URL: 'URL',
  TaxonomyFieldType: 'TaxonomyFieldType',
  TaxonomyFieldTypeMulti: 'TaxonomyFieldTypeMulti',
  Attachments: 'Attachments',
  Calculated: 'Calculated',
  Computed: 'Computed',
  Counter: 'Counter',
  File: 'Text'
};

/** Fields SharePoint shows nowhere in its own forms either. */
const SKIPPED_FIELDS = ['ComplianceAssetId', 'AppAuthor', 'AppEditor', 'FolderChildCount', 'ItemChildCount', 'ContentType'];

function numberOrUndefined(value: any): number | undefined {
  if (value === null || value === undefined || value === '') {
    return undefined;
  }
  const n = Number(value);
  // SharePoint reports "no limit" as +/- Double.MaxValue.
  return isNaN(n) || Math.abs(n) > 1e300 ? undefined : n;
}

function schemaAttr(schemaXml: string | undefined, name: string): string | undefined {
  if (!schemaXml) {
    return undefined;
  }
  const match = new RegExp(`\\s${name}="([^"]*)"`).exec(schemaXml);
  return match ? match[1] : undefined;
}

export function toFieldInfo(raw: any): IFieldInfo | undefined {
  const internalName: string = raw.InternalName;
  if (raw.Hidden || SKIPPED_FIELDS.indexOf(internalName) >= 0 || internalName.charAt(0) === '_' || internalName.indexOf('MediaService') === 0) {
    return undefined;
  }
  const kind: FieldKind = KIND_BY_TYPE[raw.TypeAsString] || 'Unsupported';
  // Projected lookup columns ("Customer: City") are read-only views of another list; leave them out.
  if ((kind === 'Lookup' || kind === 'LookupMulti') && raw.ReadOnlyField) {
    return undefined;
  }
  if (kind === 'Computed') {
    return undefined;
  }
  const schema: string | undefined = raw.SchemaXml;
  const info: IFieldInfo = {
    internalName,
    entityPropertyName: raw.EntityPropertyName || internalName,
    title: raw.Title,
    kind,
    typeAsString: raw.TypeAsString,
    description: raw.Description || '',
    required: !!raw.Required,
    readOnly: !!raw.ReadOnlyField || kind === 'Calculated' || kind === 'Counter' || kind === 'Unsupported' || raw.TypeAsString === 'File',
    defaultValue: raw.DefaultValue || undefined
  };
  switch (kind) {
    case 'Choice':
    case 'MultiChoice':
      info.choices = Array.isArray(raw.Choices) ? raw.Choices : (raw.Choices && raw.Choices.results) || [];
      info.fillIn = !!raw.FillInChoice;
      break;
    case 'Text':
      info.maxLength = numberOrUndefined(raw.MaxLength);
      break;
    case 'Note':
      info.richText = !!raw.RichText;
      info.appendOnly = !!raw.AppendOnly;
      break;
    case 'Number':
    case 'Currency':
    case 'Integer': {
      info.min = numberOrUndefined(raw.MinimumValue);
      info.max = numberOrUndefined(raw.MaximumValue);
      info.showAsPercentage = !!raw.ShowAsPercentage || schemaAttr(schema, 'Percentage') === 'TRUE';
      const decimals = schemaAttr(schema, 'Decimals');
      info.decimals = decimals !== undefined ? Number(decimals) : undefined;
      break;
    }
    case 'DateTime':
      info.dateOnly = raw.DisplayFormat === 0;
      break;
    case 'Lookup':
    case 'LookupMulti':
      info.lookupListId = raw.LookupList ? guidOnly(raw.LookupList) : undefined;
      info.lookupField = raw.LookupField || 'Title';
      break;
    case 'User':
    case 'UserMulti':
      info.userSelectionMode = raw.SelectionMode;
      info.userSelectionGroup = raw.SelectionGroup;
      break;
    case 'URL':
      info.urlFormat = raw.DisplayFormat;
      break;
    case 'TaxonomyFieldType':
    case 'TaxonomyFieldTypeMulti':
      info.termSetId = raw.TermSetId;
      info.anchorId = raw.AnchorId && raw.AnchorId !== '00000000-0000-0000-0000-000000000000' ? raw.AnchorId : undefined;
      break;
  }
  return info;
}

const MANAGE_LISTS_MASK = 0x800;

/**
 * Everything the forms need from SharePoint, through the current user's
 * session (SPHttpClient). No app registration, no Graph permissions.
 */
export class ListService {
  private fieldCache: { [listId: string]: Promise<IFieldInfo[]> } = {};
  private termCache: { [setId: string]: Promise<ITermValue[]> } = {};
  private userCache: Promise<ICurrentUser> | undefined;

  constructor(private ctx: ISpContext, public readonly webUrl: string = ctx.pageContext.web.absoluteUrl) {}

  private list(listId: string): string {
    return `${this.webUrl}/_api/web/lists(guid'${guidOnly(listId)}')`;
  }

  public async getLists(): Promise<IListSummary[]> {
    const json = await getJson<{ value: any[] }>(
      this.ctx,
      `${this.webUrl}/_api/web/lists?$select=Id,Title,BaseTemplate,ItemCount,BaseType&$filter=Hidden eq false and IsCatalog eq false&$orderby=Title`
    );
    return json.value
      .filter((l) => [100, 101, 102, 103, 104, 105, 106, 107, 108, 119, 171].indexOf(l.BaseTemplate) >= 0 || l.BaseTemplate >= 10000)
      .map((l) => ({ id: l.Id, title: l.Title, baseTemplate: l.BaseTemplate, itemCount: l.ItemCount, isLibrary: l.BaseType === 1 }));
  }

  public async getList(listId: string): Promise<IListInfo> {
    const l = await getJson<any>(
      this.ctx,
      `${this.list(listId)}?$select=Id,Title,BaseTemplate,ItemCount,BaseType,EnableAttachments,RootFolder/ServerRelativeUrl&$expand=RootFolder`
    );
    return {
      id: l.Id,
      title: l.Title,
      baseTemplate: l.BaseTemplate,
      itemCount: l.ItemCount,
      isLibrary: l.BaseType === 1,
      enableAttachments: !!l.EnableAttachments,
      rootFolderUrl: l.RootFolder.ServerRelativeUrl
    };
  }

  public getFields(listId: string): Promise<IFieldInfo[]> {
    const key = guidOnly(listId).toLowerCase();
    if (!this.fieldCache[key]) {
      this.fieldCache[key] = getJson<{ value: any[] }>(this.ctx, `${this.list(listId)}/fields?$filter=Hidden eq false`).then((json) =>
        json.value.map(toFieldInfo).filter((f): f is IFieldInfo => !!f)
      );
      this.fieldCache[key].catch(() => delete this.fieldCache[key]);
    }
    return this.fieldCache[key];
  }

  public async getContentTypes(listId: string): Promise<IContentTypeInfo[]> {
    const json = await getJson<{ value: any[] }>(this.ctx, `${this.list(listId)}/contenttypes?$select=StringId,Name,Hidden`);
    return json.value.filter((ct) => !ct.Hidden && String(ct.StringId).indexOf('0x0120') !== 0).map((ct) => ({ id: ct.StringId, name: ct.Name }));
  }

  public async canManageList(listId: string): Promise<boolean> {
    try {
      const json = await getJson<any>(this.ctx, `${this.list(listId)}/EffectiveBasePermissions`);
      const low = Number(json.Low !== undefined ? json.Low : json.EffectiveBasePermissions && json.EffectiveBasePermissions.Low);
      return (low & MANAGE_LISTS_MASK) === MANAGE_LISTS_MASK;
    } catch {
      return false;
    }
  }

  public getCurrentUser(): Promise<ICurrentUser> {
    if (!this.userCache) {
      this.userCache = getJson<any>(
        this.ctx,
        `${this.webUrl}/_api/web/currentuser?$select=Id,Title,Email,LoginName,IsSiteAdmin,Groups/Title&$expand=Groups`
      ).then((u) => ({
        id: u.Id,
        name: u.Title,
        email: u.Email,
        loginName: u.LoginName,
        isSiteAdmin: !!u.IsSiteAdmin,
        groups: (Array.isArray(u.Groups) ? u.Groups : (u.Groups && u.Groups.results) || []).map((g: any) => g.Title)
      }));
      this.userCache.catch(() => (this.userCache = undefined));
    }
    return this.userCache;
  }

  public async getSiteGroups(): Promise<string[]> {
    const json = await getJson<{ value: any[] }>(this.ctx, `${this.webUrl}/_api/web/sitegroups?$select=Title&$orderby=Title`);
    return json.value.map((g) => g.Title);
  }

  public async getItem(listId: string, itemId: number, fields: IFieldInfo[]): Promise<ILoadedItem> {
    const { select, expand } = buildSelect(fields);
    let url = `${this.list(listId)}/items(${itemId})?$select=${select.join(',')}`;
    if (expand.length) {
      url += `&$expand=${expand.join(',')}`;
    }
    // Minimal metadata, so the response carries the item's ETag for conflict detection.
    const response = await ensureOk(await this.ctx.spHttpClient.get(url, SPHttpClient.configurations.v1));
    const raw = await response.json();
    const etag = raw['@odata.etag'] || raw['odata.etag'] || response.headers.get('ETag') || '*';
    return { values: readItem(fields, raw), etag, contentTypeId: raw.ContentTypeId || '' };
  }

  /** Items of a list where `lookupField` points at `parentId` - used by the related items control. */
  public async getChildItems(listId: string, lookupField: IFieldInfo, parentId: number, fields: IFieldInfo[]): Promise<{ id: number; values: IFormValues }[]> {
    const { select, expand } = buildSelect(fields);
    let url = `${this.list(listId)}/items?$select=${select.join(',')}&$filter=${lookupField.entityPropertyName}Id eq ${parentId}&$top=500&$orderby=Id`;
    if (expand.length) {
      url += `&$expand=${expand.join(',')}`;
    }
    const json = await getJson<{ value: any[] }>(this.ctx, url);
    return json.value.map((raw) => ({ id: raw.Id, values: readItem(fields, raw) }));
  }

  public async createItem(listId: string, payload: IItemPayload, options: { folderPath?: string; contentTypeId?: string } = {}): Promise<number> {
    const body: { [key: string]: unknown } = { ...payload.body };
    if (options.contentTypeId) {
      body.ContentTypeId = options.contentTypeId;
    }
    let id: number;
    if (options.folderPath) {
      // Items in a folder can only be created through AddValidateUpdateItemUsingPath,
      // which takes text values; create the item there, then write typed values with MERGE.
      const result = await postJson<{ value: { FieldName: string; FieldValue: string; HasException: boolean; ErrorMessage: string }[] }>(
        this.ctx,
        `${this.list(listId)}/AddValidateUpdateItemUsingPath`,
        {
          listItemCreateInfo: { FolderPath: { DecodedUrl: options.folderPath }, UnderlyingObjectType: 0 },
          formValues: [],
          bNewDocumentUpdate: false
        }
      );
      const values = (result && result.value) || [];
      const failed = values.filter((v) => v.HasException)[0];
      if (failed) {
        throw new Error(`${failed.FieldName}: ${failed.ErrorMessage}`);
      }
      const idField = values.filter((v) => v.FieldName === 'Id')[0];
      id = Number(idField && idField.FieldValue);
      if (!id) {
        throw new Error('SharePoint did not return the new item id.');
      }
      if (Object.keys(body).length) {
        await this.updateItem(listId, id, { ...payload, body }, '*');
      }
    } else {
      const created = await postJson<{ Id: number }>(this.ctx, `${this.list(listId)}/items`, body);
      if (!created || !created.Id) {
        throw new Error('SharePoint did not return the new item id.');
      }
      id = created.Id;
    }
    await this.writeTaxonomy(listId, id, payload);
    return id;
  }

  /** Throws HttpError 412 when the item changed since `etag` was read. */
  public async updateItem(listId: string, itemId: number, payload: IItemPayload, etag = '*'): Promise<void> {
    if (Object.keys(payload.body).length) {
      await postJson(this.ctx, `${this.list(listId)}/items(${itemId})`, payload.body, { 'IF-MATCH': etag, 'X-HTTP-Method': 'MERGE' });
    }
    await this.writeTaxonomy(listId, itemId, payload);
  }

  private async writeTaxonomy(listId: string, itemId: number, payload: IItemPayload): Promise<void> {
    if (!payload.taxonomy.length) {
      return;
    }
    const result = await postJson<{ value: { FieldName: string; HasException: boolean; ErrorMessage: string }[] }>(
      this.ctx,
      `${this.list(listId)}/items(${itemId})/ValidateUpdateListItem`,
      { formValues: payload.taxonomy, bNewDocumentUpdate: false }
    );
    const failed = ((result && result.value) || []).filter((v) => v.HasException)[0];
    if (failed) {
      throw new Error(`${failed.FieldName}: ${failed.ErrorMessage}`);
    }
  }

  public async deleteItem(listId: string, itemId: number): Promise<void> {
    await postJson(this.ctx, `${this.list(listId)}/items(${itemId})`, undefined, { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE' });
  }

  public async addAttachment(listId: string, itemId: number, file: File): Promise<void> {
    const url = `${this.list(listId)}/items(${itemId})/AttachmentFiles/add(FileName='${quoteForUrl(file.name)}')`;
    await ensureOk(await this.ctx.spHttpClient.post(url, SPHttpClient.configurations.v1, { body: file }));
  }

  public async deleteAttachment(listId: string, itemId: number, fileName: string): Promise<void> {
    const url = `${this.list(listId)}/items(${itemId})/AttachmentFiles/getByFileName('${quoteForUrl(fileName)}')`;
    await postJson(this.ctx, url, undefined, { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE' });
  }

  /**
   * Options for a lookup field. `extraFilter` is an OData expression added
   * with "and" (cascading filters, the designer's custom filter).
   */
  public async searchLookup(field: IFieldInfo, query: string, extraFilter?: string): Promise<ILookupValue[]> {
    if (!field.lookupListId) {
      return [];
    }
    const display = field.lookupField || 'Title';
    const sourceFields = await this.getFields(field.lookupListId).catch(() => [] as IFieldInfo[]);
    const displayInfo = sourceFields.filter((f) => f.internalName === display)[0];
    const searchable = !displayInfo || displayInfo.kind === 'Text' || displayInfo.kind === 'Choice' || displayInfo.kind === 'Note';
    const filters: string[] = [];
    if (query && searchable) {
      filters.push(`substringof('${odataString(query)}',${display})`);
    }
    if (extraFilter) {
      filters.push(`(${extraFilter})`);
    }
    let url = `${this.list(field.lookupListId)}/items?$select=Id,${display}&$top=50&$orderby=${display}`;
    if (filters.length) {
      url += `&$filter=${encodeURIComponent(filters.join(' and '))}`;
    }
    const json = await getJson<{ value: any[] }>(this.ctx, url);
    let options = json.value.map((r) => ({ id: r.Id, title: r[display] === null || r[display] === undefined ? '' : String(r[display]) }));
    if (query && !searchable) {
      const q = query.toLowerCase();
      options = options.filter((o) => o.title.toLowerCase().indexOf(q) >= 0);
    }
    return options;
  }

  /** Builds the cascading-lookup filter: source field equals the value of a form field. */
  public async cascadeFilter(field: IFieldInfo, sourceField: string, formValue: unknown): Promise<string | undefined> {
    if (!field.lookupListId || formValue === null || formValue === undefined || formValue === '') {
      return undefined;
    }
    const sourceFields = await this.getFields(field.lookupListId);
    const source = sourceFields.filter((f) => f.internalName === sourceField)[0];
    const prop = source ? source.entityPropertyName : sourceField;
    const v = formValue as { id?: number; title?: string };
    if (source && (source.kind === 'Lookup' || source.kind === 'LookupMulti' || source.kind === 'User' || source.kind === 'UserMulti')) {
      return typeof v.id === 'number' ? `${prop}Id eq ${v.id}` : undefined;
    }
    if (source && (source.kind === 'Number' || source.kind === 'Currency' || source.kind === 'Integer' || source.kind === 'Counter')) {
      const n = typeof formValue === 'number' ? formValue : v.id;
      return typeof n === 'number' ? `${prop} eq ${n}` : undefined;
    }
    const text = typeof formValue === 'object' ? v.title || '' : String(formValue);
    return `${prop} eq '${odataString(text)}'`;
  }

  public async searchPeople(query: string, field?: IFieldInfo): Promise<IPersonSuggestion[]> {
    const url = `${this.webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`;
    const result = await postJson<any>(this.ctx, url, {
      queryParams: {
        AllowEmailAddresses: true,
        AllowMultipleEntities: false,
        AllUrlZones: false,
        MaximumEntitySuggestions: 20,
        PrincipalSource: 15,
        // 1 = users; 15 = users, distribution lists, security and SharePoint groups.
        PrincipalType: field && field.userSelectionMode === 1 ? 15 : 1,
        QueryString: query,
        SharePointGroupID: (field && field.userSelectionGroup) || 0
      }
    });
    const rawValue = result && (result.value !== undefined ? result.value : result.d && result.d.ClientPeoplePickerSearchUser);
    const entities: any[] = typeof rawValue === 'string' ? JSON.parse(rawValue) : [];
    return entities.map((e) => ({
      key: e.Key,
      title: e.DisplayText,
      email: (e.EntityData && (e.EntityData.Email || e.EntityData.PrincipalName)) || '',
      description: e.Description || (e.EntityData && e.EntityData.Title) || '',
      isGroup: e.EntityType === 'SPGroup',
      groupId: e.EntityData && e.EntityData.SPGroupID ? Number(e.EntityData.SPGroupID) : undefined
    }));
  }

  /** Turns a login name or e-mail into a site user with an id (adding them to the site's user list if needed). */
  public async ensureUser(logonName: string): Promise<IUserValue> {
    const user = await postJson<any>(this.ctx, `${this.webUrl}/_api/web/ensureuser`, { logonName });
    if (!user) {
      throw new HttpError(404, `Could not find a person matching "${logonName}".`);
    }
    return { id: user.Id, title: user.Title, email: user.Email, loginName: user.LoginName };
  }

  public async resolveSuggestion(s: IPersonSuggestion): Promise<IUserValue> {
    if (s.isGroup && s.groupId) {
      return { id: s.groupId, title: s.title, email: '' };
    }
    return this.ensureUser(s.key);
  }

  public async getLookupTitle(field: IFieldInfo, id: number): Promise<string> {
    if (!field.lookupListId) {
      return '';
    }
    const display = field.lookupField || 'Title';
    const r = await getJson<any>(this.ctx, `${this.list(field.lookupListId)}/items(${id})?$select=Id,${display}`);
    return r[display] === null || r[display] === undefined ? '' : String(r[display]);
  }

  /**
   * Terms of a term set (or of the anchor term), read through the SharePoint
   * v2.1 term store API. Capped so a huge term set cannot stall the form.
   */
  public getTerms(field: IFieldInfo): Promise<ITermValue[]> {
    const setId = field.termSetId;
    if (!setId) {
      return Promise.resolve([]);
    }
    const key = `${setId}|${field.anchorId || ''}`;
    if (!this.termCache[key]) {
      this.termCache[key] = this.loadTerms(setId, field.anchorId);
      this.termCache[key].catch(() => delete this.termCache[key]);
    }
    return this.termCache[key];
  }

  private async loadTerms(setId: string, anchorId?: string): Promise<ITermValue[]> {
    const base = `${this.webUrl}/_api/v2.1/termStore/sets/${setId}`;
    const lang = this.ctx.pageContext.cultureInfo.currentUICultureName.toLowerCase();
    const label = (t: any): string => {
      const labels: any[] = t.labels || [];
      const match = labels.filter((l) => (l.languageTag || '').toLowerCase() === lang)[0] || labels.filter((l) => l.isDefault)[0] || labels[0];
      return match ? match.name : t.id;
    };
    const terms: ITermValue[] = [];
    const queue: { url: string; path: string; depth: number }[] = [
      { url: anchorId ? `${base}/terms/${anchorId}/children` : `${base}/children`, path: '', depth: 0 }
    ];
    let requests = 0;
    while (queue.length && requests < 40 && terms.length < 1000) {
      const next = queue.shift() as { url: string; path: string; depth: number };
      requests++;
      const json = await getJson<{ value: any[] }>(this.ctx, `${next.url}?$top=500`);
      json.value.forEach((t) => {
        const name = label(t);
        const path = next.path ? `${next.path} > ${name}` : name;
        terms.push({ termGuid: t.id, label: name, path });
        if (next.depth < 3) {
          queue.push({ url: `${base}/terms/${t.id}/children`, path, depth: next.depth + 1 });
        }
      });
    }
    return terms;
  }
}
