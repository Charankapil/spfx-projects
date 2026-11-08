import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ApprovalDecision, ICurrentUser, IListItem, IListLoadResult, IPerson, ISearchDocument, TaskStatus } from '../models/Models';
import { mapSearchRows, ISearchRow } from '../utils/searchMapping';
import { approvalAssignee, buildKql, isPending, sameEmail } from '../utils/helpers';

export const LISTS = {
  requests: 'ES Requests',
  approvals: 'ES Approvals',
  announcements: 'ES Announcements',
  tasks: 'ES Tasks'
};

const COMMON_FIELDS = ['Id', 'Title', 'Description', 'Category', 'Owner', 'Created', 'Modified', 'Author/Id', 'Author/Title', 'Editor/Title'];

/**
 * Each list only has the columns the provisioning script creates for it. Selecting a column that
 * does not exist on a list makes SharePoint reject the whole query (HTTP 400), so the $select must
 * be list-specific.
 */
export const LIST_FIELDS: Record<string, string[]> = {
  [LISTS.requests]: [...COMMON_FIELDS, 'Status', 'DueDate', 'RequesterName', 'RequesterEmail'],
  [LISTS.approvals]: [...COMMON_FIELDS, 'Status', 'DueDate', 'ApproverEmail', 'DecisionBy', 'DecisionDate'],
  [LISTS.announcements]: [...COMMON_FIELDS],
  [LISTS.tasks]: [...COMMON_FIELDS, 'Status', 'DueDate']
};

/** SharePoint Online "Local People Results" result source. People are excluded from the default source. */
const PEOPLE_SOURCE_ID = 'b09a7990-05ea-4af9-81ef-edfab16c4e31';
const SEARCH_LIMIT = 50;
const LIST_LIMIT = 100;

export class SharePointError extends Error {
  constructor(public readonly status: number, message: string) {
    super(message);
    this.name = 'SharePointError';
    // Restore the prototype chain: subclassing Error is broken when compiled to ES5.
    Object.setPrototypeOf(this, SharePointError.prototype);
  }
}

interface IRequestOptions { method?: string; headers?: Record<string, string>; body?: string; }
interface IRawListItem extends Omit<IListItem, 'etag'> { '@odata.etag'?: string; 'odata.etag'?: string; }
interface ISearchResponse { PrimaryQueryResult?: { RelevantResults?: { Table?: { Rows?: ISearchRow[] } } }; }

export class SharePointService {
  private readonly webUrl: string;

  constructor(private readonly context: WebPartContext) {
    this.webUrl = context.pageContext.web.absoluteUrl;
  }

  private async send(url: string, options?: IRequestOptions): Promise<SPHttpClientResponse> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.fetch(url, SPHttpClient.configurations.v1, {
      method: options?.method || 'GET',
      headers: { Accept: 'application/json;odata=nometadata', ...(options?.headers || {}) },
      body: options?.body
    });
    if (!response.ok) {
      let detail = response.statusText;
      try {
        const payload = await response.json();
        detail = payload?.error?.message?.value || payload?.['odata.error']?.message?.value || payload?.error?.message || detail;
      } catch {
        /* response did not contain JSON */
      }
      if (response.status === 412) detail = 'This item was changed by someone else. Refresh and try again.';
      throw new SharePointError(response.status, `SharePoint request failed (${response.status}): ${detail}`);
    }
    return response;
  }

  private async request<T>(url: string, options?: IRequestOptions): Promise<T> {
    const response = await this.send(url, options);
    if (response.status === 204) return undefined as unknown as T;
    return response.json();
  }

  private listUrl(title: string): string {
    return `${this.webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(title.replace(/'/g, "''"))}')`;
  }

  public async getCurrentUser(): Promise<ICurrentUser> {
    return this.request<ICurrentUser>(`${this.webUrl}/_api/web/currentuser?$select=Id,Title,Email,LoginName`);
  }

  public async getListItems(listTitle: string): Promise<IListItem[]> {
    const fields = (LIST_FIELDS[listTitle] || COMMON_FIELDS).join(',');
    const url = `${this.listUrl(listTitle)}/items?$select=${encodeURIComponent(fields)}&$expand=Author,Editor&$top=${LIST_LIMIT}&$orderby=Modified desc`;
    // minimal metadata so each item carries its ETag (nometadata strips it, which silently disabled concurrency checks).
    const data = await this.request<{ value: IRawListItem[] }>(url, { headers: { Accept: 'application/json;odata.metadata=minimal' } });
    return (data.value || []).map(raw => {
      const { ['@odata.etag']: v4Etag, ['odata.etag']: v3Etag, ...rest } = raw;
      return { ...rest, etag: v4Etag || v3Etag };
    });
  }

  /** Loads a list; a missing list (404) is an empty module, any other failure is reported per list. */
  public async loadList(listTitle: string): Promise<IListLoadResult> {
    try {
      return { items: await this.getListItems(listTitle) };
    } catch (error) {
      if ((error as SharePointError).status === 404) return { items: [] };
      return { items: [], error: `${listTitle}: ${(error as Error).message}` };
    }
  }

  public async createRequest(input: { title: string; description: string; category: string; requesterName: string; requesterEmail: string }): Promise<IListItem> {
    return this.createListItem(LISTS.requests, {
      Title: input.title.trim(), Description: input.description.trim(), Category: input.category,
      Status: 'Submitted', RequesterName: input.requesterName, RequesterEmail: input.requesterEmail
    });
  }

  public async createListItem(listTitle: string, fields: Record<string, unknown>): Promise<IListItem> {
    return this.request<IListItem>(`${this.listUrl(listTitle)}/items`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json;odata=nometadata' },
      body: JSON.stringify(fields)
    });
  }

  /** MERGE-updates an item and returns the item's new ETag (when SharePoint sends one). */
  public async updateListItem(listTitle: string, id: number, fields: Record<string, unknown>, etag = '*'): Promise<string | undefined> {
    const response = await this.send(`${this.listUrl(listTitle)}/items(${id})`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json;odata=nometadata', 'IF-MATCH': etag, 'X-HTTP-Method': 'MERGE' },
      body: JSON.stringify(fields)
    });
    return response.headers.get('ETag') || undefined;
  }

  public async setApprovalStatus(item: IListItem, status: ApprovalDecision, actorEmail: string): Promise<void> {
    if (!isPending(item)) throw new Error('This approval is no longer pending. Refresh the list and try again.');
    // Unassigned items must not be actionable by anyone who happens to call this.
    if (!sameEmail(approvalAssignee(item), actorEmail)) throw new Error('This approval is not assigned to you.');
    await this.updateListItem(LISTS.approvals, item.Id, { Status: status, DecisionBy: actorEmail, DecisionDate: new Date().toISOString() }, item.etag || '*');
  }

  public async setTaskStatus(item: IListItem, status: TaskStatus): Promise<string | undefined> {
    return this.updateListItem(LISTS.tasks, item.Id, { Status: status }, item.etag || '*');
  }

  private async search(queryText: string, properties: string, extraParams = ''): Promise<Record<string, string>[]> {
    const escapedQuery = queryText.replace(/'/g, "''");
    const url = `${this.webUrl}/_api/search/query?querytext='${encodeURIComponent(escapedQuery)}'&selectproperties='${encodeURIComponent(properties)}'&rowlimit=${SEARCH_LIMIT}&trimduplicates=true${extraParams}`;
    const data = await this.request<ISearchResponse>(url);
    return mapSearchRows(data?.PrimaryQueryResult?.RelevantResults?.Table?.Rows);
  }

  public async searchDocuments(query: string): Promise<ISearchDocument[]> {
    const rows = await this.search(buildKql(query, 'IsDocument:1'), 'Title,Path,Author,LastModifiedTime,FileExtension');
    return rows.map(c => ({ title: c.Title || 'Untitled document', path: c.Path || '', author: c.Author || '', modified: c.LastModifiedTime || '', fileType: c.FileExtension || '' }));
  }

  public async searchPeople(query: string): Promise<IPerson[]> {
    // Without the people result source, the default "Local SharePoint Results" source excludes people and this returned nothing.
    const q = (query || '').trim() || '*';
    const rows = await this.search(q, 'Title,Path,WorkEmail,Department,JobTitle,BaseOfficeLocation,OfficeNumber,PictureURL,Skills', `&sourceid='${PEOPLE_SOURCE_ID}'`);
    return rows.map(c => ({
      title: c.Title || '', email: c.WorkEmail || '', department: c.Department || '', jobTitle: c.JobTitle || '',
      location: c.BaseOfficeLocation || c.OfficeNumber || '', picture: c.PictureURL || '', skills: c.Skills || '', profileUrl: c.Path || ''
    }));
  }

  public async searchSites(query: string): Promise<ISearchDocument[]> {
    const rows = await this.search(buildKql(query, '(contentclass:STS_Site OR contentclass:STS_Web)'), 'Title,Path,Author,LastModifiedTime');
    return rows.map(c => ({ title: c.Title || 'SharePoint site', path: c.Path || '', author: c.Author || '', modified: c.LastModifiedTime || '', fileType: 'Site' }));
  }
}
