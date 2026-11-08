import { LISTS, SharePointService } from '../services/SharePointService';
import { IListItem } from '../models/Models';

interface ICall { url: string; options: { method?: string; headers?: Record<string, string>; body?: string }; }

function fakeResponse(status: number, body?: unknown, headers: Record<string, string> = {}): unknown {
  return {
    ok: status >= 200 && status < 300,
    status,
    statusText: `HTTP ${status}`,
    headers: { get: (name: string) => headers[name] || headers[name.toLowerCase()] || null },
    json: () => body === undefined ? Promise.reject(new Error('no body')) : Promise.resolve(body)
  };
}

function createService(responder: (call: ICall) => unknown): { service: SharePointService; calls: ICall[] } {
  const calls: ICall[] = [];
  const context = {
    pageContext: { web: { absoluteUrl: 'https://contoso.sharepoint.com/sites/hub' } },
    spHttpClient: {
      fetch: (url: string, _config: unknown, options: ICall['options']) => {
        const call = { url, options };
        calls.push(call);
        return Promise.resolve(responder(call));
      }
    }
  };
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return { service: new SharePointService(context as any), calls };
}

const selectOf = (url: string): string[] => decodeURIComponent(/\$select=([^&]+)/.exec(url)![1]).split(',');

describe('SharePointService list reads', () => {
  it('only selects columns that exist on each list (unknown columns make SharePoint return 400)', async () => {
    const { service, calls } = createService(() => fakeResponse(200, { value: [] }));
    await service.getListItems(LISTS.announcements);
    await service.getListItems(LISTS.tasks);
    await service.getListItems(LISTS.requests);
    await service.getListItems(LISTS.approvals);
    const [ann, tasks, req, appr] = calls.map(c => selectOf(c.url));
    expect(ann).not.toEqual(expect.arrayContaining(['Status']));
    for (const f of ['RequesterEmail', 'ApproverEmail', 'DueDate']) expect(ann).not.toContain(f);
    for (const f of ['RequesterEmail', 'ApproverEmail']) expect(tasks).not.toContain(f);
    expect(req).toContain('RequesterEmail');
    expect(req).not.toContain('ApproverEmail');
    expect(appr).toContain('ApproverEmail');
    expect(appr).not.toContain('RequesterEmail');
  });

  it('maps the OData ETag onto the item', async () => {
    const { service } = createService(() => fakeResponse(200, { value: [{ Id: 4, Title: 'x', '@odata.etag': '"3"' }] }));
    const items = await service.getListItems(LISTS.tasks);
    expect(items[0]).toEqual({ Id: 4, Title: 'x', etag: '"3"' });
  });

  it('treats a missing list (404) as empty but reports other failures', async () => {
    const missing = createService(() => fakeResponse(404, { 'odata.error': { message: { value: 'List does not exist' } } }));
    expect(await missing.service.loadList(LISTS.tasks)).toEqual({ items: [] });

    const denied = createService(() => fakeResponse(403, { 'odata.error': { message: { value: 'Access denied' } } }));
    const result = await denied.service.loadList(LISTS.tasks);
    expect(result.items).toEqual([]);
    expect(result.error).toContain('(403)');
    expect(result.error).toContain('Access denied');
  });

  it('escapes apostrophes in list titles', async () => {
    const { service, calls } = createService(() => fakeResponse(200, { value: [] }));
    await service.getListItems("O'Brien");
    expect(calls[0].url).toContain("getbytitle('O''Brien')");
  });
});

describe('SharePointService approvals', () => {
  const pending: IListItem = { Id: 7, Title: 'PO', Status: 'Pending', ApproverEmail: 'Ann@contoso.com', etag: '"5"' };

  it('updates status with the item ETag', async () => {
    const { service, calls } = createService(() => fakeResponse(204, undefined, { ETag: '"6"' }));
    await service.setApprovalStatus(pending, 'Approved', 'ann@contoso.com');
    expect(calls).toHaveLength(1);
    expect(calls[0].options.headers!['IF-MATCH']).toBe('"5"');
    expect(calls[0].options.headers!['X-HTTP-Method']).toBe('MERGE');
    expect(JSON.parse(calls[0].options.body!)).toEqual(expect.objectContaining({ Status: 'Approved', DecisionBy: 'ann@contoso.com' }));
  });

  it('refuses items assigned to someone else', async () => {
    const { service, calls } = createService(() => fakeResponse(204));
    await expect(service.setApprovalStatus(pending, 'Approved', 'bob@contoso.com')).rejects.toThrow('not assigned to you');
    expect(calls).toHaveLength(0);
  });

  it('refuses unassigned items (previously anyone could decide them)', async () => {
    const { service, calls } = createService(() => fakeResponse(204));
    await expect(service.setApprovalStatus({ ...pending, ApproverEmail: undefined }, 'Rejected', 'bob@contoso.com')).rejects.toThrow();
    expect(calls).toHaveLength(0);
  });

  it('refuses items that are no longer pending', async () => {
    const { service } = createService(() => fakeResponse(204));
    await expect(service.setApprovalStatus({ ...pending, Status: 'Approved' }, 'Rejected', 'ann@contoso.com')).rejects.toThrow('no longer pending');
  });

  it('turns a 412 into a friendly concurrency message', async () => {
    const { service } = createService(() => fakeResponse(412, {}));
    await expect(service.setApprovalStatus(pending, 'Approved', 'ann@contoso.com')).rejects.toThrow('changed by someone else');
  });
});

describe('SharePointService tasks', () => {
  it('returns the new ETag after an update', async () => {
    const { service } = createService(() => fakeResponse(204, undefined, { ETag: '"9"' }));
    expect(await service.setTaskStatus({ Id: 1, Title: 't', etag: '"8"' }, 'Completed')).toBe('"9"');
  });
});

describe('SharePointService search', () => {
  const searchBody = { PrimaryQueryResult: { RelevantResults: { Table: { Rows: [{ Cells: [{ Key: 'Title', Value: 'Ann' }, { Key: 'WorkEmail', Value: 'ann@contoso.com' }, { Key: 'BaseOfficeLocation', Value: 'Leeds' }] }] } } } };

  it('queries people through the Local People Results source', async () => {
    const { service, calls } = createService(() => fakeResponse(200, searchBody));
    const people = await service.searchPeople('');
    expect(calls[0].url).toContain("sourceid='b09a7990-05ea-4af9-81ef-edfab16c4e31'");
    expect(decodeURIComponent(calls[0].url)).toContain("querytext='*'");
    expect(people[0]).toEqual(expect.objectContaining({ title: 'Ann', email: 'ann@contoso.com', location: 'Leeds' }));
  });

  it('restricts documents to IsDocument:1 and escapes quotes', async () => {
    const { service, calls } = createService(() => fakeResponse(200, { PrimaryQueryResult: {} }));
    expect(await service.searchDocuments("bob's plan")).toEqual([]);
    expect(decodeURIComponent(calls[0].url)).toContain("querytext='(bob''s plan) AND IsDocument:1'");
  });

  it('restricts sites to site and web content classes', async () => {
    const { service, calls } = createService(() => fakeResponse(200, { PrimaryQueryResult: {} }));
    await service.searchSites('');
    expect(decodeURIComponent(calls[0].url)).toContain("querytext='(contentclass:STS_Site OR contentclass:STS_Web)'");
  });
});
