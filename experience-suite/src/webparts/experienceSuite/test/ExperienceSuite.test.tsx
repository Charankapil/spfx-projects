import * as React from 'react';
import { createRoot, Root } from 'react-dom/client';
import { ExperienceSuite } from '../components/ExperienceSuite';

// React 18.3 exports act() from 'react'; the pinned @types/react (18.2) does not declare it yet.
const act = (React as unknown as { act: (cb: () => void | Promise<void>) => Promise<void> }).act;

// Tell React this is a test environment so act() works without warnings.
(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

interface IFakeCall { url: string; method: string; body?: string; }

const ME = { Id: 10, Title: 'Ann Smith', Email: 'ann@contoso.com', LoginName: 'i:0#.f|membership|ann@contoso.com' };
const DATA: Record<string, unknown[]> = {
  'ES Requests': [
    { Id: 1, Title: 'Laptop', Status: 'Submitted', RequesterEmail: 'ANN@contoso.com' },
    { Id: 2, Title: 'Someone else', Status: 'Submitted', RequesterEmail: 'bob@contoso.com' }
  ],
  'ES Approvals': [
    { Id: 3, Title: 'Invoice 42', Status: 'Pending', ApproverEmail: 'ann@contoso.com', '@odata.etag': '"1"' },
    { Id: 4, Title: 'Not mine', Status: 'Pending', ApproverEmail: 'bob@contoso.com' },
    { Id: 5, Title: 'Unassigned', Status: 'Pending' }
  ],
  'ES Announcements': [{ Id: 6, Title: 'Town hall Friday' }],
  'ES Tasks': [
    { Id: 7, Title: 'My task', Status: 'In Progress', Owner: 'ann@contoso.com', '@odata.etag': '"2"' },
    { Id: 8, Title: 'Bob task', Status: 'Not Started', Owner: 'bob@contoso.com' }
  ]
};

function respond(status: number, body?: unknown, headers: Record<string, string> = {}): unknown {
  return {
    ok: status < 300, status, statusText: String(status),
    headers: { get: (n: string) => headers[n] || null },
    json: () => Promise.resolve(body)
  };
}

function fakeContext(calls: IFakeCall[], overrides: { failList?: string } = {}): unknown {
  return {
    pageContext: { web: { absoluteUrl: 'https://contoso.sharepoint.com/sites/hub' } },
    spHttpClient: {
      fetch: (url: string, _c: unknown, o: { method?: string; body?: string }) => {
        calls.push({ url, method: o.method || 'GET', body: o.body });
        if (url.indexOf('/currentuser') >= 0) return Promise.resolve(respond(200, ME));
        const list = /getbytitle\('([^']+)'\)/.exec(decodeURIComponent(url));
        if (list && o.method === 'POST') return Promise.resolve(respond(url.indexOf('/items(') >= 0 ? 204 : 201, { Id: 99 }, { ETag: '"3"' }));
        if (list) {
          if (list[1] === overrides.failList) return Promise.resolve(respond(403, { 'odata.error': { message: { value: 'Access denied' } } }));
          return Promise.resolve(respond(200, { value: DATA[list[1]] || [] }));
        }
        if (url.indexOf('/_api/search/') >= 0) {
          return Promise.resolve(respond(200, { PrimaryQueryResult: { RelevantResults: { Table: { Rows: [
            { Cells: [{ Key: 'Title', Value: 'Result A' }, { Key: 'Path', Value: 'https://contoso.sharepoint.com/a' }] },
            // eslint-disable-next-line no-script-url
            { Cells: [{ Key: 'Title', Value: 'Evil' }, { Key: 'Path', Value: 'javascript:alert(1)' }] }
          ] } } } }));
        }
        return Promise.resolve(respond(404, {}));
      }
    }
  };
}

let container: HTMLDivElement;
let root: Root;

async function mount(ctx: unknown, enablePeopleSearch = true): Promise<void> {
  await act(async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    root.render(<ExperienceSuite title="Suite" context={ctx as any} enablePeopleSearch={enablePeopleSearch} />);
  });
}

const navButton = (label: string): HTMLButtonElement | undefined =>
  Array.from(container.querySelectorAll<HTMLButtonElement>('.es-nav-item')).filter(b => (b.textContent || '').indexOf(label) >= 0)[0];

async function click(el: Element | undefined): Promise<void> {
  if (!el) throw new Error('element not found');
  await act(async () => { (el as HTMLElement).click(); });
}

beforeEach(() => {
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
});

afterEach(async () => {
  await act(async () => root.unmount());
  container.remove();
});

describe('ExperienceSuite UI', () => {
  it('loads the user and shows only their own data on MySpace', async () => {
    const calls: IFakeCall[] = [];
    await mount(fakeContext(calls));
    expect(container.textContent).toContain('Ann Smith');
    const stats = Array.from(container.querySelectorAll('.es-stat strong')).map(s => s.textContent);
    expect(stats).toEqual(['1', '1', '1', '1']); // my requests, my pending approvals, announcements, my open tasks
    expect(container.textContent).toContain('Town hall Friday');
  });

  it('ApprovalFlow lists only approvals assigned to me and approves with the ETag', async () => {
    const calls: IFakeCall[] = [];
    await mount(fakeContext(calls));
    await click(navButton('ApprovalFlow'));
    expect(container.textContent).toContain('Invoice 42');
    expect(container.textContent).not.toContain('Not mine');
    expect(container.textContent).not.toContain('Unassigned');
    await click(container.querySelector('.es-approve') || undefined);
    const update = calls.filter(c => c.method === 'POST')[0];
    expect(update.url).toContain('/items(3)');
    expect(JSON.parse(update.body!)).toEqual(expect.objectContaining({ Status: 'Approved', DecisionBy: 'ann@contoso.com' }));
    expect(container.textContent).toContain('Approval approved.');
  });

  it('TeamPulse shows only my tasks', async () => {
    await mount(fakeContext([]));
    await click(navButton('TeamPulse'));
    expect(container.textContent).toContain('My task');
    expect(container.textContent).not.toContain('Bob task');
    expect(container.querySelector('.es-status.in-progress')).not.toBeNull();
  });

  it('DocuSmart renders results and neutralises non-http links', async () => {
    await mount(fakeContext([]));
    await click(navButton('DocuSmart'));
    await click(Array.from(container.querySelectorAll('button')).filter(b => b.textContent === 'Search')[0]);
    const links = Array.from(container.querySelectorAll<HTMLAnchorElement>('a.es-result'));
    expect(links).toHaveLength(2);
    expect(links[0].getAttribute('href')).toBe('https://contoso.sharepoint.com/a');
    expect(links[1].getAttribute('href')).toBeNull();
  });

  it('keeps working when one list is not accessible', async () => {
    await mount(fakeContext([], { failList: 'ES Approvals' }));
    expect(container.textContent).toContain('Some data could not be loaded');
    expect(container.textContent).toContain('ES Approvals');
    expect(container.textContent).toContain('Town hall Friday');
  });

  it('hides PeopleConnect when disabled', async () => {
    await mount(fakeContext([]), false);
    expect(navButton('PeopleConnect')).toBeUndefined();
  });

  it('submits a new request with the signed-in user as requester', async () => {
    const calls: IFakeCall[] = [];
    await mount(fakeContext(calls));
    await click(navButton('RequestHub'));
    await click(Array.from(container.querySelectorAll('button')).filter(b => (b.textContent || '').indexOf('New request') >= 0)[0]);
    const setValue = (el: HTMLInputElement | HTMLTextAreaElement, value: string): void => {
      const proto = el instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, 'value')!.set!.call(el, value);
      el.dispatchEvent(new Event('input', { bubbles: true }));
    };
    await act(async () => {
      setValue(container.querySelector('.es-form-panel input') as HTMLInputElement, 'New monitor');
      setValue(container.querySelector('.es-form-panel textarea') as HTMLTextAreaElement, 'Please');
    });
    await click(Array.from(container.querySelectorAll('button')).filter(b => b.textContent === 'Submit request')[0]);
    const create = calls.filter(c => c.method === 'POST')[0];
    expect(JSON.parse(create.body!)).toEqual(expect.objectContaining({ Title: 'New monitor', Status: 'Submitted', RequesterEmail: 'ann@contoso.com' }));
    expect(container.textContent).toContain('Your request was submitted successfully.');
  });
});
