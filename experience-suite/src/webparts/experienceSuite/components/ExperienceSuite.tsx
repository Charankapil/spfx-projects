import * as React from 'react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { ApprovalDecision, IExperienceSuiteProps, ICurrentUser, IListItem, IPerson, ISearchDocument, ModuleKey, TaskStatus } from '../models/Models';
import { LISTS, SharePointService } from '../services/SharePointService';
import { approvalAssignee, isPending, sameEmail } from '../utils/helpers';
import { Empty, ItemRow, ModuleHeader, PanelTitle, PersonCard, ResultsPanel, SearchBar, Stat } from './Parts';
import './ExperienceSuite.css';

type LoadState = 'loading' | 'ready' | 'error';

const ALL_MODULES: { key: ModuleKey; label: string; icon: string; hint: string }[] = [
  { key: 'home', label: 'MySpace', icon: '⌂', hint: 'Your workday at a glance' },
  { key: 'requests', label: 'RequestHub', icon: '↗', hint: 'Employee service requests' },
  { key: 'documents', label: 'DocuSmart', icon: '▤', hint: 'Find trusted documents' },
  { key: 'sites', label: 'SiteNavigator', icon: '◫', hint: 'Discover SharePoint sites' },
  { key: 'approvals', label: 'ApprovalFlow', icon: '✓', hint: 'Items awaiting your decision' },
  { key: 'team', label: 'TeamPulse', icon: '◎', hint: 'Team news and action items' },
  { key: 'people', label: 'PeopleConnect', icon: '♙', hint: 'Find people and expertise' }
];

const REQUEST_CATEGORIES = ['IT support', 'SharePoint access', 'Equipment', 'Facilities', 'HR', 'Software access', 'Other'];
const TASK_STATUSES: TaskStatus[] = ['Not Started', 'In Progress', 'Completed'];

const greeting = (): string => {
  const h = new Date().getHours();
  return h < 12 ? 'Good morning' : h < 17 ? 'Good afternoon' : 'Good evening';
};

export const ExperienceSuite: React.FC<IExperienceSuiteProps> = ({ title, context, enablePeopleSearch }) => {
  const service = useMemo(() => new SharePointService(context), [context]);
  const modules = useMemo(() => ALL_MODULES.filter(m => m.key !== 'people' || enablePeopleSearch), [enablePeopleSearch]);

  const [active, setActive] = useState<ModuleKey>('home');
  const [user, setUser] = useState<ICurrentUser | undefined>();
  const [requests, setRequests] = useState<IListItem[]>([]);
  const [approvals, setApprovals] = useState<IListItem[]>([]);
  const [announcements, setAnnouncements] = useState<IListItem[]>([]);
  const [tasks, setTasks] = useState<IListItem[]>([]);
  const [people, setPeople] = useState<IPerson[]>([]);
  const [documents, setDocuments] = useState<ISearchDocument[]>([]);
  const [sites, setSites] = useState<ISearchDocument[]>([]);
  const [searched, setSearched] = useState<Partial<Record<ModuleKey, boolean>>>({});
  const [query, setQuery] = useState('');
  const [state, setState] = useState<LoadState>('loading');
  const [loadError, setLoadError] = useState('');
  const [listWarnings, setListWarnings] = useState<string[]>([]);
  const [error, setError] = useState('');
  const [busy, setBusy] = useState(false);
  const [showRequestForm, setShowRequestForm] = useState(false);
  const [requestTitle, setRequestTitle] = useState('');
  const [requestDescription, setRequestDescription] = useState('');
  const [requestCategory, setRequestCategory] = useState(REQUEST_CATEGORIES[0]);
  const [notice, setNotice] = useState('');
  const [actionId, setActionId] = useState<number | undefined>();
  const loadSeq = useRef(0);

  /** Loads the user and all four lists. `quiet` keeps the current view on screen (used after an action). */
  const refresh = useCallback(async (quiet = false): Promise<void> => {
    const seq = ++loadSeq.current;
    if (!quiet) setState('loading');
    setLoadError('');
    try {
      const [me, req, appr, news, work] = await Promise.all([
        service.getCurrentUser(),
        service.loadList(LISTS.requests),
        service.loadList(LISTS.approvals),
        service.loadList(LISTS.announcements),
        service.loadList(LISTS.tasks)
      ]);
      if (seq !== loadSeq.current) return; // a newer refresh has started; drop this stale result
      setUser(me);
      setRequests(req.items); setApprovals(appr.items); setAnnouncements(news.items); setTasks(work.items);
      setListWarnings([req, appr, news, work].map(r => r.error).filter((e): e is string => !!e));
      setState('ready');
    } catch (e) {
      if (seq !== loadSeq.current) return;
      setLoadError((e as Error).message);
      setState('error');
    }
  }, [service]);

  useEffect(() => { refresh().catch(() => undefined); }, [refresh]);

  // If PeopleConnect is switched off while it is open, fall back to the landing view.
  useEffect(() => { if (!enablePeopleSearch && active === 'people') setActive('home'); }, [enablePeopleSearch, active]);

  const runSearch = async (): Promise<void> => {
    if (busy) return;
    const target = active;
    setBusy(true); setError(''); setNotice('');
    try {
      if (target === 'people' && enablePeopleSearch) setPeople(await service.searchPeople(query));
      else if (target === 'documents') setDocuments(await service.searchDocuments(query));
      else if (target === 'sites') setSites(await service.searchSites(query));
      setSearched(s => ({ ...s, [target]: true }));
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const nav = (key: ModuleKey): void => { setActive(key); setQuery(''); setError(''); setNotice(''); };

  const myEmail = user?.Email || '';
  const displayName = user?.Title || 'Colleague';
  const myRequests = requests.filter(x => x.RequesterEmail ? sameEmail(x.RequesterEmail, myEmail) : !!user && x.Author?.Id === user.Id);
  const myApprovals = approvals.filter(x => sameEmail(approvalAssignee(x), myEmail));
  const pending = myApprovals.filter(isPending).length;
  const myTasks = tasks.filter(x => sameEmail(x.Owner, myEmail));

  const submitRequest = async (): Promise<void> => {
    if (busy || !requestTitle.trim() || !requestDescription.trim() || !user) return;
    setBusy(true); setError(''); setNotice('');
    try {
      await service.createRequest({ title: requestTitle, description: requestDescription, category: requestCategory, requesterName: user.Title, requesterEmail: user.Email });
      setRequestTitle(''); setRequestDescription(''); setRequestCategory(REQUEST_CATEGORIES[0]); setShowRequestForm(false);
      setNotice('Your request was submitted successfully.');
      await refresh(true);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const decideApproval = async (item: IListItem, decision: ApprovalDecision): Promise<void> => {
    if (!user || actionId !== undefined) return;
    setActionId(item.Id); setError(''); setNotice('');
    try {
      await service.setApprovalStatus(item, decision, user.Email);
      setNotice(`Approval ${decision.toLowerCase()}.`);
      await refresh(true);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setActionId(undefined);
    }
  };

  const updateTask = async (item: IListItem, status: TaskStatus): Promise<void> => {
    if (actionId !== undefined) return;
    setActionId(item.Id); setError(''); setNotice('');
    try {
      const etag = await service.setTaskStatus(item, status);
      // keep the new ETag so a second change to the same task is not rejected as a conflict
      setTasks(current => current.map(x => x.Id === item.Id ? { ...x, Status: status, etag: etag || x.etag } : x));
      setNotice('Task status updated.');
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setActionId(undefined);
    }
  };

  const initial = (displayName.charAt(0) || '?').toUpperCase();

  return (
    <div className="es-shell">
      <aside className="es-sidebar">
        <div className="es-brand">
          <div className="es-brand-mark" aria-hidden="true">E</div>
          <div><strong>{title}</strong><small>SharePoint workspace</small></div>
        </div>
        <div className="es-nav-label">WORKSPACE</div>
        <nav aria-label="Experience Suite modules">
          {modules.map(m => (
            <button key={m.key} className={`es-nav-item ${active === m.key ? 'active' : ''}`} aria-current={active === m.key ? 'page' : undefined} onClick={() => nav(m.key)}>
              <span className="es-nav-icon" aria-hidden="true">{m.icon}</span>
              <span>{m.label}</span>
              {m.key === 'approvals' && pending > 0 && <b className="es-count" aria-label={`${pending} pending`}>{pending}</b>}
            </button>
          ))}
        </nav>
        <div className="es-sidebar-bottom">
          <div className="es-avatar" aria-hidden="true">{initial}</div>
          <div className="es-user-mini"><strong>{displayName}</strong><small>{myEmail || 'Signed-in user'}</small></div>
        </div>
      </aside>

      <main className="es-main">
        <header className="es-topbar">
          <div className="es-crumb">Workspace <span>/</span> {modules.find(m => m.key === active)?.label}</div>
          <div className="es-top-actions">
            <span className="es-online"><i /> SharePoint connected</span>
            <button className="es-icon-button" title="Refresh" aria-label="Refresh" onClick={() => { refresh().catch(() => undefined); }}>↻</button>
            <div className="es-avatar small" aria-hidden="true">{initial}</div>
          </div>
        </header>

        <section className="es-content">
          {state === 'loading' && <div className="es-loading" role="status"><div className="es-spinner" />Loading your workspace…</div>}
          {state === 'error' && (
            <div className="es-alert" role="alert">
              <strong>Couldn&apos;t load your SharePoint profile.</strong>
              <span>{loadError}</span>
              <button onClick={() => { refresh().catch(() => undefined); }}>Try again</button>
            </div>
          )}

          {state === 'ready' && <>
            {listWarnings.length > 0 && (
              <div className="es-inline-error es-list-warning" role="alert">
                <strong>Some data could not be loaded:</strong>
                {listWarnings.map(w => <span key={w}>{w}</span>)}
              </div>
            )}

            {active === 'home' && <>
              <div className="es-hero">
                <div>
                  <div className="es-eyebrow">YOUR WORKSPACE</div>
                  <h1>{greeting()}, {displayName.split(' ')[0]}.</h1>
                  <p>Here’s what’s happening across your work today.</p>
                </div>
                <div className="es-hero-orb" aria-hidden="true">✦</div>
              </div>
              <div className="es-stat-grid">
                <Stat label="My requests" value={myRequests.length} note="Your submitted requests" icon="↗" />
                <Stat label="Pending approvals" value={pending} note="Waiting for your decision" icon="✓" />
                <Stat label="Team updates" value={announcements.length} note="Published announcements" icon="◉" />
                <Stat label="My actions" value={myTasks.filter(t => t.Status !== 'Completed').length} note="Open tasks assigned to you" icon="▣" />
              </div>
              <div className="es-home-grid">
                <section className="es-panel">
                  <PanelTitle title="Quick access" subtitle="Jump back into the tools you use most" />
                  <div className="es-quick-grid">
                    {modules.slice(1).map(m => (
                      <button key={m.key} className="es-quick" onClick={() => nav(m.key)}>
                        <span aria-hidden="true">{m.icon}</span><strong>{m.label}</strong><small>{m.hint}</small><b aria-hidden="true">→</b>
                      </button>
                    ))}
                  </div>
                </section>
                <section className="es-panel">
                  <PanelTitle title="Latest team updates" subtitle="Recent announcements from your workspace" />
                  {announcements.length
                    ? announcements.slice(0, 5).map(x => <ItemRow key={x.Id} item={x} />)
                    : <Empty text="No announcements found. Add items to the ES Announcements list to show team news here." />}
                </section>
              </div>
            </>}

            {active === 'requests' && <>
              <ModuleHeader title="RequestHub" subtitle="Raise a request and track its progress." icon="↗">
                <button className="es-primary" onClick={() => setShowRequestForm(v => !v)}>{showRequestForm ? 'Close form' : '＋ New request'}</button>
              </ModuleHeader>
              {showRequestForm && (
                <section className="es-panel es-form-panel">
                  <PanelTitle title="Create a service request" subtitle="Your request will be recorded in SharePoint." />
                  <label>Request title
                    <input value={requestTitle} onChange={e => setRequestTitle(e.target.value)} maxLength={255} placeholder="What do you need help with?" />
                  </label>
                  <label>Category
                    <select value={requestCategory} onChange={e => setRequestCategory(e.target.value)}>
                      {REQUEST_CATEGORIES.map(c => <option key={c}>{c}</option>)}
                    </select>
                  </label>
                  <label>Description
                    <textarea value={requestDescription} onChange={e => setRequestDescription(e.target.value)} rows={4} maxLength={4000} placeholder="Add details that will help the team resolve your request." />
                  </label>
                  <div className="es-form-actions">
                    <button className="es-primary" onClick={() => { submitRequest().catch(() => undefined); }} disabled={busy || !requestTitle.trim() || !requestDescription.trim()}>
                      {busy ? 'Submitting…' : 'Submit request'}
                    </button>
                    <small>Submitted as {myEmail}</small>
                  </div>
                </section>
              )}
              <section className="es-panel">
                <PanelTitle title="My service requests" subtitle={`${myRequests.length} requests submitted by you`} />
                {myRequests.length
                  ? myRequests.map(x => <ItemRow key={x.Id} item={x} />)
                  : <Empty text="You have not submitted any requests yet. Use New request to get started." />}
              </section>
            </>}

            {active === 'documents' && <>
              <ModuleHeader title="DocuSmart" subtitle="Search documents you already have permission to access." icon="▤" />
              <SearchBar query={query} setQuery={setQuery} onSearch={() => { runSearch().catch(() => undefined); }} busy={busy} placeholder="Search policies, guides, templates…" />
              <ResultsPanel searched={!!searched.documents} empty="Search SharePoint to discover documents." items={documents} />
            </>}

            {active === 'sites' && <>
              <ModuleHeader title="SiteNavigator" subtitle="Discover SharePoint sites and workspaces you can access." icon="◫" />
              <SearchBar query={query} setQuery={setQuery} onSearch={() => { runSearch().catch(() => undefined); }} busy={busy} placeholder="Search a team, department or site…" />
              <ResultsPanel searched={!!searched.sites} empty="Search for a SharePoint site or workspace." items={sites} />
            </>}

            {active === 'approvals' && <>
              <ModuleHeader title="ApprovalFlow" subtitle="Review and decide on requests assigned to you." icon="✓" />
              <section className="es-panel">
                <PanelTitle title="My approval queue" subtitle={`${pending} pending · ${myApprovals.length} assigned records`} />
                {myApprovals.length
                  ? myApprovals.map(x => (
                    <div className="es-approval-row" key={x.Id}>
                      <ItemRow item={x} />
                      {isPending(x) && (
                        <div className="es-row-actions">
                          <button className="es-approve" disabled={actionId !== undefined} onClick={() => { decideApproval(x, 'Approved').catch(() => undefined); }}>Approve</button>
                          <button className="es-reject" disabled={actionId !== undefined} onClick={() => { decideApproval(x, 'Rejected').catch(() => undefined); }}>Reject</button>
                        </div>
                      )}
                    </div>
                  ))
                  : <Empty text="No approvals are currently assigned to your account." />}
              </section>
            </>}

            {active === 'team' && <>
              <ModuleHeader title="TeamPulse" subtitle="Team announcements and your action items." icon="◎" />
              <div className="es-two-col">
                <section className="es-panel">
                  <PanelTitle title="Announcements" subtitle="Latest team news" />
                  {announcements.length ? announcements.map(x => <ItemRow key={x.Id} item={x} />) : <Empty text="No announcements published yet." />}
                </section>
                <section className="es-panel">
                  <PanelTitle title="Action items" subtitle="Update tasks assigned to you" />
                  {myTasks.length
                    ? myTasks.map(x => (
                      <div className="es-task-row" key={x.Id}>
                        <ItemRow item={x} />
                        <select
                          aria-label={`Update ${x.Title}`}
                          value={x.Status || 'Not Started'}
                          disabled={actionId !== undefined}
                          onChange={e => { updateTask(x, e.target.value as TaskStatus).catch(() => undefined); }}
                        >
                          {TASK_STATUSES.map(s => <option key={s}>{s}</option>)}
                        </select>
                      </div>
                    ))
                    : <Empty text="No tasks are assigned to you." />}
                </section>
              </div>
            </>}

            {active === 'people' && enablePeopleSearch && <>
              <ModuleHeader title="PeopleConnect" subtitle="Find colleagues by name, role, department or expertise." icon="♙" />
              <div className="es-people-search">
                <SearchBar query={query} setQuery={setQuery} onSearch={() => { runSearch().catch(() => undefined); }} busy={busy} placeholder="Try ‘SharePoint’, ‘Finance’ or a colleague’s name" buttonText="Find people" />
                <small>Search results respect the signed-in user&apos;s SharePoint search permissions and available user profile properties.</small>
              </div>
              <div className="es-people-grid">{people.map((p, i) => <PersonCard key={`${p.email}-${i}`} person={p} />)}</div>
              {people.length === 0 && <Empty text={searched.people ? 'No people matched your search.' : 'Search for a person, skill or department to discover colleagues.'} />}
            </>}

            {notice && <div className="es-success" role="status">{notice}</div>}
            {error && <div className="es-inline-error" role="alert">{error}</div>}
          </>}
        </section>
        <footer className="es-footer">SharePoint Experience Suite <span>·</span> Uses your existing SharePoint sign-in and permissions</footer>
      </main>
    </div>
  );
};
