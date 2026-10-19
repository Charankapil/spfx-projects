import * as React from 'react';
import { DefaultButton, DetailsList, DetailsListLayoutMode, Dropdown, IColumn, MessageBar, MessageBarType, Pivot, PivotItem, PrimaryButton, SearchBox, Selection, SelectionMode, TextField } from '@fluentui/react';
import { parseAddresses } from '../../services/addresses';
import { downloadCsv } from '../../services/exportCsv';
import { formatDate } from '../../services/format';
import { IDirectoryObject, IM365Group } from '../../services/GraphApi';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { GraphGate, graphErrorText } from '../shared/graph';
import { useLoader } from '../shared/hooks';
import { Card, Empty, ErrorBar, Kpi, Loading, Pill, ViewHeader } from '../shared/ui';

const MAX_BULK = 100;

/**
 * Microsoft 365 groups behind team sites (Microsoft Graph, Group.ReadWrite.All + User.Read.All).
 * SharePoint's own "Owners" group of a group-connected site is synced from the Microsoft 365
 * group, so owners and members are managed here, where the change actually sticks.
 */
export const GroupsView: React.FC = () => {
  return (
    <div className={styles.view}>
      <ViewHeader title="Microsoft 365 groups & owners" hint="Owners and members of the Microsoft 365 group behind a team site, and groups across the tenant that have nobody in charge." />
      <GraphGate scopes={['Group.ReadWrite.All', 'User.Read.All']} feature="Microsoft 365 group management">
        <Pivot aria-label="Group views">
          <PivotItem headerText="This site's group">
            <div style={{ paddingTop: 12 }}>
              <SiteGroupTab />
            </div>
          </PivotItem>
          <PivotItem headerText="Ownerless groups (tenant)">
            <div style={{ paddingTop: 12 }}>
              <OwnerlessTab />
            </div>
          </PivotItem>
        </Pivot>
      </GraphGate>
    </div>
  );
};

// ---- the selected site's group ---------------------------------------------------

const SiteGroupTab: React.FC = () => {
  const ctx = useAdmin();
  const { api, graph, target } = ctx;
  const loader = useLoader(async () => {
    const site = await api.getSite(target.siteUrl);
    if (!site.groupId || /^0{8}-/.test(site.groupId)) {
      return { group: undefined };
    }
    const [group, owners, members] = await Promise.all([graph.getGroup(site.groupId), graph.getGroupPeople(site.groupId, 'owners'), graph.getGroupPeople(site.groupId, 'members')]);
    return { group, owners: owners.people, members: members.people, truncated: owners.truncated || members.truncated };
  }, [target.siteUrl]);

  if (loader.loading && !loader.data) {
    return <Loading text="Loading the Microsoft 365 group…" />;
  }
  if (loader.error || !loader.data) {
    return <ErrorBar error={loader.error || 'Nothing loaded.'} onRetry={loader.reload} />;
  }
  const { group, owners, members, truncated } = loader.data;
  if (!group || !owners || !members) {
    return (
      <Card>
        <Empty text={`${target.title} is not connected to a Microsoft 365 group, so its people are managed in People & permissions.`} />
      </Card>
    );
  }
  const userOwners = owners.filter((o) => o.type === 'user');
  return (
    <>
      <div className={styles.kpiRow}>
        <Kpi label="Group" value={group.displayName} sub={group.mail} />
        <Kpi label="Owners" value={String(userOwners.length)} sub={userOwners.length < 2 ? 'add a second owner' : undefined} />
        <Kpi label="Members" value={members.length.toLocaleString()} sub={`${members.filter((m) => (m.userType || '').toLowerCase() === 'guest').length} guests`} />
        <Kpi label="Privacy" value={group.visibility || '—'} sub={group.isTeam ? 'Microsoft Teams team' : 'no team'} />
      </div>
      {userOwners.length === 0 && <MessageBar messageBarType={MessageBarType.severeWarning}>This group has no owners. Nobody can manage its membership or respond to access requests. Add at least two owners.</MessageBar>}
      {userOwners.length === 1 && <MessageBar messageBarType={MessageBarType.warning}>Only one owner. If that person leaves, the group and the site are left without an owner.</MessageBar>}
      {truncated && <MessageBar messageBarType={MessageBarType.info}>Very large group: the first 5,000 entries are shown.</MessageBar>}
      <div className={styles.grid}>
        <PeopleCard groupId={group.id} kind="owners" people={owners} onChanged={loader.reload} />
        <PeopleCard groupId={group.id} kind="members" people={members} onChanged={loader.reload} />
      </div>
    </>
  );
};

const PeopleCard: React.FC<{ groupId: string; kind: 'owners' | 'members'; people: IDirectoryObject[]; onChanged: () => void }> = ({ groupId, kind, people, onChanged }) => {
  const ctx = useAdmin();
  const [input, setInput] = React.useState('');
  const [busy, setBusy] = React.useState(false);
  const [filter, setFilter] = React.useState('');
  const label = kind === 'owners' ? 'owner' : 'member';
  const userOwners = people.filter((p) => p.type === 'user');

  const add = async (): Promise<void> => {
    const addresses = parseAddresses(input).slice(0, MAX_BULK);
    if (!addresses.length) {
      return;
    }
    setBusy(true);
    let ok = 0;
    const failed: string[] = [];
    for (const a of addresses) {
      try {
        const user = await ctx.graph.findUser(a);
        if (!user) {
          failed.push(`${a} (not found in the directory)`);
          continue;
        }
        if (people.some((p) => p.id === user.id)) {
          ok++;
          continue;
        }
        await ctx.graph.addToGroup(groupId, kind, user.id);
        ctx.log(`Add group ${label}`, `${user.displayName} → ${groupId}`, true);
        ok++;
      } catch (e) {
        const msg = graphErrorText(e);
        failed.push(`${a} (${msg.substring(0, 120)})`);
        ctx.log(`Add group ${label}`, a, false, msg);
      }
    }
    setBusy(false);
    setInput('');
    ctx.notify(failed.length ? `${ok} added. Failed: ${failed.join('; ')}` : `${ok} ${label}${ok === 1 ? '' : 's'} added.`, failed.length ? 'error' : 'success');
    onChanged();
  };

  const remove = async (p: IDirectoryObject): Promise<void> => {
    if (kind === 'owners' && p.type === 'user' && userOwners.length <= 1) {
      ctx.notify('This is the last owner. Add another owner first so the group is never left without one.', 'error');
      return;
    }
    const self = [p.userPrincipalName, p.mail].some((x) => !!x && x.toLowerCase() === ctx.currentUser);
    const ok = await ctx.confirm({
      title: `Remove ${label}?`,
      message: `${p.displayName} will be removed as ${label} of the Microsoft 365 group${kind === 'members' ? ' and loses access to the site, the team and the mailbox' : ''}.${self ? ' This is you: you may lose the ability to manage this group.' : ''}`,
      confirmText: 'Remove',
      danger: true
    });
    if (!ok) {
      return;
    }
    setBusy(true);
    try {
      await ctx.graph.removeFromGroup(groupId, kind, p.id);
      ctx.log(`Remove group ${label}`, `${p.displayName} → ${groupId}`, true);
      ctx.notify(`${p.displayName} removed.`, 'success');
      onChanged();
    } catch (e) {
      ctx.log(`Remove group ${label}`, p.displayName, false, graphErrorText(e));
      ctx.notify(graphErrorText(e), 'error');
    } finally {
      setBusy(false);
    }
  };

  const f = filter.trim().toLowerCase();
  const shown = people.filter((p) => !f || (p.displayName + ' ' + (p.mail || '') + ' ' + (p.userPrincipalName || '')).toLowerCase().indexOf(f) >= 0).slice(0, 300);
  return (
    <Card title={kind === 'owners' ? `Owners (${people.length})` : `Members (${people.length})`}>
      <TextField
        label={`Add ${label}s`}
        placeholder="E-mail addresses, one per line or separated by ; or ,"
        multiline
        rows={2}
        value={input}
        onChange={(_, v) => setInput(v || '')}
        disabled={busy}
      />
      <div className={styles.actions} style={{ margin: '8px 0 12px' }}>
        <PrimaryButton onClick={add} disabled={busy || !input.trim()}>
          {busy ? 'Working…' : `Add ${label}s`}
        </PrimaryButton>
      </div>
      {people.length > 10 && <SearchBox placeholder={`Filter ${kind}`} value={filter} onChange={(_, v) => setFilter(v || '')} />}
      {people.length === 0 ? (
        <Empty text={`No ${kind}.`} />
      ) : (
        <ul className={styles.memberList}>
          {shown.map((p) => (
            <li key={p.id} className={styles.memberRow}>
              <span style={{ minWidth: 0 }}>
                <div>
                  {p.displayName} {(p.userType || '').toLowerCase() === 'guest' && <Pill kind="warning">Guest</Pill>} {p.type !== 'user' && <Pill kind="info">{p.type === 'servicePrincipal' ? 'App' : 'Group'}</Pill>}
                </div>
                <div className={styles.muted} style={{ fontSize: 12, overflow: 'hidden', textOverflow: 'ellipsis' }}>
                  {p.mail || p.userPrincipalName}
                </div>
              </span>
              <button type="button" className={styles.link} onClick={() => remove(p)} disabled={busy}>
                Remove
              </button>
            </li>
          ))}
        </ul>
      )}
    </Card>
  );
};

// ---- tenant-wide ownerless groups -----------------------------------------------

type OwnerFilter = 'none' | 'one' | 'all';

const OwnerlessTab: React.FC = () => {
  const ctx = useAdmin();
  const [state, setState] = React.useState<{ groups?: IM365Group[]; truncated?: boolean; loading: boolean; progress: number; error?: string; at?: Date }>({ loading: false, progress: 0 });
  const [show, setShow] = React.useState<OwnerFilter>('none');
  const [query, setQuery] = React.useState('');
  const [owner, setOwner] = React.useState('');
  const [busy, setBusy] = React.useState(false);
  const [selectedCount, setSelectedCount] = React.useState(0);
  const selection = React.useMemo(() => new Selection({ getKey: (g) => (g as IM365Group).id, onSelectionChanged: () => setSelectedCount(selection.getSelectedCount()) }), []);

  const scan = async (): Promise<void> => {
    setState({ loading: true, progress: 0 });
    try {
      const res = await ctx.graph.listGroupsWithOwners((n) => setState((s) => ({ ...s, progress: n })));
      setState({ loading: false, progress: res.groups.length, groups: res.groups, truncated: res.truncated, at: new Date() });
    } catch (e) {
      setState({ loading: false, progress: 0, error: graphErrorText(e) });
    }
  };

  const groups = state.groups || [];
  const counts = React.useMemo(() => ({ none: groups.filter((g) => g.ownerCount === 0).length, one: groups.filter((g) => g.ownerCount === 1).length }), [groups]);
  const rows = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    return groups
      .filter((g) => (show === 'none' ? g.ownerCount === 0 : show === 'one' ? g.ownerCount === 1 : true))
      .filter((g) => !q || (g.displayName + ' ' + (g.mail || '')).toLowerCase().indexOf(q) >= 0);
  }, [groups, show, query]);
  // A stable array keeps the selection while typing in the owner box.
  const visible = React.useMemo(() => rows.slice(0, 2000), [rows]);

  const openSite = async (g: IM365Group): Promise<void> => {
    try {
      const res = await ctx.search.run(ctx.homeWebUrl, { query: `contentclass:STS_Site GroupId:${g.id}`, select: ['Path', 'Title'], rowLimit: 1 });
      const url = res.rows[0] && res.rows[0].Path;
      if (!url) {
        ctx.notify('The site for this group was not found in search (it may not be indexed yet or you have no access to it).', 'info');
        return;
      }
      const [web, site] = await Promise.all([ctx.api.getWeb(url), ctx.api.getSite(url)]);
      ctx.setTarget({ webUrl: web.url, siteUrl: site.url, title: web.title });
      ctx.notify(`Now managing ${web.title}.`, 'success');
    } catch (e) {
      ctx.notify(graphErrorText(e), 'error');
    }
  };

  const addOwner = async (): Promise<void> => {
    const picked = selection.getSelection() as IM365Group[];
    if (!picked.length || !owner.trim()) {
      return;
    }
    setBusy(true);
    try {
      const user = await ctx.graph.findUser(owner.trim());
      if (!user) {
        ctx.notify(`${owner} was not found in the directory.`, 'error');
        return;
      }
      const ok = await ctx.confirm({ title: 'Add owner', message: `Add ${user.displayName} as an owner of ${picked.length} group${picked.length === 1 ? '' : 's'}?`, confirmText: 'Add owner' });
      if (!ok) {
        return;
      }
      let done = 0;
      const failed: string[] = [];
      for (const g of picked.slice(0, MAX_BULK)) {
        try {
          await ctx.graph.addToGroup(g.id, 'owners', user.id);
          ctx.log('Add group owner', `${user.displayName} → ${g.displayName}`, true);
          done++;
          g.ownerCount = (g.ownerCount || 0) + 1;
        } catch (e) {
          failed.push(g.displayName);
          ctx.log('Add group owner', g.displayName, false, graphErrorText(e));
        }
      }
      setState((s) => ({ ...s, groups: (s.groups || []).slice() }));
      selection.setAllSelected(false);
      ctx.notify(failed.length ? `${done} updated; failed: ${failed.slice(0, 5).join(', ')}${failed.length > 5 ? '…' : ''}` : `${user.displayName} is now an owner of ${done} group${done === 1 ? '' : 's'}.`, failed.length ? 'error' : 'success');
    } catch (e) {
      ctx.notify(graphErrorText(e), 'error');
    } finally {
      setBusy(false);
    }
  };

  const columns: IColumn[] = [
    { key: 'n', name: 'Group', minWidth: 200, isResizable: true, onRender: (g: IM365Group) => <span title={g.mail}>{g.displayName}</span> },
    { key: 'o', name: 'Owners', minWidth: 70, onRender: (g: IM365Group) => (g.ownerCount === 0 ? <Pill kind="critical">None</Pill> : g.ownerCount === 1 ? <Pill kind="warning">1</Pill> : String(g.ownerCount)) },
    { key: 't', name: 'Team', minWidth: 60, onRender: (g: IM365Group) => (g.isTeam ? 'Yes' : '') },
    { key: 'v', name: 'Privacy', minWidth: 70, onRender: (g: IM365Group) => g.visibility || '' },
    { key: 'c', name: 'Created', minWidth: 90, onRender: (g: IM365Group) => formatDate(g.createdDateTime ? new Date(g.createdDateTime) : undefined) },
    {
      key: 'a',
      name: '',
      minWidth: 90,
      onRender: (g: IM365Group) => (
        <button type="button" className={styles.link} onClick={() => openSite(g)}>
          Manage site
        </button>
      )
    }
  ];

  return (
    <>
      <Card>
        <p className={styles.muted} style={{ marginTop: 0 }}>
          Reads every Microsoft 365 group with its owners, 100 groups per request, paced like everything else (about 2 minutes for 17,000 groups). It only runs when you click Scan.
        </p>
        <div className={styles.actions}>
          <PrimaryButton iconProps={{ iconName: 'Search' }} onClick={scan} disabled={state.loading}>
            {state.loading ? `Scanning… ${state.progress.toLocaleString()} groups` : state.groups ? 'Scan again' : 'Scan all groups'}
          </PrimaryButton>
          {state.groups && (
            <DefaultButton
              iconProps={{ iconName: 'Download' }}
              onClick={() => downloadCsv('m365-group-owners.csv', ['Group', 'Mail', 'Owners', 'Owner names', 'Team', 'Privacy', 'Created', 'Group id'], groups.map((g) => [g.displayName, g.mail, g.ownerCount, (g.owners || []).map((o) => o.userPrincipalName || o.displayName).join('; '), g.isTeam ? 'Yes' : 'No', g.visibility, g.createdDateTime, g.id]))}
            >
              Export all
            </DefaultButton>
          )}
        </div>
        {state.error && <MessageBar messageBarType={MessageBarType.error} isMultiline>{state.error}</MessageBar>}
        {state.truncated && <MessageBar messageBarType={MessageBarType.warning}>Stopped at 50,000 groups.</MessageBar>}
      </Card>
      {state.groups && (
        <>
          <div className={styles.kpiRow}>
            <Kpi label="Microsoft 365 groups" value={groups.length.toLocaleString()} sub={`${groups.filter((g) => g.isTeam).length.toLocaleString()} with a team`} onClick={() => setShow('all')} />
            <Kpi label="No owner" value={counts.none.toLocaleString()} sub="nobody manages access" onClick={() => setShow('none')} />
            <Kpi label="Single owner" value={counts.one.toLocaleString()} sub="at risk when they leave" onClick={() => setShow('one')} />
          </div>
          <Card
            title={`${rows.length.toLocaleString()} groups`}
            right={
              <Dropdown
                ariaLabel="Show"
                selectedKey={show}
                onChange={(_, o) => o && setShow(o.key as OwnerFilter)}
                options={[
                  { key: 'none', text: 'No owner' },
                  { key: 'one', text: 'Single owner' },
                  { key: 'all', text: 'All groups' }
                ]}
                styles={{ root: { width: 160 } }}
              />
            }
          >
            <div className={styles.actions} style={{ marginBottom: 8 }}>
              <SearchBox placeholder="Filter by name or e-mail" value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { width: 260 } }} />
              <TextField ariaLabel="Owner e-mail" placeholder="Owner e-mail to add" value={owner} onChange={(_, v) => setOwner(v || '')} styles={{ root: { width: 260 } }} />
              <PrimaryButton onClick={addOwner} disabled={busy || !owner.trim() || selectedCount === 0}>
                Add as owner to {selectedCount || ''} selected
              </PrimaryButton>
            </div>
            {rows.length === 0 ? (
              <Empty text="No groups match." />
            ) : (
              <DetailsList items={visible} columns={columns} selection={selection} selectionMode={SelectionMode.multiple} layoutMode={DetailsListLayoutMode.justified} />
            )}
            {rows.length > 2000 && <p className={styles.muted}>Showing the first 2,000; use the filter or Export all.</p>}
          </Card>
        </>
      )}
    </>
  );
};
