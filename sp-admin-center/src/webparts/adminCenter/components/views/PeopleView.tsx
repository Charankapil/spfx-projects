import * as React from 'react';
import {
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  Dropdown,
  IColumn,
  MessageBar,
  MessageBarType,
  Pivot,
  PivotItem,
  PrimaryButton,
  SearchBox,
  SelectionMode,
  TextField,
  Toggle
} from '@fluentui/react';
import { IGroupInfo, IUserInfo } from '../../models';
import { parseAddresses } from '../../services/addresses';
import { isOrgWideLogin, matchUsers } from '../../services/AdminApi';
import { IRoleAssignment } from '../../models';
import { downloadCsv } from '../../services/exportCsv';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { useLoader } from '../shared/hooks';
import { Card, Empty, ErrorBar, Kpi, Loading, Pill, ViewHeader } from '../shared/ui';

const MAX_BULK = 200;

function kindPill(u: IUserInfo): React.ReactNode {
  switch (u.kind) {
    case 'Guest':
      return <Pill kind="warning">Guest</Pill>;
    case 'OrgWide':
      return <Pill kind="critical">Organisation-wide</Pill>;
    case 'Group':
      return <Pill kind="info">Group</Pill>;
    case 'System':
      return <Pill>System</Pill>;
    default:
      return <Pill kind="good">Member</Pill>;
  }
}

export const PeopleView: React.FC = () => {
  const ctx = useAdmin();
  const { api, target } = ctx;
  const people = useLoader(async () => {
    const [u, groups] = await Promise.all([api.getUsers(target.webUrl), api.getGroups(target.webUrl)]);
    return { users: u.users, truncated: u.truncated, groups };
  }, [target.webUrl]);

  if (people.loading && !people.data) {
    return <Loading text="Loading people and groups…" />;
  }
  if (people.error || !people.data) {
    return <ErrorBar error={people.error || 'Nothing loaded.'} onRetry={people.reload} />;
  }
  const { users, groups, truncated } = people.data;
  const guests = users.filter((u) => u.kind === 'Guest');
  const admins = users.filter((u) => u.isSiteAdmin);

  return (
    <div className={styles.view}>
      <ViewHeader title="People & permissions" hint={`Who can do what in ${target.title}. Everything here is security-trimmed to what you can manage.`}>
        <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={people.reload} disabled={people.loading}>
          Refresh
        </DefaultButton>
      </ViewHeader>
      <div className={styles.kpiRow}>
        <Kpi label="People on site" value={users.filter((u) => u.kind === 'Member').length.toLocaleString()} />
        <Kpi label="Guests" value={guests.length.toLocaleString()} sub="external" />
        <Kpi label="Site admins" value={String(admins.length)} />
        <Kpi label="Groups" value={String(groups.length)} />
      </div>
      {truncated && <MessageBar messageBarType={MessageBarType.warning}>Only the first 5,000 users are loaded to protect SharePoint from heavy paging.</MessageBar>}

      <Pivot aria-label="People views">
        <PivotItem headerText="Groups">
          <div style={{ paddingTop: 12 }}>
            <GroupsTab groups={groups} onChanged={people.reload} />
          </div>
        </PivotItem>
        <PivotItem headerText="Users & guests">
          <div style={{ paddingTop: 12 }}>
            <UsersTab users={users} onChanged={people.reload} />
          </div>
        </PivotItem>
        <PivotItem headerText="Direct access">
          <div style={{ paddingTop: 12 }}>
            <DirectAccessTab />
          </div>
        </PivotItem>
        <PivotItem headerText="Access checker">
          <div style={{ paddingTop: 12 }}>
            <AccessChecker users={users} />
          </div>
        </PivotItem>
        <PivotItem headerText="Bulk add">
          <div style={{ paddingTop: 12 }}>
            <BulkAdd groups={groups} onChanged={people.reload} />
          </div>
        </PivotItem>
        <PivotItem headerText="Bulk remove">
          <div style={{ paddingTop: 12 }}>
            <BulkRemove users={users} onChanged={people.reload} />
          </div>
        </PivotItem>
        <PivotItem headerText="Everyone groups" itemCount={users.filter((u) => u.kind === 'OrgWide').length || undefined}>
          <div style={{ paddingTop: 12 }}>
            <OrgWideTab users={users} groups={groups} onChanged={people.reload} />
          </div>
        </PivotItem>
      </Pivot>
    </div>
  );
};

// ---- Groups -------------------------------------------------------------------

const GroupsTab: React.FC<{ groups: IGroupInfo[]; onChanged: () => void }> = ({ groups }) => {
  return (
    <div className={styles.grid} style={{ gridTemplateColumns: 'repeat(auto-fit, minmax(340px, 1fr))' }}>
      {groups.map((g) => (
        <GroupCard key={g.id} group={g} />
      ))}
      {groups.length === 0 && <Empty text="No SharePoint groups on this site." />}
    </div>
  );
};

const GroupCard: React.FC<{ group: IGroupInfo }> = ({ group }) => {
  const ctx = useAdmin();
  const { api, target } = ctx;
  const [open, setOpen] = React.useState(false);
  const members = useLoader(() => api.getGroupMembers(target.webUrl, group.id), [group.id, target.webUrl], open);
  const [email, setEmail] = React.useState('');
  const [busy, setBusy] = React.useState(false);

  const add = async (): Promise<void> => {
    const list = parseAddresses(email);
    if (!list.length) {
      ctx.notify('Enter an e-mail address.', 'error');
      return;
    }
    setBusy(true);
    try {
      const u = await api.ensureUser(target.webUrl, list[0]);
      await api.addUserToGroup(target.webUrl, group.id, u.loginName);
      ctx.log('Add to group', `${u.title || list[0]} → ${group.title}`, true);
      ctx.notify(`${u.title || list[0]} added to ${group.title}.`, 'success');
      setEmail('');
      members.reload();
    } catch (e) {
      ctx.log('Add to group', `${list[0]} → ${group.title}`, false, (e as Error).message);
      ctx.notify((e as Error).message, 'error');
    } finally {
      setBusy(false);
    }
  };

  const remove = async (u: IUserInfo): Promise<void> => {
    const ok = await ctx.confirm({
      title: 'Remove from group',
      message: `Remove ${u.title} from "${group.title}"?${group.role === 'Owners' ? ' Make sure the site keeps at least one owner.' : ''}`,
      confirmText: 'Remove',
      danger: true
    });
    if (!ok) {
      return;
    }
    try {
      await api.removeUserFromGroup(target.webUrl, group.id, u.id);
      ctx.log('Remove from group', `${u.title} ← ${group.title}`, true);
      members.reload();
    } catch (e) {
      ctx.log('Remove from group', `${u.title} ← ${group.title}`, false, (e as Error).message);
      ctx.notify((e as Error).message, 'error');
    }
  };

  return (
    <Card
      title={group.title}
      right={group.role ? <Pill kind="info">{group.role}</Pill> : undefined}
    >
      {group.description && <p className={styles.muted} style={{ marginTop: 0 }}>{group.description}</p>}
      <div className={styles.muted}>Owner: {group.ownerTitle || '-'}</div>
      <div style={{ marginTop: 8 }}>
        <DefaultButton iconProps={{ iconName: open ? 'ChevronUp' : 'People' }} onClick={() => setOpen(!open)} aria-expanded={open}>
          {open ? 'Hide members' : 'Show members'}
        </DefaultButton>
      </div>
      {open && (
        <div style={{ marginTop: 8 }}>
          {members.loading && !members.data && <Loading />}
          {members.error && <ErrorBar error={members.error} onRetry={members.reload} />}
          {members.data && (
            <ul className={styles.memberList}>
              {members.data.length === 0 && <li className={styles.muted}>No members.</li>}
              {members.data.slice(0, 200).map((m) => (
                <li key={m.id} className={styles.memberRow}>
                  <span>
                    {m.title} {m.email && <span className={styles.muted}>· {m.email}</span>} {m.kind !== 'Member' && kindPill(m)}
                  </span>
                  <button type="button" className={styles.link} onClick={() => remove(m)} aria-label={`Remove ${m.title} from ${group.title}`}>
                    Remove
                  </button>
                </li>
              ))}
              {members.data.length > 200 && <li className={styles.muted}>…and {members.data.length - 200} more.</li>}
            </ul>
          )}
          <div className={styles.filters} style={{ marginTop: 10 }}>
            <TextField className={styles.field} placeholder="name@company.com" value={email} onChange={(_, v) => setEmail(v || '')} ariaLabel={`Add a person to ${group.title}`} onKeyDown={(e) => e.key === 'Enter' && !busy && add()} />
            <PrimaryButton onClick={add} disabled={busy || !email.trim()}>
              Add
            </PrimaryButton>
          </div>
        </div>
      )}
    </Card>
  );
};

// ---- Users & guests ---------------------------------------------------------

const UsersTab: React.FC<{ users: IUserInfo[]; onChanged: () => void }> = ({ users, onChanged }) => {
  const ctx = useAdmin();
  const { api, target } = ctx;
  const [text, setText] = React.useState('');
  const [kind, setKind] = React.useState('Member');

  const rows = users.filter((u) => {
    if (kind === 'Admins' ? !u.isSiteAdmin : kind !== 'all' && u.kind !== kind) { // admins filter ignores type
      return false;
    }
    const q = text.trim().toLowerCase();
    return !q || u.title.toLowerCase().indexOf(q) >= 0 || u.email.toLowerCase().indexOf(q) >= 0 || u.loginName.toLowerCase().indexOf(q) >= 0;
  });

  const removeUser = async (u: IUserInfo): Promise<void> => {
    const ok = await ctx.confirm({
      title: 'Remove from site',
      message: `Remove ${u.title} from this site collection? They lose all access granted on it. Their existing content is not deleted.`,
      confirmText: 'Remove user',
      danger: true
    });
    if (!ok) {
      return;
    }
    try {
      await api.removeUserFromSite(target.webUrl, u.id);
      ctx.log('Remove user from site', u.title, true);
      ctx.notify(`${u.title} removed.`, 'success');
      onChanged();
    } catch (e) {
      ctx.log('Remove user from site', u.title, false, (e as Error).message);
      ctx.notify((e as Error).message, 'error');
    }
  };

  const toggleAdmin = async (u: IUserInfo): Promise<void> => {
    const make = !u.isSiteAdmin;
    const ok = await ctx.confirm({
      title: make ? 'Make site collection admin' : 'Remove site collection admin',
      message: make ? `${u.title} will get full control of the whole site collection.` : `${u.title} will no longer be a site collection administrator.`,
      confirmText: make ? 'Make admin' : 'Remove admin',
      danger: !make
    });
    if (!ok) {
      return;
    }
    try {
      await api.setSiteAdmin(target.webUrl, u.id, make);
      ctx.log(make ? 'Grant site admin' : 'Revoke site admin', u.title, true);
      ctx.notify(make ? `${u.title} is now an admin.` : `${u.title} is no longer an admin.`, 'success');
      onChanged();
    } catch (e) {
      ctx.log(make ? 'Grant site admin' : 'Revoke site admin', u.title, false, (e as Error).message);
      ctx.notify((e as Error).message, 'error');
    }
  };

  const columns: IColumn[] = [
    { key: 'name', name: 'Name', minWidth: 150, isResizable: true, onRender: (u: IUserInfo) => <strong>{u.title}</strong> },
    { key: 'email', name: 'E-mail / login', minWidth: 200, isResizable: true, onRender: (u: IUserInfo) => u.email || u.loginName },
    { key: 'kind', name: 'Type', minWidth: 120, onRender: (u: IUserInfo) => kindPill(u) },
    { key: 'admin', name: 'Site admin', minWidth: 80, onRender: (u: IUserInfo) => (u.isSiteAdmin ? <Pill kind="warning">Admin</Pill> : '') },
    {
      key: 'act',
      name: 'Actions',
      minWidth: 220,
      onRender: (u: IUserInfo) =>
        u.kind === 'System' ? (
          ''
        ) : (
          <div className={styles.actions}>
            {u.kind === 'Member' && (
              <button type="button" className={styles.link} onClick={() => toggleAdmin(u)}>
                {u.isSiteAdmin ? 'Remove admin' : 'Make admin'}
              </button>
            )}
            <button type="button" className={styles.link} onClick={() => removeUser(u)}>
              Remove from site
            </button>
          </div>
        )
    }
  ];

  return (
    <>
      <div className={styles.filters}>
        <SearchBox className={styles.field} placeholder="Search name or e-mail" value={text} onChange={(_, v) => setText(v || '')} />
        <Dropdown
          className={styles.field}
          label="Show"
          selectedKey={kind}
          onChange={(_, o) => setKind(String(o ? o.key : 'Member'))}
          options={[
            { key: 'Member', text: 'People' },
            { key: 'Guest', text: 'Guests (external)' },
            { key: 'Admins', text: 'Site collection admins' },
            { key: 'OrgWide', text: 'Organisation-wide groups' },
            { key: 'Group', text: 'Security / SharePoint groups' },
            { key: 'all', text: 'Everyone' }
          ]}
        />
        <DefaultButton
          iconProps={{ iconName: 'Download' }}
          onClick={() => downloadCsv('users.csv', ['Name', 'E-mail', 'Login', 'Type', 'Site admin'], rows.map((u) => [u.title, u.email, u.loginName, u.kind, u.isSiteAdmin]))}
        >
          Export CSV
        </DefaultButton>
        <span className={styles.muted}>{rows.length.toLocaleString()} shown</span>
      </div>
      <div className={styles.tableWrap} style={{ marginTop: 10 }}>
        {rows.length === 0 ? <Empty text="Nobody matches." /> : <DetailsList items={rows.slice(0, 500)} columns={columns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} getKey={(u: IUserInfo) => String(u.id)} />}
        {rows.length > 500 && <div className={styles.empty}>Showing the first 500. Search or export the CSV for the rest.</div>}
      </div>
    </>
  );
};

// ---- Direct access ----------------------------------------------------------

const DirectAccessTab: React.FC = () => {
  const ctx = useAdmin();
  const { api, target } = ctx;
  const ra = useLoader(() => api.getRoleAssignments(target.webUrl), [target.webUrl]);
  if (ra.loading && !ra.data) {
    return <Loading />;
  }
  if (ra.error || !ra.data) {
    return <ErrorBar error={ra.error || 'Nothing loaded.'} onRetry={ra.reload} />;
  }
  return (
    <Card title="Who is granted permission directly on this site">
      <p className={styles.muted} style={{ marginTop: 0 }}>
        Groups and people with a permission level on the site itself (not inherited by individual lists).
      </p>
      <div className={styles.tableWrap}>
        <DetailsList
          items={ra.data}
          selectionMode={SelectionMode.none}
          layoutMode={DetailsListLayoutMode.justified}
          getKey={(r: { principalId: number }) => String(r.principalId)}
          columns={[
            { key: 'p', name: 'Person or group', minWidth: 200, isResizable: true, onRender: (r: { principalTitle: string }) => <strong>{r.principalTitle}</strong> },
            { key: 't', name: 'Kind', minWidth: 110, onRender: (r: { principalType: number }) => (r.principalType === 8 ? 'SharePoint group' : r.principalType === 4 ? 'Security group' : r.principalType === 1 ? 'Person' : 'Other') },
            { key: 'r', name: 'Permission levels', minWidth: 220, onRender: (r: { roles: string[] }) => r.roles.join(', ') }
          ]}
        />
      </div>
    </Card>
  );
};

// ---- Access checker ---------------------------------------------------------

const AccessChecker: React.FC<{ users: IUserInfo[] }> = ({ users }) => {
  const ctx = useAdmin();
  const { api, target } = ctx;
  const [q, setQ] = React.useState('');
  const [picked, setPicked] = React.useState<IUserInfo | undefined>();
  const [result, setResult] = React.useState<{ groups: string[]; level: string } | undefined>();
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>();

  const matches = q.trim().length < 2 ? [] : users.filter((u) => u.kind !== 'System' && (u.title.toLowerCase().indexOf(q.toLowerCase()) >= 0 || u.email.toLowerCase().indexOf(q.toLowerCase()) >= 0)).slice(0, 6);

  const check = async (u: IUserInfo): Promise<void> => {
    setPicked(u);
    setBusy(true);
    setError(undefined);
    setResult(undefined);
    try {
      const [groups, level] = await Promise.all([api.getUserGroups(target.webUrl, u.id), api.getEffectiveLevel(target.webUrl, u.loginName)]);
      setResult({ groups: groups.map((g) => g.title), level });
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  return (
    <Card title="What can this person do here?">
      <SearchBox className={styles.field} placeholder="Type a name or e-mail (min. 2 letters)" value={q} onChange={(_, v) => setQ(v || '')} />
      {matches.length > 0 && (
        <ul className={styles.memberList}>
          {matches.map((u) => (
            <li key={u.id} className={styles.memberRow}>
              <span>
                {u.title} <span className={styles.muted}>· {u.email || u.loginName}</span>
              </span>
              <button type="button" className={styles.link} onClick={() => check(u)}>
                Check access
              </button>
            </li>
          ))}
        </ul>
      )}
      {busy && <Loading />}
      {error && <ErrorBar error={error} />}
      {picked && result && (
        <div style={{ marginTop: 14 }}>
          <strong>{picked.title}</strong> has <Pill kind={result.level === 'No access' ? 'critical' : 'good'}>{result.level}</Pill> on {target.title}.
          <div className={styles.muted} style={{ marginTop: 6 }}>
            Member of: {result.groups.length ? result.groups.join(', ') : 'no SharePoint groups (access comes from a direct grant or an organisation group)'}
          </div>
        </div>
      )}
    </Card>
  );
};

// ---- Bulk add ---------------------------------------------------------------

const BulkAdd: React.FC<{ groups: IGroupInfo[]; onChanged: () => void }> = ({ groups, onChanged }) => {
  const ctx = useAdmin();
  const { api, target } = ctx;
  const [groupId, setGroupId] = React.useState<number | undefined>(() => {
    const m = groups.filter((g) => g.role === 'Members')[0] || groups[0];
    return m ? m.id : undefined;
  });
  const [text, setText] = React.useState('');
  const [results, setResults] = React.useState<Array<{ who: string; ok: boolean; message?: string }>>([]);
  const [busy, setBusy] = React.useState(false);
  const list = parseAddresses(text);

  const run = async (): Promise<void> => {
    const group = groups.filter((g) => g.id === groupId)[0];
    if (!group || !list.length) {
      return;
    }
    if (list.length > MAX_BULK) {
      ctx.notify(`Please add at most ${MAX_BULK} people per run to stay within SharePoint's limits.`, 'error');
      return;
    }
    const ok = await ctx.confirm({ title: 'Add people', message: `Add ${list.length} people to "${group.title}" on ${target.title}?`, confirmText: 'Add them' });
    if (!ok) {
      return;
    }
    setBusy(true);
    setResults([]);
    const out: Array<{ who: string; ok: boolean; message?: string }> = [];
    for (const who of list) {
      try {
        const u = await api.ensureUser(target.webUrl, who);
        await api.addUserToGroup(target.webUrl, group.id, u.loginName);
        out.push({ who, ok: true });
        ctx.log('Bulk add to group', `${who} → ${group.title}`, true);
      } catch (e) {
        out.push({ who, ok: false, message: (e as Error).message });
        ctx.log('Bulk add to group', `${who} → ${group.title}`, false, (e as Error).message);
      }
      setResults(out.slice());
    }
    setBusy(false);
    ctx.notify(`${out.filter((r) => r.ok).length} of ${out.length} added.`, 'info');
    onChanged();
  };

  return (
    <Card title="Add many people to a group at once">
      <div className={styles.filters}>
        <Dropdown
          className={styles.field}
          label="Group"
          selectedKey={groupId}
          onChange={(_, o) => setGroupId(o ? Number(o.key) : undefined)}
          options={groups.map((g) => ({ key: g.id, text: g.title }))}
        />
      </div>
      <TextField
        label="E-mail addresses (one per line, or separated by commas / semicolons)"
        multiline
        rows={6}
        value={text}
        onChange={(_, v) => setText(v || '')}
        placeholder={'alex@contoso.com\nsam@contoso.com'}
      />
      <div className={styles.actions} style={{ marginTop: 10 }}>
        <PrimaryButton onClick={run} disabled={busy || !list.length || groupId === undefined}>
          Add {list.length || ''} {list.length === 1 ? 'person' : 'people'}
        </PrimaryButton>
        <span className={styles.muted}>Up to {MAX_BULK} per run; one at a time so SharePoint is never flooded.</span>
      </div>
      {results.length > 0 && (
        <ul className={styles.memberList} style={{ marginTop: 12 }}>
          {results.map((r) => (
            <li key={r.who} className={styles.memberRow}>
              <span>{r.who}</span>
              {r.ok ? <Pill kind="good">Added</Pill> : <Pill kind="critical">{r.message ? r.message.substring(0, 80) : 'Failed'}</Pill>}
            </li>
          ))}
        </ul>
      )}
    </Card>
  );
};

// ---- Bulk remove ------------------------------------------------------------

type RemovePlan = { user: IUserInfo; skip?: string };

const BulkRemove: React.FC<{ users: IUserInfo[]; onChanged: () => void }> = ({ users, onChanged }) => {
  const ctx = useAdmin();
  const { api, target } = ctx;
  const [text, setText] = React.useState('');
  const [includeAdmins, setIncludeAdmins] = React.useState(false);
  const [plan, setPlan] = React.useState<{ rows: RemovePlan[]; unmatched: string[] } | undefined>();
  const [results, setResults] = React.useState<Array<{ who: string; ok: boolean; message?: string }>>([]);
  const [busy, setBusy] = React.useState(false);

  const isMe = (u: IUserInfo): boolean => {
    const me = ctx.currentUser;
    const login = u.loginName.toLowerCase();
    return !!me && (u.email.toLowerCase() === me || login.substring(login.lastIndexOf('|') + 1) === me);
  };

  const find = (): void => {
    const entries = parseAddresses(text);
    const m = matchUsers(entries, users);
    setResults([]);
    setPlan({
      unmatched: m.unmatched,
      rows: m.matched.map((u) => ({
        user: u,
        skip: isMe(u) ? 'That is you' : u.kind === 'System' ? 'System account' : u.isSiteAdmin && !includeAdmins ? 'Site collection admin' : undefined
      }))
    });
  };

  const removable = plan ? plan.rows.filter((r) => !r.skip) : [];

  const run = async (): Promise<void> => {
    if (!removable.length) {
      return;
    }
    if (removable.length > MAX_BULK) {
      ctx.notify(`Please remove at most ${MAX_BULK} people per run to stay within SharePoint's limits.`, 'error');
      return;
    }
    const ok = await ctx.confirm({
      title: 'Remove people from the site collection',
      message: `Remove ${removable.length} ${removable.length === 1 ? 'person' : 'people'} from ${target.siteUrl}? They lose every permission in this site collection (site, lists, libraries and items) and are removed from all its SharePoint groups. Their content is not deleted. For a Microsoft 365 group-connected site, also remove them from the group, or they keep access through it.`,
      confirmText: `Remove ${removable.length}`,
      danger: true
    });
    if (!ok) {
      return;
    }
    setBusy(true);
    const out: Array<{ who: string; ok: boolean; message?: string }> = [];
    for (const r of removable) {
      const who = r.user.email || r.user.title;
      try {
        await api.removeUserFromSite(target.siteUrl, r.user.id);
        out.push({ who, ok: true });
        ctx.log('Bulk remove from site collection', `${who} ← ${target.siteUrl}`, true);
      } catch (e) {
        out.push({ who, ok: false, message: (e as Error).message });
        ctx.log('Bulk remove from site collection', `${who} ← ${target.siteUrl}`, false, (e as Error).message);
      }
      setResults(out.slice());
    }
    setBusy(false);
    setPlan(undefined);
    ctx.notify(`${out.filter((r) => r.ok).length} of ${out.length} removed.`, out.every((r) => r.ok) ? 'success' : 'error');
    onChanged();
  };

  return (
    <Card title="Remove many people from this site collection">
      <TextField
        label="E-mail addresses or logins (one per line, or separated by commas / semicolons)"
        multiline
        rows={6}
        value={text}
        onChange={(_, v) => {
          setText(v || '');
          setPlan(undefined);
        }}
        placeholder={'alex@contoso.com\npartner_fabrikam.com#ext#@contoso.onmicrosoft.com'}
      />
      <div className={styles.actions} style={{ marginTop: 10 }}>
        <DefaultButton onClick={find} disabled={busy || !parseAddresses(text).length}>
          Find these people
        </DefaultButton>
        <Toggle
          label="Also remove site collection admins"
          inlineLabel
          checked={includeAdmins}
          onChange={(_, c) => {
            setIncludeAdmins(!!c);
            setPlan(undefined);
          }}
        />
      </div>

      {plan && (
        <div style={{ marginTop: 12 }}>
          {plan.unmatched.length > 0 && (
            <MessageBar messageBarType={MessageBarType.warning} isMultiline>
              Not found on this site ({plan.unmatched.length}): {plan.unmatched.slice(0, 15).join(', ')}
              {plan.unmatched.length > 15 ? '…' : ''}
            </MessageBar>
          )}
          {plan.rows.length === 0 ? (
            <Empty text="None of these people are on this site." />
          ) : (
            <ul className={styles.memberList}>
              {plan.rows.map((r) => (
                <li key={r.user.id} className={styles.memberRow}>
                  <span>
                    <strong>{r.user.title}</strong> <span className={styles.muted}>· {r.user.email || r.user.loginName}</span> {kindPill(r.user)}
                  </span>
                  {r.skip ? <Pill>Skipped: {r.skip}</Pill> : <Pill kind="warning">Will be removed</Pill>}
                </li>
              ))}
            </ul>
          )}
          <div className={styles.actions} style={{ marginTop: 10 }}>
            <PrimaryButton onClick={run} disabled={busy || removable.length === 0} styles={{ root: { background: '#c4314b', borderColor: '#c4314b' }, rootHovered: { background: '#a4262c', borderColor: '#a4262c' } }}>
              Remove {removable.length} from the site collection
            </PrimaryButton>
            <span className={styles.muted}>Up to {MAX_BULK} per run, one at a time.</span>
          </div>
        </div>
      )}

      {results.length > 0 && (
        <ul className={styles.memberList} style={{ marginTop: 12 }}>
          {results.map((r) => (
            <li key={r.who} className={styles.memberRow}>
              <span>{r.who}</span>
              {r.ok ? <Pill kind="good">Removed</Pill> : <Pill kind="critical">{r.message ? r.message.substring(0, 80) : 'Failed'}</Pill>}
            </li>
          ))}
        </ul>
      )}
    </Card>
  );
};

// ---- Everyone / Everyone except external users --------------------------------

const MAX_GROUP_SCAN = 60;

interface IOrgWideFinding {
  user: IUserInfo;
  groups: IGroupInfo[];
  direct?: IRoleAssignment;
}

const OrgWideTab: React.FC<{ users: IUserInfo[]; groups: IGroupInfo[]; onChanged: () => void }> = ({ users, groups, onChanged }) => {
  const ctx = useAdmin();
  const { api, target } = ctx;
  const orgUsers = users.filter((u) => u.kind === 'OrgWide');
  // One request for the site's direct permissions plus one per group (at most 60), only when this tab is opened.
  const scan = useLoader(async (): Promise<IOrgWideFinding[]> => {
    if (!orgUsers.length) {
      return [];
    }
    const assignments = await api.getRoleAssignments(target.webUrl).catch((): IRoleAssignment[] => []);
    const inGroups: { [userId: number]: IGroupInfo[] } = {};
    for (const g of groups.slice(0, MAX_GROUP_SCAN)) {
      const members = await api.getGroupMembers(target.webUrl, g.id).catch((): IUserInfo[] => []);
      members.filter((m) => isOrgWideLogin(m.loginName)).forEach((m) => (inGroups[m.id] = (inGroups[m.id] || []).concat(g)));
    }
    return orgUsers.map((u) => ({
      user: u,
      groups: inGroups[u.id] || [],
      direct: assignments.filter((a) => a.principalId === u.id || a.principalLogin.toLowerCase() === u.loginName.toLowerCase())[0]
    }));
  }, [target.webUrl, orgUsers.length, groups.length]);

  const act = async (title: string, message: string, label: string, fn: () => Promise<void>, logText: string): Promise<void> => {
    const ok = await ctx.confirm({ title, message, confirmText: label, danger: true });
    if (!ok) {
      return;
    }
    try {
      await fn();
      ctx.log(title, logText, true);
      ctx.notify(`${title}: done.`, 'success');
      scan.reload();
      onChanged();
    } catch (e) {
      ctx.log(title, logText, false, (e as Error).message);
      ctx.notify((e as Error).message, 'error');
    }
  };

  if (orgUsers.length === 0) {
    return (
      <Card title="Everyone and Everyone except external users">
        <FindingRowLite good text="Neither group has been added to this site collection. Content here is not opened up to the whole organisation through them." />
      </Card>
    );
  }
  if (scan.loading && !scan.data) {
    return <Loading text={`Checking ${Math.min(groups.length, MAX_GROUP_SCAN)} groups and the site's permissions…`} />;
  }
  if (scan.error || !scan.data) {
    return <ErrorBar error={scan.error || 'Nothing loaded.'} onRetry={scan.reload} />;
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
      <MessageBar messageBarType={MessageBarType.warning} isMultiline>
        These claims stand for the whole organisation (<em>Everyone except external users</em>) or everyone including guests (<em>Everyone</em>). Wherever they have a permission, every one of those people can open that content, and it shows up in their search results and Copilot answers.
      </MessageBar>
      {scan.data.map((f) => (
        <Card key={f.user.id} title={f.user.title} right={<Pill kind="critical">Organisation-wide</Pill>}>
          <div className={styles.muted} style={{ fontSize: 12, marginBottom: 8 }}>
            {f.user.loginName}
          </div>

          <strong style={{ fontSize: 13 }}>Member of SharePoint groups</strong>
          {f.groups.length === 0 ? (
            <p className={styles.muted} style={{ marginTop: 4 }}>
              None{groups.length > MAX_GROUP_SCAN ? ` (first ${MAX_GROUP_SCAN} groups checked)` : ''}.
            </p>
          ) : (
            <ul className={styles.memberList}>
              {f.groups.map((g) => (
                <li key={g.id} className={styles.memberRow}>
                  <span>
                    {g.title} {g.role && <Pill kind="info">{g.role}</Pill>}
                  </span>
                  <button
                    type="button"
                    className={styles.link}
                    onClick={() => act('Remove from group', `Remove "${f.user.title}" from "${g.title}"? People who only had access through it lose that access.`, 'Remove', () => api.removeUserFromGroup(target.webUrl, g.id, f.user.id), `${f.user.title} ← ${g.title}`)}
                  >
                    Remove from group
                  </button>
                </li>
              ))}
            </ul>
          )}

          <strong style={{ fontSize: 13, display: 'block', marginTop: 10 }}>Direct permission on {target.title}</strong>
          {f.direct ? (
            <div className={styles.memberRow}>
              <span>{f.direct.roles.join(', ') || 'Limited access'}</span>
              <button
                type="button"
                className={styles.link}
                onClick={() => act('Remove direct permission', `Remove the "${f.direct ? f.direct.roles.join(', ') : ''}" permission that "${f.user.title}" has directly on ${target.title}?`, 'Remove permission', () => api.removeRoleAssignment(target.webUrl, f.user.id), `${f.user.title} on ${target.webUrl}`)}
              >
                Remove permission
              </button>
            </div>
          ) : (
            <p className={styles.muted} style={{ marginTop: 4 }}>
              None.
            </p>
          )}

          <div className={styles.actions} style={{ marginTop: 12 }}>
            <PrimaryButton
              styles={{ root: { background: '#c4314b', borderColor: '#c4314b' }, rootHovered: { background: '#a4262c', borderColor: '#a4262c' } }}
              onClick={() =>
                act(
                  'Remove from site collection',
                  `Remove "${f.user.title}" from the whole site collection ${target.siteUrl}? This also clears permissions it was given on individual lists, libraries, folders and items, which this page cannot list.`,
                  'Remove everywhere',
                  () => api.removeUserFromSite(target.siteUrl, f.user.id),
                  `${f.user.title} ← ${target.siteUrl}`
                )
              }
            >
              Remove from the whole site collection
            </PrimaryButton>
            {f.groups.length === 0 && !f.direct && <span className={styles.muted}>Not in a group or on the site itself; it may still be granted on lists, libraries or items.</span>}
          </div>
        </Card>
      ))}
    </div>
  );
};

const FindingRowLite: React.FC<{ good?: boolean; text: string }> = ({ text }) => (
  <div className={styles.memberRow}>
    <span>
      <Pill kind="good">All clear</Pill> {text}
    </span>
  </div>
);
