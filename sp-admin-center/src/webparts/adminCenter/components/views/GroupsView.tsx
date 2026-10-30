import * as React from 'react';
import { MessageBar, MessageBarType, PrimaryButton, SearchBox, TextField } from '@fluentui/react';
import { parseAddresses } from '../../services/addresses';
import { IDirectoryObject } from '../../services/GraphApi';
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
      <ViewHeader title="Microsoft 365 group" hint="Owners and members of the Microsoft 365 group behind the selected site. Tenant-wide group tools are under Microsoft 365 in the menu." />
      <GraphGate scopes={['Group.ReadWrite.All', 'User.Read.All']} feature="Microsoft 365 group management">
        <SiteGroupTab />
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
