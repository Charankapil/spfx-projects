import * as React from 'react';
import { Checkbox, DefaultButton, DetailsList, DetailsListLayoutMode, IColumn, MessageBar, MessageBarType, PrimaryButton, Selection, SelectionMode, TextField } from '@fluentui/react';
import { downloadCsv } from '../../services/exportCsv';
import { IDirectoryObject } from '../../services/GraphApi';
import { IOffboardRow, planOffboarding } from '../../services/m365Review';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { GraphGate, graphErrorText } from '../shared/graph';
import { manageGroupSite } from '../shared/groupSite';
import { Card, Empty, Pill, StatTile, ViewHeader } from '../shared/ui';

const MAX_GROUPS = 200;

/**
 * Hands over a leaving person's Microsoft 365 groups and teams: every group they own or
 * belong to, the ones where they are the only owner first. Their ownership moves to a new
 * owner and they are removed. (User.Read.All + Group.ReadWrite.All.)
 */
export const OffboardingView: React.FC = () => (
  <div className={styles.view}>
    <ViewHeader title="Offboarding" hint="Find every Microsoft 365 group and team a leaving person owns or belongs to, hand their ownership to someone else and remove them, in one go." />
    <GraphGate scopes={['User.Read.All', 'Group.ReadWrite.All']} feature="Offboarding">
      <OffboardingBody />
    </GraphGate>
  </div>
);

interface IPlan {
  leaver: IDirectoryObject;
  rows: IOffboardRow[];
}

const OffboardingBody: React.FC = () => {
  const ctx = useAdmin();
  const [leaverText, setLeaverText] = React.useState('');
  const [newOwnerText, setNewOwnerText] = React.useState('');
  const [removeMembership, setRemoveMembership] = React.useState(true);
  const [plan, setPlan] = React.useState<IPlan | undefined>();
  const [busy, setBusy] = React.useState<string | undefined>();
  const [error, setError] = React.useState<string | undefined>();
  const [results, setResults] = React.useState<Array<{ group: string; outcome: string; ok: boolean }> | undefined>();
  const [selectedCount, setSelectedCount] = React.useState(0);
  const selection = React.useMemo(() => new Selection({ getKey: (r) => (r as IOffboardRow).group.id, onSelectionChanged: () => setSelectedCount(selection.getSelectedCount()) }), []);

  // Everything is selected by default; untick groups to leave them alone.
  React.useEffect(() => {
    if (plan) {
      const t = setTimeout(() => selection.setAllSelected(true), 0);
      return () => clearTimeout(t);
    }
    return undefined;
  }, [plan, selection]);

  const find = async (): Promise<void> => {
    setError(undefined);
    setPlan(undefined);
    setResults(undefined);
    setBusy('Finding the person…');
    try {
      const leaver = await ctx.graph.findUser(leaverText.trim());
      if (!leaver) {
        setError(`${leaverText.trim()} was not found in the directory.`);
        return;
      }
      setBusy('Reading their groups…');
      const [owned, memberOf] = await Promise.all([ctx.graph.userGroups(leaver.id, 'ownedObjects'), ctx.graph.userGroups(leaver.id, 'memberOf')]);
      setBusy('Checking the other owners…');
      const owners = await ctx.graph.ownersMany(owned.map((g) => g.id));
      const rows = planOffboarding(owned, memberOf, owners, leaver.id);
      setPlan({ leaver, rows });
    } catch (e) {
      setError(graphErrorText(e));
    } finally {
      setBusy(undefined);
    }
  };

  const run = async (): Promise<void> => {
    if (!plan) {
      return;
    }
    const picked = (selection.getSelection() as IOffboardRow[]).slice(0, MAX_GROUPS);
    let newOwner: IDirectoryObject | undefined;
    if (newOwnerText.trim()) {
      newOwner = await ctx.graph.findUser(newOwnerText.trim()).catch(() => undefined);
      if (!newOwner) {
        ctx.notify(`${newOwnerText.trim()} was not found in the directory.`, 'error');
        return;
      }
      if (newOwner.id === plan.leaver.id) {
        ctx.notify('The new owner must be someone else.', 'error');
        return;
      }
    }
    const ownedCount = picked.filter((r) => r.isOwner).length;
    const soleCount = picked.filter((r) => r.isOwner && r.otherOwners === 0).length;
    const ok = await ctx.confirm({
      title: `Offboard ${plan.leaver.displayName}`,
      message:
        `${picked.length} group${picked.length === 1 ? '' : 's'} selected. ` +
        (ownedCount ? (newOwner ? `${newOwner.displayName} becomes an owner of the ${ownedCount} group${ownedCount === 1 ? '' : 's'} ${plan.leaver.displayName} owns, then ${plan.leaver.displayName} is removed as owner. ` : `${plan.leaver.displayName} is removed as owner where other owners remain${soleCount ? ` (${soleCount} group${soleCount === 1 ? '' : 's'} where they are the only owner are skipped: name a new owner to include them)` : ''}. `) : '') +
        (removeMembership ? `${plan.leaver.displayName} is removed as a member of every selected group.` : 'Memberships are kept.'),
      confirmText: 'Offboard',
      danger: true
    });
    if (!ok) {
      return;
    }
    const out: Array<{ group: string; outcome: string; ok: boolean }> = [];
    const leaverId = plan.leaver.id;
    for (let i = 0; i < picked.length; i++) {
      const r = picked[i];
      setBusy(`Working… ${i + 1} of ${picked.length}`);
      const steps: string[] = [];
      let ok2 = true;
      try {
        if (r.isOwner) {
          if (newOwner) {
            await ctx.graph.addToGroup(r.group.id, 'owners', newOwner.id).catch((e) => {
              // Already an owner is fine.
              if (!/already exist/i.test(graphErrorText(e))) {
                throw e;
              }
            });
            steps.push(`${newOwner.displayName} added as owner`);
          }
          if (newOwner || (r.otherOwners || 0) > 0) {
            await ctx.graph.removeFromGroup(r.group.id, 'owners', leaverId);
            steps.push('removed as owner');
          } else {
            steps.push('kept as owner (only owner; no new owner given)');
            ok2 = false;
          }
        }
        if (removeMembership && r.isMember) {
          await ctx.graph.removeFromGroup(r.group.id, 'members', leaverId);
          steps.push('removed as member');
        }
        ctx.log('Offboard', `${plan.leaver.displayName} → ${r.group.displayName}: ${steps.join(', ')}`, ok2);
      } catch (e) {
        ok2 = false;
        steps.push(`failed: ${graphErrorText(e).substring(0, 120)}`);
        ctx.log('Offboard', `${plan.leaver.displayName} → ${r.group.displayName}`, false, graphErrorText(e));
      }
      out.push({ group: r.group.displayName, outcome: steps.join('; ') || 'nothing to do', ok: ok2 });
    }
    setBusy(undefined);
    setResults(out);
    const failed = out.filter((o) => !o.ok).length;
    ctx.notify(failed ? `${out.length - failed} done, ${failed} need attention (see the results).` : `${plan.leaver.displayName} offboarded from ${out.length} group${out.length === 1 ? '' : 's'}.`, failed ? 'error' : 'success');
  };

  const rows = plan ? plan.rows : [];
  const sole = rows.filter((r) => r.isOwner && r.otherOwners === 0).length;
  const columns: IColumn[] = [
    { key: 'g', name: 'Group', minWidth: 220, isResizable: true, onRender: (r: IOffboardRow) => <span title={r.group.mail}>{r.group.displayName}</span> },
    {
      key: 'r',
      name: 'Role',
      minWidth: 140,
      onRender: (r: IOffboardRow) => (
        <span style={{ display: 'flex', gap: 4 }}>
          {r.isOwner && <Pill kind={r.otherOwners === 0 ? 'critical' : 'warning'}>{r.otherOwners === 0 ? 'Only owner' : 'Owner'}</Pill>}
          {r.isMember && <Pill>Member</Pill>}
        </span>
      )
    },
    { key: 'o', name: 'Other owners', minWidth: 90, onRender: (r: IOffboardRow) => (r.isOwner ? (r.otherOwners === undefined ? '?' : String(r.otherOwners)) : '') },
    {
      key: 'a',
      name: '',
      minWidth: 90,
      onRender: (r: IOffboardRow) => (
        <button type="button" className={styles.link} onClick={() => manageGroupSite(ctx, r.group.id)}>
          Manage site
        </button>
      )
    }
  ];

  return (
    <>
      <Card>
        <div className={styles.actions} style={{ alignItems: 'flex-end' }}>
          <TextField label="Person leaving" placeholder="name@contoso.com" value={leaverText} onChange={(_, v) => setLeaverText(v || '')} onKeyDown={(e) => e.key === 'Enter' && leaverText.trim() && !busy && find()} styles={{ root: { width: 300 } }} />
          <PrimaryButton iconProps={{ iconName: 'Search' }} onClick={find} disabled={!leaverText.trim() || !!busy}>
            {busy && !plan ? busy : 'Find their groups'}
          </PrimaryButton>
        </div>
        {error && <MessageBar messageBarType={MessageBarType.error} isMultiline>{error}</MessageBar>}
        <p className={styles.muted} style={{ marginBottom: 0, fontSize: 12 }}>
          Covers Microsoft 365 groups and teams (and so their sites). SharePoint-only sites they have access to are in People &amp; access; their OneDrive is handled by Microsoft 365 when the account is deleted.
        </p>
      </Card>

      {plan && (
        <>
          <div className={styles.statGrid}>
            <StatTile icon="Contact" label={plan.leaver.displayName} value={rows.length.toLocaleString()} sub="groups and teams" tone="brand" />
            <StatTile icon="Crown" label="Owner of" value={rows.filter((r) => r.isOwner).length.toLocaleString()} />
            <StatTile icon="Warning" label="Only owner of" value={sole.toLocaleString()} tone={sole ? 'critical' : 'good'} sub={sole ? 'name a new owner' : undefined} />
            <StatTile icon="People" label="Member of" value={rows.filter((r) => r.isMember).length.toLocaleString()} />
          </div>
          <Card title="Hand over and remove">
            <div className={styles.actions} style={{ alignItems: 'flex-end' }}>
              <TextField label="New owner (recommended)" placeholder="manager@contoso.com" value={newOwnerText} onChange={(_, v) => setNewOwnerText(v || '')} styles={{ root: { width: 300 } }} />
              <Checkbox label="Also remove them as a member" checked={removeMembership} onChange={(_, v) => setRemoveMembership(!!v)} />
              <PrimaryButton iconProps={{ iconName: 'UserRemove' }} onClick={run} disabled={!!busy || selectedCount === 0}>
                {busy && plan ? busy : `Offboard from ${selectedCount} selected`}
              </PrimaryButton>
              <DefaultButtonExport plan={plan} />
            </div>
            {rows.length === 0 ? (
              <Empty text="This person does not own or belong to any Microsoft 365 group." />
            ) : (
              <DetailsList items={rows} columns={columns} selection={selection} selectionMode={SelectionMode.multiple} layoutMode={DetailsListLayoutMode.justified} />
            )}
          </Card>
        </>
      )}

      {results && (
        <Card title="Results">
          <ul className={styles.memberList}>
            {results.map((r) => (
              <li key={r.group} className={styles.memberRow}>
                <span>
                  <div>{r.group}</div>
                  <div className={styles.muted} style={{ fontSize: 12 }}>
                    {r.outcome}
                  </div>
                </span>
                {r.ok ? <Pill kind="good">Done</Pill> : <Pill kind="critical">Check</Pill>}
              </li>
            ))}
          </ul>
        </Card>
      )}
    </>
  );
};

const DefaultButtonExport: React.FC<{ plan: IPlan }> = ({ plan }) => (
  <DefaultButton
    iconProps={{ iconName: 'Download' }}
    onClick={() =>
      downloadCsv(
        `offboarding-${(plan.leaver.mail || plan.leaver.userPrincipalName || plan.leaver.displayName).replace(/[^A-Za-z0-9_.-]+/g, '-')}.csv`,
        ['Person', 'Group', 'Group e-mail', 'Owner', 'Only owner', 'Member', 'Other owners', 'Group id'],
        plan.rows.map((r) => [plan.leaver.mail || plan.leaver.userPrincipalName, r.group.displayName, r.group.mail, r.isOwner ? 'Yes' : 'No', r.isOwner && r.otherOwners === 0 ? 'Yes' : 'No', r.isMember ? 'Yes' : 'No', r.otherOwners, r.group.id])
      )
    }
  >
    Export
  </DefaultButton>
);
