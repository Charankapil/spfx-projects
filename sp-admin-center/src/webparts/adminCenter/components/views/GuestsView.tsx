import * as React from 'react';
import { DefaultButton, DetailsList, DetailsListLayoutMode, Dropdown, IColumn, MessageBar, MessageBarType, PrimaryButton, SearchBox, Selection, SelectionMode } from '@fluentui/react';
import { hasScope } from '../../services/GraphApi';
import { downloadCsv } from '../../services/exportCsv';
import { daysAgo, formatDate } from '../../services/format';
import { IGuest, isStaleGuest, STALE_DAYS } from '../../services/GraphApi';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { GraphGate, graphErrorText, useGraphScopes } from '../shared/graph';
import { Card, Empty, Pill, StatTile, ViewHeader } from '../shared/ui';

type GuestFilter = 'all' | 'stale' | 'never' | 'pending' | 'disabled';

const MAX_GUESTS = 100;

/**
 * Tenant-wide guest review (Microsoft Graph, User.Read.All; last sign-in needs AuditLog.Read.All).
 * Read-only: guests are removed from sites in People & permissions, and from the directory by an
 * Entra ID administrator using the exported list.
 */
export const GuestsView: React.FC = () => {
  const ctx = useAdmin();
  const [state, setState] = React.useState<{ guests?: IGuest[]; signInKnown?: boolean; truncated?: boolean; loading: boolean; progress: number; error?: string }>({ loading: false, progress: 0 });
  const [show, setShow] = React.useState<GuestFilter>('all');
  const [query, setQuery] = React.useState('');
  const access = useGraphScopes();
  const canRemove = !!access.scopes && hasScope(access.scopes, 'Group.ReadWrite.All');
  const [busy, setBusy] = React.useState<string | undefined>();
  const [selectedCount, setSelectedCount] = React.useState(0);
  const selection = React.useMemo(() => new Selection({ getKey: (g) => (g as IGuest).id, onSelectionChanged: () => setSelectedCount(selection.getSelectedCount()) }), []);

  /**
   * Removes the selected guests from every Microsoft 365 group (and team) they are in, as member
   * and as owner. Their guest account stays (deleting accounts needs a directory role).
   */
  const removeFromGroups = async (): Promise<void> => {
    const picked = (selection.getSelection() as IGuest[]).slice(0, MAX_GUESTS);
    if (!picked.length) {
      return;
    }
    const ok = await ctx.confirm({
      title: 'Remove guests from groups',
      message: `Remove ${picked.length} guest${picked.length === 1 ? '' : 's'} from every Microsoft 365 group and team they belong to? They lose access to those sites, teams and mailboxes. Their guest accounts are kept, and direct access to SharePoint-only sites is not changed.`,
      confirmText: 'Remove',
      danger: true
    });
    if (!ok) {
      return;
    }
    let groupsDone = 0;
    const failed: string[] = [];
    for (let i = 0; i < picked.length; i++) {
      const g = picked[i];
      setBusy(`Removing… guest ${i + 1} of ${picked.length}`);
      try {
        const [memberOf, owned] = await Promise.all([ctx.graph.userGroups(g.id, 'memberOf'), ctx.graph.userGroups(g.id, 'ownedObjects')]);
        for (const grp of owned) {
          try {
            await ctx.graph.removeFromGroup(grp.id, 'owners', g.id);
          } catch (e) {
            failed.push(`${g.displayName} as owner of ${grp.displayName}`);
          }
        }
        for (const grp of memberOf) {
          try {
            await ctx.graph.removeFromGroup(grp.id, 'members', g.id);
            groupsDone++;
          } catch (e) {
            failed.push(`${g.displayName} from ${grp.displayName}`);
          }
        }
        ctx.log('Remove guest from groups', `${g.displayName}: ${memberOf.length} groups`, true);
      } catch (e) {
        failed.push(`${g.displayName} (${graphErrorText(e).substring(0, 80)})`);
        ctx.log('Remove guest from groups', g.displayName, false, graphErrorText(e));
      }
    }
    setBusy(undefined);
    selection.setAllSelected(false);
    ctx.notify(failed.length ? `${groupsDone} removals done; ${failed.length} failed: ${failed.slice(0, 4).join('; ')}${failed.length > 4 ? '…' : ''}` : `Removed ${picked.length} guest${picked.length === 1 ? '' : 's'} from ${groupsDone} group membership${groupsDone === 1 ? '' : 's'}.`, failed.length ? 'error' : 'success');
  };

  const load = async (): Promise<void> => {
    setState({ loading: true, progress: 0 });
    try {
      const res = await ctx.graph.listGuests((n) => setState((s) => ({ ...s, progress: n })));
      setState({ loading: false, progress: res.guests.length, ...res });
    } catch (e) {
      setState({ loading: false, progress: 0, error: graphErrorText(e) });
    }
  };

  const guests = state.guests || [];
  const signInKnown = !!state.signInKnown;
  const counts = React.useMemo(
    () => ({
      stale: guests.filter((g) => isStaleGuest(g, signInKnown)).length,
      never: guests.filter((g) => !g.lastSignIn).length,
      pending: guests.filter((g) => g.state === 'PendingAcceptance').length,
      disabled: guests.filter((g) => !g.enabled).length
    }),
    [guests, signInKnown]
  );
  const rows = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    return guests
      .filter((g) =>
        show === 'stale' ? isStaleGuest(g, signInKnown) : show === 'never' ? !g.lastSignIn : show === 'pending' ? g.state === 'PendingAcceptance' : show === 'disabled' ? !g.enabled : true
      )
      .filter((g) => !q || (g.displayName + ' ' + g.mail + ' ' + g.userPrincipalName).toLowerCase().indexOf(q) >= 0)
      .slice(0, 3000);
  }, [guests, show, query, signInKnown]);

  const columns: IColumn[] = [
    { key: 'n', name: 'Guest', minWidth: 180, isResizable: true, onRender: (g: IGuest) => <span title={g.userPrincipalName}>{g.displayName}</span> },
    { key: 'm', name: 'E-mail', minWidth: 200, isResizable: true, onRender: (g: IGuest) => g.mail },
    { key: 'c', name: 'Invited', minWidth: 90, onRender: (g: IGuest) => formatDate(g.created) },
    { key: 's', name: 'Invitation', minWidth: 110, onRender: (g: IGuest) => (g.state === 'PendingAcceptance' ? <Pill kind="warning">Not accepted</Pill> : g.state === 'Accepted' ? 'Accepted' : g.state || '—') },
    ...(signInKnown
      ? [{ key: 'l', name: 'Last sign-in', minWidth: 110, onRender: (g: IGuest) => (g.lastSignIn ? `${formatDate(g.lastSignIn)} (${daysAgo(g.lastSignIn)}d)` : <Pill kind="critical">Never</Pill>) } as IColumn]
      : []),
    { key: 'e', name: 'Account', minWidth: 80, onRender: (g: IGuest) => (g.enabled ? 'Enabled' : <Pill>Disabled</Pill>) }
  ];

  return (
    <div className={styles.view}>
      <ViewHeader title="Guest users" hint="Every guest account in the tenant, with invitation state and (when available) last sign-in, to find guests who no longer need access." />
      <GraphGate scopes={['User.Read.All']} feature="Guest review">
        <Card>
          <div className={styles.actions}>
            <PrimaryButton iconProps={{ iconName: 'People' }} onClick={load} disabled={state.loading}>
              {state.loading ? `Loading… ${state.progress.toLocaleString()} guests` : state.guests ? 'Reload' : 'Load guests'}
            </PrimaryButton>
            {state.guests && (
              <DefaultButton
                iconProps={{ iconName: 'Download' }}
                onClick={() =>
                  downloadCsv('guests.csv', ['Name', 'E-mail', 'UPN', 'Invited', 'Invitation state', 'Last sign-in', 'Enabled', 'Stale', 'Object id'], guests.map((g) => [g.displayName, g.mail, g.userPrincipalName, g.created, g.state, signInKnown ? g.lastSignIn || 'Never' : 'n/a', g.enabled ? 'Yes' : 'No', isStaleGuest(g, signInKnown) ? 'Yes' : 'No', g.id]))
                }
              >
                Export
              </DefaultButton>
            )}
          </div>
          <p className={styles.muted} style={{ marginBottom: 0 }}>
            Read in pages of 100 to 500, paced; only when you click Load. Select guests (for example the stale ones) to remove them from every Microsoft 365 group and team at once; direct access to SharePoint-only sites is removed in People &amp; access (Bulk remove accepts this exported list).
          </p>
          {state.error && <MessageBar messageBarType={MessageBarType.error} isMultiline>{state.error}</MessageBar>}
          {state.guests && !signInKnown && (
            <MessageBar messageBarType={MessageBarType.info} isMultiline>
              Last sign-in dates are not available (needs the AuditLog.Read.All permission, Microsoft Entra ID P1 and a Reports Reader or Global Reader role). Stale means: invitation never accepted after {STALE_DAYS} days.
            </MessageBar>
          )}
        </Card>
        {state.guests && (
          <>
            <div className={styles.statGrid}>
              <StatTile icon="People" label="Guests" value={guests.length.toLocaleString()} onClick={() => setShow('all')} tone="brand" />
              <StatTile icon="Clock" label={`Stale (${STALE_DAYS}+ days)`} value={counts.stale.toLocaleString()} tone={counts.stale ? 'warning' : 'good'} onClick={() => setShow('stale')} />
              {signInKnown && <StatTile icon="Blocked2" label="Never signed in" value={counts.never.toLocaleString()} onClick={() => setShow('never')} />}
              <StatTile icon="Mail" label="Invitation not accepted" value={counts.pending.toLocaleString()} onClick={() => setShow('pending')} />
              <StatTile icon="UserRemove" label="Disabled" value={counts.disabled.toLocaleString()} onClick={() => setShow('disabled')} />
            </div>
            <Card
              title={`${rows.length.toLocaleString()} guests`}
              right={
                <Dropdown
                  ariaLabel="Show"
                  selectedKey={show}
                  onChange={(_, o) => o && setShow(o.key as GuestFilter)}
                  options={[
                    { key: 'all', text: 'All guests' },
                    { key: 'stale', text: 'Stale' },
                    ...(signInKnown ? [{ key: 'never', text: 'Never signed in' }] : []),
                    { key: 'pending', text: 'Not accepted' },
                    { key: 'disabled', text: 'Disabled' }
                  ]}
                  styles={{ root: { width: 170 } }}
                />
              }
            >
              <div className={styles.actions} style={{ marginBottom: 8 }}>
                <SearchBox placeholder="Filter by name or e-mail" value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { width: 300 } }} />
                {canRemove && (
                  <DefaultButton iconProps={{ iconName: 'UserRemove' }} onClick={removeFromGroups} disabled={!!busy || selectedCount === 0} styles={{ root: { borderColor: '#c4314b', color: '#a4262c' } }}>
                    {busy || `Remove ${selectedCount ? Math.min(selectedCount, MAX_GUESTS) : ''} selected from all groups`}
                  </DefaultButton>
                )}
              </div>
              {rows.length === 0 ? (
                <Empty text="No guests match." />
              ) : (
                <DetailsList items={rows} columns={columns} selection={selection} selectionMode={canRemove ? SelectionMode.multiple : SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} />
              )}
              {state.truncated && <p className={styles.muted}>Stopped at 100,000 guests.</p>}
            </Card>
          </>
        )}
      </GraphGate>
    </div>
  );
};
