import * as React from 'react';
import { DefaultButton, DetailsList, DetailsListLayoutMode, Dropdown, IColumn, MessageBar, MessageBarType, PrimaryButton, SearchBox, SelectionMode } from '@fluentui/react';
import { downloadCsv } from '../../services/exportCsv';
import { daysAgo, formatDate } from '../../services/format';
import { IGuest, isStaleGuest, STALE_DAYS } from '../../services/GraphApi';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { GraphGate, graphErrorText } from '../shared/graph';
import { Card, Empty, Pill, StatTile, ViewHeader } from '../shared/ui';

type GuestFilter = 'all' | 'stale' | 'never' | 'pending' | 'disabled';

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
            Read in pages of 100 to 500, paced; only when you click Load. Remove a guest from a site in People &amp; permissions (Bulk remove accepts this exported list).
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
              <SearchBox placeholder="Filter by name or e-mail" value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { maxWidth: 320, marginBottom: 8 } }} />
              {rows.length === 0 ? <Empty text="No guests match." /> : <DetailsList items={rows} columns={columns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} />}
              {state.truncated && <p className={styles.muted}>Stopped at 100,000 guests.</p>}
            </Card>
          </>
        )}
      </GraphGate>
    </div>
  );
};
