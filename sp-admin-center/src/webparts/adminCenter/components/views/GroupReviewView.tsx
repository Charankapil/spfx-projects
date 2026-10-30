import * as React from 'react';
import { DefaultButton, DetailsList, DetailsListLayoutMode, Dropdown, IColumn, MessageBar, MessageBarType, PrimaryButton, SearchBox, Selection, SelectionMode, TextField } from '@fluentui/react';
import { downloadCsv } from '../../services/exportCsv';
import { formatDate } from '../../services/format';
import { hasScope, IGroupActivity, IM365Group } from '../../services/GraphApi';
import { REVIEW_FLAGS, ReviewFlag, reviewFlags } from '../../services/m365Review';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { GraphGate, graphErrorText, useGraphScopes } from '../shared/graph';
import { manageGroupSite } from '../shared/groupSite';
import { Card, Empty, Pill, StatTile, ViewHeader } from '../shared/ui';

const MAX_BULK = 100;

interface IReviewed extends IM365Group {
  flags: ReviewFlag[];
  guests?: number;
}

/**
 * Every Microsoft 365 group with what needs attention: no owner, a single owner, a guest as
 * owner, a public group with guests in it, no description. Group.Read(Write).All; guest
 * counts come from the groups usage report when Reports.Read.All is approved.
 */
export const GroupReviewView: React.FC = () => (
  <div className={styles.view}>
    <ViewHeader title="Group review" hint="Every Microsoft 365 group in the tenant with what needs attention, and a quick fix: add an owner to many groups at once." />
    <GraphGate scopes={['Group.Read.All']} feature="Group review">
      <ReviewBody />
    </GraphGate>
  </div>
);

const ReviewBody: React.FC = () => {
  const ctx = useAdmin();
  const access = useGraphScopes();
  const reportsOk = !!access.scopes && hasScope(access.scopes, 'Reports.Read.All');
  const [state, setState] = React.useState<{ groups?: IReviewed[]; truncated?: boolean; loading?: string; error?: string; guestsKnown?: boolean }>({});
  const [show, setShow] = React.useState<ReviewFlag | 'any' | 'all'>('any');
  const [query, setQuery] = React.useState('');
  const [owner, setOwner] = React.useState('');
  const [busy, setBusy] = React.useState(false);
  const [selectedCount, setSelectedCount] = React.useState(0);
  const selection = React.useMemo(() => new Selection({ getKey: (g) => (g as IReviewed).id, onSelectionChanged: () => setSelectedCount(selection.getSelectedCount()) }), []);

  const scan = async (): Promise<void> => {
    setState({ loading: 'Reading groups and owners…' });
    try {
      const res = await ctx.graph.listGroupsWithOwners((n) => setState((s) => ({ ...s, loading: `Reading groups and owners… ${n.toLocaleString()}` })));
      let activity: { [id: string]: IGroupActivity } | undefined;
      if (reportsOk) {
        setState((s) => ({ ...s, loading: 'Reading guest counts from the usage report…' }));
        try {
          activity = {};
          (await ctx.graph.groupActivity('D7')).forEach((a) => {
            if (a.id) {
              (activity as { [id: string]: IGroupActivity })[a.id] = a;
            }
          });
        } catch {
          activity = undefined;
        }
      }
      const groups = res.groups.map((g) => {
        const a = activity ? activity[g.id.toLowerCase()] : undefined;
        return { ...g, flags: reviewFlags(g, a), guests: a ? a.guests : undefined };
      });
      setState({ groups, truncated: res.truncated, guestsKnown: !!activity && Object.keys(activity).length > 0 });
    } catch (e) {
      setState({ error: graphErrorText(e) });
    }
  };

  const groups = state.groups || [];
  const counts = React.useMemo(() => {
    const c: { [f: string]: number } = {};
    REVIEW_FLAGS.forEach((f) => (c[f] = groups.filter((g) => g.flags.indexOf(f) >= 0).length));
    return c;
  }, [groups]);
  const rows = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    return groups
      .filter((g) => (show === 'all' ? true : show === 'any' ? g.flags.some((f) => f !== 'No description') : g.flags.indexOf(show) >= 0))
      .filter((g) => !q || (g.displayName + ' ' + (g.mail || '')).toLowerCase().indexOf(q) >= 0);
  }, [groups, show, query]);
  // A stable array keeps the selection while typing in the owner box.
  const visible = React.useMemo(() => rows.slice(0, 2000), [rows]);

  const addOwner = async (): Promise<void> => {
    const picked = selection.getSelection() as IReviewed[];
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
      const ok = await ctx.confirm({ title: 'Add owner', message: `Add ${user.displayName} as an owner of ${Math.min(picked.length, MAX_BULK)} group${picked.length === 1 ? '' : 's'}?`, confirmText: 'Add owner' });
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
          g.owners = (g.owners || []).concat({ ...user, type: 'user' });
          g.ownerCount = (g.ownerCount || 0) + 1;
          g.flags = reviewFlags(g, g.guests === undefined ? undefined : ({ guests: g.guests } as IGroupActivity));
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
    { key: 'n', name: 'Group', minWidth: 200, isResizable: true, onRender: (g: IReviewed) => <span title={g.mail}>{g.displayName}</span> },
    {
      key: 'f',
      name: 'Needs attention',
      minWidth: 220,
      isResizable: true,
      onRender: (g: IReviewed) =>
        g.flags.length ? (
          <span style={{ display: 'flex', gap: 4, flexWrap: 'wrap' }}>
            {g.flags.map((f) => (
              <Pill key={f} kind={f === 'No owner' || f === 'Guest owner' ? 'critical' : f === 'No description' ? undefined : 'warning'}>
                {f}
              </Pill>
            ))}
          </span>
        ) : (
          <Pill kind="good">OK</Pill>
        )
    },
    { key: 'o', name: 'Owners', minWidth: 60, onRender: (g: IReviewed) => String(g.ownerCount || 0) },
    { key: 'g', name: 'Guests', minWidth: 60, onRender: (g: IReviewed) => (g.guests === undefined ? '—' : String(g.guests)) },
    { key: 'v', name: 'Privacy', minWidth: 70, onRender: (g: IReviewed) => g.visibility || '' },
    { key: 't', name: 'Team', minWidth: 50, onRender: (g: IReviewed) => (g.isTeam ? 'Yes' : '') },
    { key: 'c', name: 'Created', minWidth: 90, onRender: (g: IReviewed) => formatDate(g.createdDateTime ? new Date(g.createdDateTime) : undefined) },
    {
      key: 'a',
      name: '',
      minWidth: 90,
      onRender: (g: IReviewed) => (
        <button type="button" className={styles.link} onClick={() => manageGroupSite(ctx, g.id)}>
          Manage site
        </button>
      )
    }
  ];

  return (
    <>
      <Card>
        <p className={styles.muted} style={{ marginTop: 0 }}>
          Reads every Microsoft 365 group with its owners (100 groups per request, about 2 minutes for 17,000 groups){reportsOk ? ', plus one usage report for guest counts' : ''}. It only runs when you click Scan.
        </p>
        <div className={styles.actions}>
          <PrimaryButton iconProps={{ iconName: 'Search' }} onClick={scan} disabled={!!state.loading}>
            {state.loading || (state.groups ? 'Scan again' : 'Scan all groups')}
          </PrimaryButton>
          {state.groups && (
            <DefaultButton
              iconProps={{ iconName: 'Download' }}
              onClick={() =>
                downloadCsv(
                  'm365-group-review.csv',
                  ['Group', 'E-mail', 'Needs attention', 'Owners (count)', 'Owners', 'Guests', 'Privacy', 'Team', 'Description', 'Created', 'Group id'],
                  groups.map((g) => [g.displayName, g.mail, g.flags.join('; '), g.ownerCount, (g.owners || []).map((o) => o.mail || o.userPrincipalName || o.displayName).join(';'), g.guests, g.visibility, g.isTeam ? 'Yes' : 'No', g.description, g.createdDateTime, g.id])
                )
              }
            >
              Export all
            </DefaultButton>
          )}
        </div>
        {state.error && <MessageBar messageBarType={MessageBarType.error} isMultiline>{state.error}</MessageBar>}
        {state.truncated && <MessageBar messageBarType={MessageBarType.warning}>Stopped at 50,000 groups.</MessageBar>}
        {state.groups && !state.guestsKnown && (
          <MessageBar messageBarType={MessageBarType.info} isMultiline>
            Guest counts are not available ({reportsOk ? 'the usage report hides group ids while names are concealed in reports' : 'needs Reports.Read.All'}), so &ldquo;Public with guests&rdquo; is not checked.
          </MessageBar>
        )}
      </Card>
      {state.groups && (
        <>
          <div className={styles.statGrid}>
            <StatTile icon="Group" label="Microsoft 365 groups" value={groups.length.toLocaleString()} tone="brand" onClick={() => setShow('all')} />
            <StatTile icon="Blocked2" label="No owner" value={counts['No owner'].toLocaleString()} tone={counts['No owner'] ? 'critical' : 'good'} onClick={() => setShow('No owner')} />
            <StatTile icon="Contact" label="One owner" value={counts['One owner'].toLocaleString()} tone={counts['One owner'] ? 'warning' : 'good'} onClick={() => setShow('One owner')} />
            <StatTile icon="UserWarning" label="Guest owner" value={counts['Guest owner'].toLocaleString()} tone={counts['Guest owner'] ? 'critical' : 'good'} onClick={() => setShow('Guest owner')} />
            <StatTile icon="Globe" label="Public with guests" value={state.guestsKnown ? counts['Public with guests'].toLocaleString() : '—'} onClick={() => setShow('Public with guests')} />
            <StatTile icon="TextDocument" label="No description" value={counts['No description'].toLocaleString()} onClick={() => setShow('No description')} />
          </div>
          <Card
            title={`${rows.length.toLocaleString()} groups`}
            right={
              <Dropdown
                ariaLabel="Show"
                selectedKey={show}
                onChange={(_, o) => o && setShow(o.key as ReviewFlag | 'any' | 'all')}
                options={[{ key: 'any', text: 'Needs attention' }, ...REVIEW_FLAGS.map((f) => ({ key: f, text: f })), { key: 'all', text: 'All groups' }]}
                styles={{ root: { width: 190 } }}
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
