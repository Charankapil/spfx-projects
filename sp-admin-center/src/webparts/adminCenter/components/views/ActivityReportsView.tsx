import * as React from 'react';
import { DefaultButton, DetailsList, DetailsListLayoutMode, Dropdown, IColumn, MessageBar, MessageBarType, PrimaryButton, SearchBox, Selection, SelectionMode } from '@fluentui/react';
import { downloadCsv } from '../../services/exportCsv';
import { formatBytes, formatDate } from '../../services/format';
import { IGroupActivity, IGroupSummary, ITeamActivity } from '../../services/GraphApi';
import { daysSince } from '../../services/m365Review';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { GraphGate, graphErrorText } from '../shared/graph';
import { manageGroupSite } from '../shared/groupSite';
import { Card, Empty, Pill, StatTile, ViewHeader } from '../shared/ui';

const MAX_DELETE = 50;
const IDLE_OPTIONS = [
  { key: 90, text: 'Idle 90+ days' },
  { key: 180, text: 'Idle 180+ days' },
  { key: 365, text: 'Idle 365+ days' },
  { key: -1, text: 'Never active' },
  { key: 0, text: 'All' }
];

function idleText(d: Date | undefined): React.ReactNode {
  const n = daysSince(d);
  if (n === undefined) {
    return <Pill kind="critical">Never</Pill>;
  }
  return (
    <span>
      {formatDate(d)} <span className={styles.muted}>({n.toLocaleString()} d)</span>
    </span>
  );
}

function matchesIdle(d: Date | undefined, idle: number): boolean {
  if (idle === 0) {
    return true;
  }
  if (idle === -1) {
    return !d;
  }
  const n = daysSince(d);
  return n === undefined || n >= idle;
}

// ---- inactive Microsoft 365 groups -----------------------------------------------------

/**
 * Microsoft 365 groups by last activity (mail, files, Teams), from the groups usage report
 * (Reports.Read.All, one report request). Selected groups can be deleted (Group.ReadWrite.All);
 * a deleted group, with its site, mailbox and team, can be restored for 30 days.
 */
export const InactiveGroupsView: React.FC = () => (
  <div className={styles.view}>
    <ViewHeader title="Inactive groups & Teams" hint="Microsoft 365 groups and teams nobody has used for months, from Microsoft's usage reports. Delete the ones you no longer need; they can be restored for 30 days." />
    <GraphGate scopes={['Reports.Read.All']} feature="Inactive groups">
      <InactiveBody />
    </GraphGate>
  </div>
);

interface IActivityRow extends IGroupActivity {
  isTeam?: boolean;
  visibility?: string;
}

const InactiveBody: React.FC = () => {
  const ctx = useAdmin();
  const [state, setState] = React.useState<{ rows?: IActivityRow[]; loading?: string; error?: string; concealed?: number }>({});
  const [idle, setIdle] = React.useState(180);
  const [query, setQuery] = React.useState('');
  const [busy, setBusy] = React.useState(false);
  const [selectedCount, setSelectedCount] = React.useState(0);
  const selection = React.useMemo(() => new Selection({ getKey: (g) => (g as IActivityRow).id || (g as IActivityRow).name, onSelectionChanged: () => setSelectedCount(selection.getSelectedCount()) }), []);

  const load = async (): Promise<void> => {
    setState({ loading: 'Reading the groups usage report…' });
    try {
      const report = await ctx.graph.groupActivity('D180', (n) => setState((s) => ({ ...s, loading: `Reading the groups usage report… ${n.toLocaleString()}` })));
      let dir: { [id: string]: IGroupSummary } = {};
      try {
        dir = await ctx.graph.groupDirectory();
      } catch {
        /* Teams flag and privacy are optional */
      }
      const live = report.filter((r) => !r.isDeleted);
      const rows = live.filter((r) => !!r.id).map((r) => ({ ...r, isTeam: dir[r.id] ? dir[r.id].isTeam : undefined, visibility: dir[r.id] ? dir[r.id].visibility : r.type }));
      setState({ rows, concealed: live.length - rows.length });
    } catch (e) {
      setState({ error: graphErrorText(e) });
    }
  };

  const all = state.rows || [];
  const rows = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    return all
      .filter((r) => matchesIdle(r.lastActivity, idle))
      .filter((r) => !q || (r.name + ' ' + r.owner).toLowerCase().indexOf(q) >= 0)
      .sort((a, b) => (a.lastActivity ? a.lastActivity.getTime() : 0) - (b.lastActivity ? b.lastActivity.getTime() : 0));
  }, [all, idle, query]);
  // A stable array keeps the selection while filtering state elsewhere changes.
  const visible = React.useMemo(() => rows.slice(0, 3000), [rows]);
  const counts = React.useMemo(() => ({ d90: all.filter((r) => matchesIdle(r.lastActivity, 90)).length, d365: all.filter((r) => matchesIdle(r.lastActivity, 365)).length, never: all.filter((r) => !r.lastActivity).length }), [all]);

  const remove = async (): Promise<void> => {
    const picked = (selection.getSelection() as IActivityRow[]).slice(0, MAX_DELETE);
    if (!picked.length) {
      return;
    }
    const ok = await ctx.confirm({
      title: `Delete ${picked.length} group${picked.length === 1 ? '' : 's'}?`,
      message: `This deletes ${picked.length === 1 ? `"${picked[0].name}"` : `${picked.length} Microsoft 365 groups`} with their SharePoint sites, mailboxes, Planner plans and teams. They can be restored for 30 days (Microsoft 365 > Deleted groups), after which they are gone for good.`,
      confirmText: 'Delete',
      danger: true
    });
    if (!ok) {
      return;
    }
    setBusy(true);
    let done = 0;
    const failed: string[] = [];
    for (const g of picked) {
      try {
        await ctx.graph.deleteGroup(g.id);
        ctx.log('Delete Microsoft 365 group', g.name, true);
        done++;
      } catch (e) {
        failed.push(g.name);
        ctx.log('Delete Microsoft 365 group', g.name, false, graphErrorText(e));
      }
    }
    const gone = picked.filter((p) => failed.indexOf(p.name) < 0).map((p) => p.id);
    setState((s) => ({ ...s, rows: (s.rows || []).filter((r) => gone.indexOf(r.id) < 0) }));
    selection.setAllSelected(false);
    setBusy(false);
    ctx.notify(failed.length ? `${done} deleted; failed: ${failed.slice(0, 5).join(', ')}` : `${done} group${done === 1 ? '' : 's'} deleted. Restore them under Microsoft 365 > Deleted groups within 30 days if needed.`, failed.length ? 'error' : 'success');
  };

  const columns: IColumn[] = [
    { key: 'n', name: 'Group', minWidth: 200, isResizable: true, onRender: (r: IActivityRow) => <span title={r.owner}>{r.name}</span> },
    { key: 'l', name: 'Last activity', minWidth: 150, onRender: (r: IActivityRow) => idleText(r.lastActivity) },
    { key: 'm', name: 'Members', minWidth: 70, onRender: (r: IActivityRow) => r.members.toLocaleString() },
    { key: 'g', name: 'Guests', minWidth: 60, onRender: (r: IActivityRow) => r.guests.toLocaleString() },
    { key: 'f', name: 'Files', minWidth: 70, onRender: (r: IActivityRow) => r.files.toLocaleString() },
    { key: 's', name: 'Site storage', minWidth: 90, onRender: (r: IActivityRow) => formatBytes(r.siteBytes) },
    { key: 't', name: 'Team', minWidth: 50, onRender: (r: IActivityRow) => (r.isTeam ? 'Yes' : '') },
    { key: 'v', name: 'Privacy', minWidth: 70, onRender: (r: IActivityRow) => r.visibility || '' },
    {
      key: 'a',
      name: '',
      minWidth: 90,
      onRender: (r: IActivityRow) => (
        <button type="button" className={styles.link} onClick={() => manageGroupSite(ctx, r.id)}>
          Manage site
        </button>
      )
    }
  ];

  return (
    <>
      <Card>
        <p className={styles.muted} style={{ marginTop: 0 }}>
          One usage report covers every group (Microsoft refreshes it daily; it is usually 1&ndash;2 days behind). Activity means e-mail received, files used in the group&rsquo;s site, or Teams activity.
        </p>
        <div className={styles.actions}>
          <PrimaryButton iconProps={{ iconName: 'Download' }} onClick={load} disabled={!!state.loading}>
            {state.loading || (state.rows ? 'Reload' : 'Load group activity')}
          </PrimaryButton>
          {state.rows && (
            <DefaultButton
              iconProps={{ iconName: 'ExcelDocument' }}
              onClick={() =>
                downloadCsv(
                  'm365-group-activity.csv',
                  ['Group', 'Owner', 'Last activity', 'Days idle', 'Members', 'Guests', 'E-mails received', 'Active files', 'Files', 'Site storage (bytes)', 'Team', 'Privacy', 'Group id'],
                  rows.map((r) => [r.name, r.owner, r.lastActivity ? r.lastActivity.toISOString().substring(0, 10) : 'Never', daysSince(r.lastActivity), r.members, r.guests, r.emailsReceived, r.activeFiles, r.files, r.siteBytes, r.isTeam ? 'Yes' : 'No', r.visibility, r.id])
                )
              }
            >
              Export shown
            </DefaultButton>
          )}
        </div>
        {state.error && <MessageBar messageBarType={MessageBarType.error} isMultiline>{state.error}</MessageBar>}
        {!!state.concealed && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline>
            {state.concealed.toLocaleString()} groups are hidden because names are concealed in usage reports (Microsoft 365 admin center &gt; Settings &gt; Org settings &gt; Reports).
          </MessageBar>
        )}
      </Card>
      {state.rows && (
        <>
          <div className={styles.statGrid}>
            <StatTile icon="Group" label="Active groups in report" value={all.length.toLocaleString()} tone="brand" onClick={() => setIdle(0)} />
            <StatTile icon="Clock" label="Idle 90+ days" value={counts.d90.toLocaleString()} tone={counts.d90 ? 'warning' : 'good'} onClick={() => setIdle(90)} />
            <StatTile icon="Snooze" label="Idle 365+ days" value={counts.d365.toLocaleString()} tone={counts.d365 ? 'critical' : 'good'} onClick={() => setIdle(365)} />
            <StatTile icon="Blocked2" label="Never active" value={counts.never.toLocaleString()} onClick={() => setIdle(-1)} />
          </div>
          <Card
            title={`${rows.length.toLocaleString()} groups`}
            right={<Dropdown ariaLabel="Show" selectedKey={idle} onChange={(_, o) => o && setIdle(Number(o.key))} options={IDLE_OPTIONS} styles={{ root: { width: 170 } }} />}
          >
            <div className={styles.actions} style={{ marginBottom: 8 }}>
              <SearchBox placeholder="Filter by name or owner" value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { width: 280 } }} />
              <DefaultButton
                iconProps={{ iconName: 'Delete' }}
                onClick={remove}
                disabled={busy || selectedCount === 0}
                styles={{ root: { borderColor: '#c4314b', color: '#a4262c' } }}
              >
                Delete {selectedCount ? Math.min(selectedCount, MAX_DELETE) : ''} selected
              </DefaultButton>
            </div>
            {rows.length === 0 ? (
              <Empty text="No groups match." />
            ) : (
              <DetailsList items={visible} columns={columns} selection={selection} selectionMode={SelectionMode.multiple} layoutMode={DetailsListLayoutMode.justified} />
            )}
            <p className={styles.muted} style={{ fontSize: 12 }}>
              Up to {MAX_DELETE} groups per delete. You need to be an owner of the group, or a SharePoint, Groups or Global Administrator.
            </p>
          </Card>
        </>
      )}
    </>
  );
};

// ---- Teams usage ----------------------------------------------------------------------

/** Per-team activity from the Teams usage report (Reports.Read.All, one report request). */
export const TeamsUsageView: React.FC = () => (
  <div className={styles.view}>
    <ViewHeader title="Teams usage" hint="Active users, messages and meetings for every team over the last 30 days, to find teams nobody uses." />
    <GraphGate scopes={['Reports.Read.All']} feature="Teams usage">
      <TeamsBody />
    </GraphGate>
  </div>
);

const TeamsBody: React.FC = () => {
  const ctx = useAdmin();
  const [state, setState] = React.useState<{ rows?: ITeamActivity[]; loading?: string; error?: string; concealed?: number }>({});
  const [show, setShow] = React.useState<'all' | 'unused' | 'quiet'>('unused');
  const [query, setQuery] = React.useState('');

  const load = async (): Promise<void> => {
    setState({ loading: 'Reading the Teams usage report…' });
    try {
      const all = await ctx.graph.teamsActivity('D30', (n) => setState((s) => ({ ...s, loading: `Reading the Teams usage report… ${n.toLocaleString()}` })));
      const rows = all.filter((t) => !!t.name);
      setState({ rows, concealed: all.length - rows.length });
    } catch (e) {
      setState({ error: graphErrorText(e) });
    }
  };

  const all = state.rows || [];
  const rows = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    return all
      .filter((t) => (show === 'unused' ? t.activeUsers === 0 : show === 'quiet' ? t.activeUsers > 0 && t.messages + t.meetings === 0 : true))
      .filter((t) => !q || t.name.toLowerCase().indexOf(q) >= 0)
      .sort((a, b) => a.activeUsers - b.activeUsers || a.messages - b.messages);
  }, [all, show, query]);
  const unused = all.filter((t) => t.activeUsers === 0).length;

  const columns: IColumn[] = [
    { key: 'n', name: 'Team', minWidth: 200, isResizable: true, onRender: (t: ITeamActivity) => t.name },
    { key: 'u', name: 'Active users', minWidth: 90, onRender: (t: ITeamActivity) => (t.activeUsers ? t.activeUsers.toLocaleString() : <Pill kind="warning">0</Pill>) },
    { key: 'c', name: 'Active channels', minWidth: 100, onRender: (t: ITeamActivity) => t.activeChannels.toLocaleString() },
    { key: 'm', name: 'Messages', minWidth: 80, onRender: (t: ITeamActivity) => t.messages.toLocaleString() },
    { key: 'e', name: 'Meetings', minWidth: 80, onRender: (t: ITeamActivity) => t.meetings.toLocaleString() },
    { key: 'g', name: 'Guests', minWidth: 60, onRender: (t: ITeamActivity) => t.guests.toLocaleString() },
    { key: 'l', name: 'Last activity', minWidth: 150, onRender: (t: ITeamActivity) => idleText(t.lastActivity) },
    { key: 'p', name: 'Type', minWidth: 70, onRender: (t: ITeamActivity) => t.type },
    {
      key: 'a',
      name: '',
      minWidth: 90,
      onRender: (t: ITeamActivity) =>
        t.id ? (
          <button type="button" className={styles.link} onClick={() => manageGroupSite(ctx, t.id)}>
            Manage site
          </button>
        ) : null
    }
  ];

  return (
    <>
      <Card>
        <div className={styles.actions}>
          <PrimaryButton iconProps={{ iconName: 'TeamsLogo' }} onClick={load} disabled={!!state.loading}>
            {state.loading || (state.rows ? 'Reload' : 'Load Teams usage')}
          </PrimaryButton>
          {state.rows && (
            <DefaultButton
              iconProps={{ iconName: 'ExcelDocument' }}
              onClick={() =>
                downloadCsv(
                  'teams-usage-30-days.csv',
                  ['Team', 'Active users', 'Active channels', 'Messages', 'Meetings', 'Guests', 'Last activity', 'Type', 'Team id'],
                  rows.map((t) => [t.name, t.activeUsers, t.activeChannels, t.messages, t.meetings, t.guests, t.lastActivity ? t.lastActivity.toISOString().substring(0, 10) : 'Never', t.type, t.id])
                )
              }
            >
              Export shown
            </DefaultButton>
          )}
        </div>
        {state.error && <MessageBar messageBarType={MessageBarType.error} isMultiline>{state.error}</MessageBar>}
        {!!state.concealed && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline>
            {state.concealed.toLocaleString()} teams are hidden because names are concealed in usage reports.
          </MessageBar>
        )}
      </Card>
      {state.rows && (
        <>
          <div className={styles.statGrid}>
            <StatTile icon="TeamsLogo" label="Teams" value={all.length.toLocaleString()} tone="brand" onClick={() => setShow('all')} />
            <StatTile icon="Blocked2" label="No active users (30 days)" value={unused.toLocaleString()} tone={unused ? 'warning' : 'good'} onClick={() => setShow('unused')} />
            <StatTile icon="Chat" label="Users but no messages or meetings" value={all.filter((t) => t.activeUsers > 0 && t.messages + t.meetings === 0).length.toLocaleString()} onClick={() => setShow('quiet')} />
          </div>
          <Card
            title={`${rows.length.toLocaleString()} teams`}
            right={
              <Dropdown
                ariaLabel="Show"
                selectedKey={show}
                onChange={(_, o) => o && setShow(o.key as 'all' | 'unused' | 'quiet')}
                options={[
                  { key: 'unused', text: 'No active users' },
                  { key: 'quiet', text: 'No messages or meetings' },
                  { key: 'all', text: 'All teams' }
                ]}
                styles={{ root: { width: 200 } }}
              />
            }
          >
            <SearchBox placeholder="Filter by name" value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { maxWidth: 300, marginBottom: 8 } }} />
            {rows.length === 0 ? <Empty text="No teams match." /> : <DetailsList items={rows.slice(0, 3000)} columns={columns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} />}
            <p className={styles.muted} style={{ fontSize: 12 }}>
              To remove an unused team, delete its group under Microsoft 365 &gt; Inactive groups (it can be restored for 30 days).
            </p>
          </Card>
        </>
      )}
    </>
  );
};
