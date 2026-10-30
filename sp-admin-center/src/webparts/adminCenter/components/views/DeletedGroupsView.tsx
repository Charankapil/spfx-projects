import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, MessageBar, MessageBarType, PrimaryButton, SearchBox, Selection, SelectionMode } from '@fluentui/react';
import { formatDateTime } from '../../services/format';
import { IDeletedGroup } from '../../services/GraphApi';
import { daysSince } from '../../services/m365Review';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { GraphGate, graphErrorText } from '../shared/graph';
import { useLoader } from '../shared/hooks';
import { Card, Empty, ErrorBar, Loading, Pill, ViewHeader } from '../shared/ui';

const KEEP_DAYS = 30;

/**
 * Microsoft 365 groups deleted in the last 30 days (Group.ReadWrite.All). Restoring a group
 * brings back its SharePoint site, mailbox, Planner plans and team.
 */
export const DeletedGroupsView: React.FC = () => (
  <div className={styles.view}>
    <ViewHeader title="Deleted groups" hint="Microsoft 365 groups deleted in the last 30 days. Restoring one also brings back its SharePoint site, mailbox, plans and team." />
    <GraphGate scopes={['Group.ReadWrite.All']} feature="Restoring deleted groups">
      <DeletedBody />
    </GraphGate>
  </div>
);

const DeletedBody: React.FC = () => {
  const ctx = useAdmin();
  const loader = useLoader(() => ctx.graph.deletedGroups(), []);
  const [query, setQuery] = React.useState('');
  const [busy, setBusy] = React.useState(false);
  const [restored, setRestored] = React.useState<string[]>([]);
  const [selectedCount, setSelectedCount] = React.useState(0);
  const selection = React.useMemo(() => new Selection({ getKey: (g) => (g as IDeletedGroup).id, onSelectionChanged: () => setSelectedCount(selection.getSelectedCount()) }), []);

  const groups = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    return (loader.data || [])
      .filter((g) => restored.indexOf(g.id) < 0)
      .filter((g) => !q || (g.displayName + ' ' + g.mail).toLowerCase().indexOf(q) >= 0)
      .sort((a, b) => (b.deleted ? b.deleted.getTime() : 0) - (a.deleted ? a.deleted.getTime() : 0));
  }, [loader.data, query, restored]);

  if (loader.loading && !loader.data) {
    return <Loading text="Loading deleted groups…" />;
  }
  if (loader.error || !loader.data) {
    return <ErrorBar error={loader.error || 'Nothing loaded.'} onRetry={loader.reload} />;
  }

  const restore = async (): Promise<void> => {
    const picked = selection.getSelection() as IDeletedGroup[];
    if (!picked.length) {
      return;
    }
    const ok = await ctx.confirm({ title: 'Restore groups', message: `Restore ${picked.length === 1 ? `"${picked[0].displayName}"` : `${picked.length} groups`} with their sites, mailboxes and teams? It can take a few minutes for everything to come back.`, confirmText: 'Restore' });
    if (!ok) {
      return;
    }
    setBusy(true);
    const done: string[] = [];
    const failed: string[] = [];
    for (const g of picked.slice(0, 50)) {
      try {
        await ctx.graph.restoreGroup(g.id);
        ctx.log('Restore Microsoft 365 group', g.displayName, true);
        done.push(g.id);
      } catch (e) {
        failed.push(`${g.displayName} (${graphErrorText(e).substring(0, 80)})`);
        ctx.log('Restore Microsoft 365 group', g.displayName, false, graphErrorText(e));
      }
    }
    setRestored((r) => r.concat(done));
    selection.setAllSelected(false);
    setBusy(false);
    ctx.notify(failed.length ? `${done.length} restored; failed: ${failed.join('; ')}` : `${done.length} group${done.length === 1 ? '' : 's'} restored. The site and team may take a few minutes to reappear.`, failed.length ? 'error' : 'success');
  };

  const columns: IColumn[] = [
    { key: 'n', name: 'Group', minWidth: 220, isResizable: true, onRender: (g: IDeletedGroup) => <span title={g.mail}>{g.displayName}</span> },
    { key: 'm', name: 'E-mail', minWidth: 200, isResizable: true, onRender: (g: IDeletedGroup) => g.mail },
    { key: 'd', name: 'Deleted', minWidth: 150, onRender: (g: IDeletedGroup) => formatDateTime(g.deleted) },
    {
      key: 'l',
      name: 'Days left',
      minWidth: 80,
      onRender: (g: IDeletedGroup) => {
        const left = KEEP_DAYS - (daysSince(g.deleted) || 0);
        return left <= 5 ? <Pill kind="critical">{Math.max(0, left)}</Pill> : String(left);
      }
    },
    { key: 'v', name: 'Privacy', minWidth: 70, onRender: (g: IDeletedGroup) => g.visibility }
  ];

  return (
    <Card title={`${groups.length.toLocaleString()} deleted groups`}>
      <div className={styles.actions} style={{ marginBottom: 8 }}>
        <SearchBox placeholder="Filter by name or e-mail" value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { width: 280 } }} />
        <PrimaryButton iconProps={{ iconName: 'Undo' }} onClick={restore} disabled={busy || selectedCount === 0}>
          Restore {selectedCount || ''} selected
        </PrimaryButton>
      </div>
      {restored.length > 0 && <MessageBar messageBarType={MessageBarType.success}>{restored.length} restored in this session.</MessageBar>}
      {groups.length === 0 ? (
        <Empty text="No deleted Microsoft 365 groups in the last 30 days." />
      ) : (
        <DetailsList items={groups} columns={columns} selection={selection} selectionMode={SelectionMode.multiple} layoutMode={DetailsListLayoutMode.justified} />
      )}
      <p className={styles.muted} style={{ fontSize: 12 }}>
        Restoring needs the SharePoint, Groups or Global Administrator role, or ownership of the group. A group can only be restored if its name and e-mail address have not been reused.
      </p>
    </Card>
  );
};
