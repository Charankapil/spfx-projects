import * as React from 'react';
import { DefaultButton, DetailsList, DetailsListLayoutMode, Dropdown, IColumn, MessageBar, MessageBarType, Selection, SelectionMode, SearchBox } from '@fluentui/react';
import { IRecycleItem } from '../../models';
import { downloadCsv } from '../../services/exportCsv';
import { formatBytes, formatDateTime } from '../../services/format';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { useLoader } from '../shared/hooks';
import { Card, Empty, ErrorBar, Kpi, Loading, Pill, ViewHeader } from '../shared/ui';

const TOP = 300;
const MAX_ACTION = 100;

/** Recycle bin of the site collection (newest 300 items) with bulk restore / delete. */
export const RecycleView: React.FC = () => {
  const ctx = useAdmin();
  const { api, target } = ctx;
  const loader = useLoader(() => api.getRecycleBin(target.webUrl, TOP), [target.webUrl]);
  const [text, setText] = React.useState('');
  const [stage, setStage] = React.useState('all');
  const [selected, setSelected] = React.useState<IRecycleItem[]>([]);
  const [busy, setBusy] = React.useState(false);
  const selection = React.useMemo(() => new Selection({ getKey: (item) => (item as IRecycleItem).id, onSelectionChanged: () => setSelected(selection.getSelection() as IRecycleItem[]) }), []);

  const q = text.trim().toLowerCase();
  // Memoised: DetailsList resets its selection whenever it receives a new items array.
  const rows = React.useMemo(
    () => (loader.data || []).filter((i) => (stage === 'all' || String(i.stage) === stage) && (!q || i.leafName.toLowerCase().indexOf(q) >= 0 || i.title.toLowerCase().indexOf(q) >= 0 || i.deletedBy.toLowerCase().indexOf(q) >= 0)),
    [loader.data, stage, q]
  );

  if (loader.loading && !loader.data) {
    return <Loading text="Loading the recycle bin…" />;
  }
  if (loader.error || !loader.data) {
    return <ErrorBar error={loader.error || 'Nothing loaded.'} onRetry={loader.reload} />;
  }
  const items = loader.data;
  const capped = items.length >= TOP;
  const total = items.reduce((s, i) => s + i.size, 0);

  const run = async (verb: 'restore' | 'delete'): Promise<void> => {
    const list = selected.slice(0, MAX_ACTION);
    if (!list.length) {
      return;
    }
    const ok = await ctx.confirm({
      title: verb === 'restore' ? 'Restore items' : 'Delete permanently',
      message:
        verb === 'restore'
          ? `Restore ${list.length} item${list.length === 1 ? '' : 's'} to their original location?`
          : `Permanently delete ${list.length} item${list.length === 1 ? '' : 's'}? This cannot be undone.`,
      confirmText: verb === 'restore' ? 'Restore' : 'Delete forever',
      danger: verb === 'delete'
    });
    if (!ok) {
      return;
    }
    setBusy(true);
    let done = 0;
    for (const it of list) {
      const name = it.leafName || it.title;
      try {
        if (verb === 'restore') {
          await api.restoreRecycleItem(target.webUrl, it.id);
        } else {
          await api.deleteRecycleItem(target.webUrl, it.id);
        }
        ctx.log(verb === 'restore' ? 'Restore from recycle bin' : 'Delete from recycle bin', name, true);
        done++;
      } catch (e) {
        ctx.log(verb === 'restore' ? 'Restore from recycle bin' : 'Delete from recycle bin', name, false, (e as Error).message);
      }
    }
    setBusy(false);
    selection.setAllSelected(false);
    ctx.notify(`${done} of ${list.length} ${verb === 'restore' ? 'restored' : 'deleted'}.`, done === list.length ? 'success' : 'error');
    loader.reload();
  };

  const columns: IColumn[] = [
    { key: 'name', name: 'Name', minWidth: 200, isResizable: true, onRender: (i: IRecycleItem) => <strong>{i.leafName || i.title}</strong> },
    { key: 'dir', name: 'Original location', minWidth: 200, isResizable: true, onRender: (i: IRecycleItem) => <span className={styles.muted}>/{i.dirName}</span> },
    { key: 'size', name: 'Size', minWidth: 70, onRender: (i: IRecycleItem) => formatBytes(i.size) },
    { key: 'by', name: 'Deleted by', minWidth: 120, onRender: (i: IRecycleItem) => i.deletedBy },
    { key: 'when', name: 'Deleted', minWidth: 140, onRender: (i: IRecycleItem) => formatDateTime(i.deletedDate) },
    { key: 'stage', name: 'Stage', minWidth: 90, onRender: (i: IRecycleItem) => (i.stage === 2 ? <Pill kind="warning">Second stage</Pill> : <Pill>First stage</Pill>) }
  ];

  return (
    <div className={styles.view}>
      <ViewHeader title="Recycle bin" hint="Restore accidental deletions, or free storage by deleting for good.">
        <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={loader.reload} disabled={loader.loading || busy}>
          Refresh
        </DefaultButton>
        <DefaultButton iconProps={{ iconName: 'Download' }} onClick={() => downloadCsv('recycle-bin.csv', ['Name', 'Location', 'Size', 'Deleted by', 'Deleted', 'Stage'], rows.map((i) => [i.leafName || i.title, i.dirName, i.size, i.deletedBy, i.deletedDate, i.stage]))}>
          Export CSV
        </DefaultButton>
      </ViewHeader>

      {capped && (
        <MessageBar messageBarType={MessageBarType.info}>Showing the {TOP} most recently deleted items. Restore or purge some to see older ones. (Larger pulls are avoided on purpose to keep SharePoint responsive.)</MessageBar>
      )}

      <div className={styles.kpiRow}>
        <Kpi label="Items shown" value={`${items.length}${capped ? '+' : ''}`} />
        <Kpi label="Size shown" value={formatBytes(total)} />
        <Kpi label="Second stage" value={String(items.filter((i) => i.stage === 2).length)} sub="site collection recycle bin" />
      </div>

      <div className={styles.filters}>
        <SearchBox className={styles.field} placeholder="Search name or person" value={text} onChange={(_, v) => setText(v || '')} />
        <Dropdown
          className={styles.field}
          label="Stage"
          selectedKey={stage}
          onChange={(_, o) => setStage(String(o ? o.key : 'all'))}
          options={[
            { key: 'all', text: 'All' },
            { key: '1', text: 'First stage' },
            { key: '2', text: 'Second stage' }
          ]}
        />
        {selected.length > 0 && (
          <>
            <span className={styles.muted}>{selected.length} selected{selected.length > MAX_ACTION ? ` (first ${MAX_ACTION} will be used)` : ''}</span>
            <DefaultButton disabled={busy} iconProps={{ iconName: 'Undo' }} onClick={() => run('restore')}>
              Restore
            </DefaultButton>
            <DefaultButton disabled={busy} iconProps={{ iconName: 'Delete' }} onClick={() => run('delete')}>
              Delete permanently
            </DefaultButton>
          </>
        )}
      </div>

      <Card>
        {rows.length === 0 ? (
          <Empty text={items.length ? 'Nothing matches.' : 'The recycle bin is empty.'} />
        ) : (
          <DetailsList items={rows} columns={columns} selection={selection} selectionMode={SelectionMode.multiple} layoutMode={DetailsListLayoutMode.justified} />
        )}
      </Card>
    </div>
  );
};
