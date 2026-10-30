import * as React from 'react';
import { DefaultButton, DetailsList, DetailsListLayoutMode, Dropdown, IColumn, Selection, SelectionMode, SearchBox, Toggle } from '@fluentui/react';
import { IListInfo } from '../../models';
import { downloadCsv } from '../../services/exportCsv';
import { LIST_VIEW_THRESHOLD } from '../../services/HealthEngine';
import { formatDate, formatCompact, relativeTime } from '../../services/format';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { useLoader } from '../shared/hooks';
import { Card, ErrorBar, Kpi, Loading, Pill, ViewHeader } from '../shared/ui';

type Change = { EnableVersioning?: boolean; NoCrawl?: boolean };

/** Lists and libraries of the target web, with the settings admins most often fix. */
export const ContentView: React.FC = () => {
  const ctx = useAdmin();
  const { api, target } = ctx;
  const loader = useLoader(() => api.getLists(target.webUrl), [target.webUrl]);
  const [kind, setKind] = React.useState('all');
  const [text, setText] = React.useState('');
  const [showHidden, setShowHidden] = React.useState(false);
  const [selected, setSelected] = React.useState<IListInfo[]>([]);
  const [busy, setBusy] = React.useState(false);
  const selection = React.useMemo(
    () =>
      new Selection({
        getKey: (item) => (item as IListInfo).id,
        onSelectionChanged: () => setSelected(selection.getSelection() as IListInfo[])
      }),
    []
  );

  const apply = async (items: IListInfo[], changes: Change, label: string): Promise<void> => {
    if (!items.length) {
      return;
    }
    const ok = await ctx.confirm({
      title: label,
      message: `${label} on ${items.length} list${items.length === 1 ? '' : 's'} in ${target.title}? ${items.slice(0, 5).map((i) => i.title).join(', ')}${items.length > 5 ? '…' : ''}`,
      confirmText: 'Apply'
    });
    if (!ok) {
      return;
    }
    setBusy(true);
    let done = 0;
    for (const l of items) {
      try {
        await api.updateList(target.webUrl, l.id, changes);
        ctx.log(label, l.title, true);
        done++;
      } catch (e) {
        ctx.log(label, l.title, false, (e as Error).message);
        ctx.notify(`${l.title}: ${(e as Error).message}`, 'error');
      }
    }
    setBusy(false);
    ctx.notify(`${label}: ${done} of ${items.length} updated.`, done === items.length ? 'success' : 'info');
    loader.reload();
  };

  const q = text.trim().toLowerCase();
  // Memoised: DetailsList resets its selection whenever it receives a new items array.
  const rows = React.useMemo(
    () =>
      (loader.data || []).filter((l) => {
        if (!showHidden && (l.hidden || l.isSystem)) {
          return false;
        }
        if (kind !== 'all' && l.kind !== kind) {
          return false;
        }
        return !q || l.title.toLowerCase().indexOf(q) >= 0;
      }),
    [loader.data, showHidden, kind, q]
  );

  if (loader.loading && !loader.data) {
    return <Loading text="Loading lists and libraries…" />;
  }
  if (loader.error || !loader.data) {
    return <ErrorBar error={loader.error || 'Nothing loaded.'} onRetry={loader.reload} />;
  }

  const all = loader.data;
  const visible = all.filter((l) => !l.hidden && !l.isSystem);
  const big = visible.filter((l) => l.itemCount >= LIST_VIEW_THRESHOLD).length;
  const noVer = visible.filter((l) => l.kind === 'Library' && !l.versioning).length;
  const unique = visible.filter((l) => l.uniquePermissions).length;

  const columns: IColumn[] = [
    {
      key: 'title',
      name: 'Name',
      minWidth: 170,
      isResizable: true,
      onRender: (l: IListInfo) => (
        <a href={l.url} target="_blank" rel="noopener noreferrer" className={styles.link}>
          {l.title}
        </a>
      )
    },
    { key: 'kind', name: 'Type', minWidth: 70, onRender: (l: IListInfo) => l.kind },
    {
      key: 'items',
      name: 'Items',
      minWidth: 80,
      onRender: (l: IListInfo) =>
        l.itemCount >= LIST_VIEW_THRESHOLD ? <Pill kind="warning">{formatCompact(l.itemCount)}</Pill> : <span>{formatCompact(l.itemCount)}</span>
    },
    { key: 'mod', name: 'Last change', minWidth: 100, onRender: (l: IListInfo) => relativeTime(l.lastModified) },
    {
      key: 'ver',
      name: 'Versioning',
      minWidth: 100,
      onRender: (l: IListInfo) =>
        l.kind === 'Other' ? (
          '-'
        ) : (
          <Toggle
            checked={l.versioning}
            disabled={busy}
            ariaLabel={`Version history for ${l.title}`}
            onChange={(_, c) => apply([l], { EnableVersioning: !!c }, c ? 'Turn on version history' : 'Turn off version history')}
          />
        )
    },
    {
      key: 'search',
      name: 'In search',
      minWidth: 90,
      onRender: (l: IListInfo) => (
        <Toggle
          checked={!l.noCrawl}
          disabled={busy}
          ariaLabel={`Appears in search: ${l.title}`}
          onChange={(_, c) => apply([l], { NoCrawl: !c }, c ? 'Show in search results' : 'Hide from search results')}
        />
      )
    },
    { key: 'perm', name: 'Permissions', minWidth: 90, onRender: (l: IListInfo) => (l.uniquePermissions ? <Pill kind="info">Unique</Pill> : <span className={styles.muted}>Inherited</span>) },
    { key: 'created', name: 'Created', minWidth: 90, onRender: (l: IListInfo) => formatDate(l.created) }
  ];

  return (
    <div className={styles.view}>
      <ViewHeader title="Lists & libraries" hint={`Content containers in ${target.title}. Changes apply immediately; select several rows for bulk actions.`}>
        <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={loader.reload} disabled={loader.loading}>
          Refresh
        </DefaultButton>
        <DefaultButton
          iconProps={{ iconName: 'Download' }}
          onClick={() =>
            downloadCsv(
              'lists-and-libraries.csv',
              ['Name', 'Type', 'Items', 'Version history', 'Appears in search', 'Unique permissions', 'Hidden', 'URL', 'Created', 'Last change'],
              rows.map((l) => [l.title, l.kind, l.itemCount, l.versioning, !l.noCrawl, l.uniquePermissions, l.hidden, l.url, l.created, l.lastModified])
            )
          }
        >
          Export CSV
        </DefaultButton>
      </ViewHeader>

      <div className={styles.kpiRow}>
        <Kpi label="Containers" value={visible.length.toLocaleString()} sub={`${visible.filter((l) => l.kind === 'Library').length} libraries`} />
        <Kpi label={`Over ${LIST_VIEW_THRESHOLD.toLocaleString()} items`} value={String(big)} sub="list view threshold" />
        <Kpi label="No version history" value={String(noVer)} sub="libraries" />
        <Kpi label="Unique permissions" value={String(unique)} sub="broken inheritance" />
      </div>

      <div className={styles.filters}>
        <SearchBox className={styles.field} placeholder="Search by name" value={text} onChange={(_, v) => setText(v || '')} />
        <Dropdown
          className={styles.field}
          label="Show"
          selectedKey={kind}
          onChange={(_, o) => setKind(String(o ? o.key : 'all'))}
          options={[
            { key: 'all', text: 'Libraries and lists' },
            { key: 'Library', text: 'Libraries only' },
            { key: 'List', text: 'Lists only' }
          ]}
        />
        <Toggle label="System / hidden" inlineLabel checked={showHidden} onChange={(_, c) => setShowHidden(!!c)} />
        {selected.length > 0 && (
          <>
            <span className={styles.muted}>{selected.length} selected</span>
            <DefaultButton disabled={busy} onClick={() => apply(selected, { EnableVersioning: true }, 'Turn on version history')}>
              Turn on versioning
            </DefaultButton>
            <DefaultButton disabled={busy} onClick={() => apply(selected, { NoCrawl: false }, 'Show in search results')}>
              Show in search
            </DefaultButton>
          </>
        )}
      </div>

      <Card>
        {rows.length === 0 ? (
          <div className={styles.empty}>Nothing matches.</div>
        ) : (
          <DetailsList items={rows} columns={columns} selection={selection} selectionMode={SelectionMode.multiple} layoutMode={DetailsListLayoutMode.justified} />
        )}
      </Card>
    </div>
  );
};
