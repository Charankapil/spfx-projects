import * as React from 'react';
import { Icon, ContextualMenu, IContextualMenuItem, DirectionalHint, Checkbox, Callout, Dialog, DialogFooter, PrimaryButton, DefaultButton, TextField } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { IBoard, IBoardColumn, IBoardGroup, IWorkItem } from '../../models/types';
import { IBoardRights } from '../../services/permissions';
import { IBoardActions } from '../board/useBoard';
import { Cell, columnWidth, COLUMN_TYPES } from '../cells/Cells';
import { IViewState, visibleRows, summarize, isDone, IGroupRows } from '../../engine/viewQuery';
import { Swatches, useConfirm } from '../common/Common';
import { formatRange } from '../../engine/dates';

export interface ITableViewProps {
  board: IBoard;
  items: IWorkItem[];
  rights: IBoardRights;
  view: IViewState;
  setView: (v: IViewState) => void;
  actions: IBoardActions;
  updateCounts: { [itemId: number]: number };
  onOpenItem: (itemId: number) => void;
  onEditLabels: (column: IBoardColumn) => void;
}

const EDGE = 6;
const CHECK = 36;
const TITLE = 340;
const ADD = 44;

function gridTemplate(cols: IBoardColumn[]): string {
  return [EDGE, CHECK, TITLE].concat(cols.map(columnWidth)).concat([ADD]).map(w => w + 'px').join(' ');
}

/* ---------- Title cell with inline rename ---------- */

function TitleCell(props: {
  item: IWorkItem; board: IBoard; canEdit: boolean; isSub: boolean; subCount: number; expanded: boolean;
  updateCount: number; onToggle: () => void; onOpen: () => void; onRename: (t: string) => void; onMenu: (target: HTMLElement) => void;
  draggable: boolean; onDragStart: (e: React.DragEvent) => void;
}): JSX.Element {
  const [editing, setEditing] = React.useState(false);
  const [draft, setDraft] = React.useState(props.item.title);
  const commit = (): void => {
    setEditing(false);
    const t = draft.trim();
    if (t && t !== props.item.title) {
      props.onRename(t);
    }
  };
  return (
    <div
      className={`${styles.cell} ${styles.titleCell} ${props.isSub ? styles.subitemTitle : ''}`}
      draggable={props.draggable && !editing}
      onDragStart={props.onDragStart}
    >
      {!props.isSub && (
        <button type="button" className={styles.iconBtn} style={{ width: 20, visibility: props.subCount > 0 ? 'visible' : 'hidden' }}
          onClick={props.onToggle} aria-label={props.expanded ? 'Hide subitems' : `Show ${props.subCount} subitems`} aria-expanded={props.expanded}>
          <Icon iconName={props.expanded ? 'ChevronDown' : 'ChevronRight'} style={{ fontSize: 10 }} />
        </button>
      )}
      {editing ? (
        <input className={styles.inlineInput} autoFocus value={draft} onChange={e => setDraft(e.target.value)} onBlur={commit}
          onKeyDown={e => { if (e.key === 'Enter') { commit(); } else if (e.key === 'Escape') { setEditing(false); } }} aria-label="Item name" />
      ) : (
        <span className={styles.titleText} onClick={() => { if (props.canEdit) { setDraft(props.item.title); setEditing(true); } }}
          title={props.item.title} role={props.canEdit ? 'button' : undefined} tabIndex={props.canEdit ? 0 : undefined}
          onKeyDown={e => { if (e.key === 'Enter' && props.canEdit) { setDraft(props.item.title); setEditing(true); } }}>
          {props.item.title || <span className={styles.muted}>Untitled</span>}
        </span>
      )}
      {!props.isSub && props.subCount > 0 && !editing && <span className={styles.pill} title={`${props.subCount} subitems`}>{props.subCount}</span>}
      <button type="button" className={styles.iconBtn} style={{ width: 24 }} onClick={e => props.onMenu(e.currentTarget)} aria-label="Item actions">
        <Icon iconName="More" />
      </button>
      <button type="button" className={`${styles.openBtn} ${props.updateCount > 0 ? styles.hasUpdates : ''}`} onClick={props.onOpen}
        aria-label={`Open ${props.item.title}${props.updateCount ? `, ${props.updateCount} updates` : ''}`}>
        <Icon iconName="Comment" />{props.updateCount > 0 ? props.updateCount : ''}
      </button>
    </div>
  );
}

/* ---------- Add item row ---------- */

function AddRow(props: { cols: IBoardColumn[]; placeholder: string; indent?: boolean; onAdd: (title: string) => void; onDrop?: (e: React.DragEvent) => void }): JSX.Element {
  const [value, setValue] = React.useState('');
  const submit = (): void => {
    const t = value.trim();
    if (t) {
      props.onAdd(t);
      setValue('');
    }
  };
  return (
    <div style={{ display: 'contents' }} onDragOver={e => { if (props.onDrop) { e.preventDefault(); } }} onDrop={props.onDrop}>
      <div className={`${styles.cell} ${styles.edge}`} />
      <div className={styles.cell} />
      <div className={`${styles.cell} ${styles.addRow}`} style={{ gridColumn: `span ${props.cols.length + 2}`, paddingLeft: props.indent ? 32 : 0 }}>
        <input className={styles.addInput} placeholder={props.placeholder} value={value} onChange={e => setValue(e.target.value)}
          onKeyDown={e => { if (e.key === 'Enter') { submit(); } else if (e.key === 'Escape') { setValue(''); } }} onBlur={submit} aria-label={props.placeholder} />
      </div>
    </div>
  );
}

/* ---------- Summary row ---------- */

function SummaryRow(props: { cols: IBoardColumn[]; items: IWorkItem[] }): JSX.Element {
  return (
    <div style={{ display: 'contents' }}>
      <div className={`${styles.cell} ${styles.edge}`} style={{ borderBottom: 0 }} />
      <div className={styles.cell} style={{ borderLeft: 0, borderBottom: 0, background: 'transparent' }} />
      <div className={styles.cell} style={{ borderBottom: 0, background: 'transparent' }} />
      {props.cols.map(col => {
        const s = summarize(col, props.items);
        let content: React.ReactNode = null;
        if (s.kind === 'labels' && s.total > 0) {
          content = (
            <div className={styles.mixBar} style={{ width: '100%', height: 20, borderRadius: 3 }}
              title={s.parts.map(p => `${p.label.text}: ${p.count}`).join('\n')}>
              {s.parts.map(p => <span key={p.label.text} style={{ flex: p.count, background: p.label.color }} />)}
            </div>
          );
        } else if (s.kind === 'sum') {
          content = <strong style={{ color: 'var(--wb-text)' }}>{Math.round(s.value * 100) / 100}{col.unit ? ' ' + col.unit : ''}</strong>;
        } else if (s.kind === 'range' && (s.start || s.end)) {
          content = formatRange(s.start, s.end);
        } else if (s.kind === 'checked' && s.total > 0) {
          content = `${s.checked} / ${s.total}`;
        }
        return <div key={col.id} className={`${styles.cell} ${styles.summaryCell}`}>{content}</div>;
      })}
      <div className={`${styles.cell}`} style={{ background: 'transparent', borderBottom: 0, borderRight: 0 }} />
    </div>
  );
}

/* ---------- Table view ---------- */

export function TableView(props: ITableViewProps): JSX.Element {
  const { board, rights, view, actions } = props;
  const cols = board.config.columns.filter(c => view.hiddenColumnIds.indexOf(c.id) < 0);
  const template = gridTemplate(cols);
  const grouped = React.useMemo(() => visibleRows(board, props.items, view), [board, props.items, view]);
  const [expanded, setExpanded] = React.useState<number[]>([]);
  const [selected, setSelected] = React.useState<number[]>([]);
  const [hoverId, setHoverId] = React.useState<number | null>(null);
  const [dragId, setDragId] = React.useState<number | null>(null);
  const [dragOverId, setDragOverId] = React.useState<number | null>(null);
  const [menu, setMenu] = React.useState<{ target: HTMLElement; items: IContextualMenuItem[] } | null>(null);
  const [colorFor, setColorFor] = React.useState<{ target: HTMLElement; group: IBoardGroup } | null>(null);
  const [renameCol, setRenameCol] = React.useState<{ column: IBoardColumn; title: string } | null>(null);
  const [addingSubFor, setAddingSubFor] = React.useState<number | null>(null);
  const [confirmDialog, confirm] = useConfirm();
  const canDrag = rights.canEdit && !view.sort;

  const subsByParent = React.useMemo(() => {
    const m: { [id: number]: number } = {};
    props.items.forEach(i => {
      if (i.parentId) {
        m[i.parentId] = (m[i.parentId] || 0) + 1;
      }
    });
    return m;
  }, [props.items]);

  const toggleExpanded = (id: number): void => setExpanded(e => (e.indexOf(id) >= 0 ? e.filter(x => x !== id) : e.concat([id])));
  const toggleSelected = (id: number): void => setSelected(s => (s.indexOf(id) >= 0 ? s.filter(x => x !== id) : s.concat([id])));
  const toggleCollapsed = (groupId: string): void => {
    const c = view.collapsedGroupIds;
    props.setView({ ...view, collapsedGroupIds: c.indexOf(groupId) >= 0 ? c.filter(x => x !== groupId) : c.concat([groupId]) });
  };

  const groupMenu = (group: IBoardGroup, index: number, target: HTMLElement): void => {
    const items: IContextualMenuItem[] = [
      { key: 'color', text: 'Change colour', iconProps: { iconName: 'Color' }, onClick: () => setColorFor({ target, group }) },
      { key: 'up', text: 'Move up', iconProps: { iconName: 'Up' }, disabled: index === 0, onClick: () => { actions.moveGroup(group.id, -1).catch(() => undefined); } },
      { key: 'down', text: 'Move down', iconProps: { iconName: 'Down' }, disabled: index === board.config.groups.length - 1, onClick: () => { actions.moveGroup(group.id, 1).catch(() => undefined); } },
      { key: 'add', text: 'Add group', iconProps: { iconName: 'Add' }, onClick: () => { actions.addGroup().catch(() => undefined); } },
      {
        key: 'delete', text: 'Delete group', iconProps: { iconName: 'Delete' }, disabled: board.config.groups.length <= 1,
        onClick: () => {
          const count = props.items.filter(i => i.groupId === group.id && !i.parentId).length;
          confirm({
            title: `Delete "${group.title}"?`,
            message: count > 0 ? `Its ${count} item(s) and their subitems go to the site recycle bin, where a site owner can restore them.` : 'The group is empty.',
            confirmText: 'Delete group',
            danger: true
          }).then(ok => { if (ok) { actions.deleteGroup(group.id).catch(() => undefined); } }).catch(() => undefined);
        }
      }
    ];
    setMenu({ target, items });
  };

  const columnMenu = (col: IBoardColumn, target: HTMLElement): void => {
    const index = board.config.columns.map(c => c.id).indexOf(col.id);
    const items: IContextualMenuItem[] = [
      { key: 'asc', text: 'Sort ascending', iconProps: { iconName: 'SortUp' }, onClick: () => props.setView({ ...view, sort: { columnId: col.id, dir: 'asc' } }) },
      { key: 'desc', text: 'Sort descending', iconProps: { iconName: 'SortDown' }, onClick: () => props.setView({ ...view, sort: { columnId: col.id, dir: 'desc' } }) },
      { key: 'hide', text: 'Hide column', iconProps: { iconName: 'Hide3' }, onClick: () => props.setView({ ...view, hiddenColumnIds: view.hiddenColumnIds.concat([col.id]) }) }
    ];
    if (rights.canManage) {
      items.push(
        { key: 'div', itemType: 1 },
        { key: 'rename', text: 'Rename', iconProps: { iconName: 'Rename' }, onClick: () => setRenameCol({ column: col, title: col.title }) }
      );
      if (col.type === 'status' || col.type === 'dropdown') {
        items.push({ key: 'labels', text: col.type === 'status' ? 'Edit labels' : 'Edit options', iconProps: { iconName: 'Tag' }, onClick: () => props.onEditLabels(col) });
      }
      items.push(
        { key: 'left', text: 'Move left', iconProps: { iconName: 'ChevronLeft' }, disabled: index <= 0, onClick: () => { actions.moveColumn(col.id, index - 1).catch(() => undefined); } },
        { key: 'right', text: 'Move right', iconProps: { iconName: 'ChevronRight' }, disabled: index >= board.config.columns.length - 1, onClick: () => { actions.moveColumn(col.id, index + 1).catch(() => undefined); } },
        {
          key: 'delete', text: 'Delete column', iconProps: { iconName: 'Delete' },
          onClick: () => {
            confirm({
              title: `Delete the "${col.title}" column?`,
              message: col.field.indexOf('WB_c_') === 0
                ? 'The column and its values are removed from every item. This cannot be undone.'
                : 'The column is removed from the board. Its values stay in the SharePoint list, and adding the same kind of column again brings them back.',
              confirmText: 'Delete column',
              danger: true
            }).then(ok => { if (ok) { actions.deleteColumn(col.id).catch(() => undefined); } }).catch(() => undefined);
          }
        }
      );
    }
    setMenu({ target, items });
  };

  const addColumnMenu = (target: HTMLElement): void => {
    setMenu({
      target,
      items: (COLUMN_TYPES.map(t => ({
        key: t.type,
        text: t.title,
        title: t.description,
        iconProps: { iconName: t.icon },
        onClick: () => { actions.addColumn(t.type, t.title, cols.length ? cols[cols.length - 1].id : undefined).catch(() => undefined); }
      })) as IContextualMenuItem[]).concat([
        { key: 'roles-divider', itemType: 1 },
        {
          key: 'roles', text: 'Project roles', iconProps: { iconName: 'Contact' },
          subMenuProps: {
            items: ['Project lead', 'Project head', 'Project sponsor'].map(title => ({
              key: title,
              text: title,
              title: `A People column named ${title}`,
              disabled: board.config.columns.some(c => c.title.toLowerCase() === title.toLowerCase()),
              onClick: () => { actions.addColumn('people', title, cols.length ? cols[cols.length - 1].id : undefined).catch(() => undefined); }
            }))
          }
        }
      ])
    });
  };

  const itemMenu = (item: IWorkItem, target: HTMLElement): void => {
    const items: IContextualMenuItem[] = [
      { key: 'open', text: 'Open', iconProps: { iconName: 'OpenPane' }, onClick: () => props.onOpenItem(item.id) }
    ];
    if (rights.canAdd && !item.parentId) {
      items.push({
        key: 'sub', text: 'Add subitem', iconProps: { iconName: 'Add' },
        onClick: () => { setExpanded(e => (e.indexOf(item.id) >= 0 ? e : e.concat([item.id]))); setAddingSubFor(item.id); }
      });
    }
    if (rights.canEdit && !item.parentId) {
      items.push({
        key: 'move', text: 'Move to group', iconProps: { iconName: 'MoveToFolder' },
        subMenuProps: {
          items: board.config.groups.filter(g => g.id !== item.groupId).map(g => ({
            key: g.id, text: g.title, onClick: () => actions.moveItem(item.id, g.id, null)
          }))
        }
      });
    }
    if (rights.canDelete) {
      items.push({
        key: 'delete', text: 'Delete', iconProps: { iconName: 'Delete' },
        onClick: () => {
          confirm({
            title: `Delete "${item.title}"?`,
            message: 'The item and its subitems go to the site recycle bin, where they can be restored.',
            confirmText: 'Delete',
            danger: true
          }).then(ok => { if (ok) { actions.deleteItems([item.id]).catch(() => undefined); } }).catch(() => undefined);
        }
      });
    }
    setMenu({ target, items });
  };

  const onDropOnRow = (target: IWorkItem): ((e: React.DragEvent) => void) => (e: React.DragEvent): void => {
    e.preventDefault();
    const id = dragId;
    setDragId(null);
    setDragOverId(null);
    if (id === null || id === target.id) {
      return;
    }
    const dragged = props.items.filter(i => i.id === id)[0];
    if (!dragged) {
      return;
    }
    // Items move among items; subitems move among their siblings.
    if ((dragged.parentId || null) !== (target.parentId || null)) {
      return;
    }
    actions.moveItem(id, target.groupId, target.id);
  };

  const renderRow = (group: IBoardGroup, item: IWorkItem, isSub: boolean): JSX.Element => {
    const done = isDone(board, item);
    const rowClass = [
      selected.indexOf(item.id) >= 0 ? styles.rowSelected : hoverId === item.id ? styles.rowHover : '',
      dragOverId === item.id ? styles.rowDragOver : ''
    ].join(' ');
    return (
      <div key={item.id} style={{ display: 'contents' }} className={rowClass}
        onMouseEnter={() => setHoverId(item.id)} onMouseLeave={() => setHoverId(h => (h === item.id ? null : h))}
        onDragOver={e => { if (dragId !== null) { e.preventDefault(); setDragOverId(item.id); } }}
        onDragLeave={() => setDragOverId(d => (d === item.id ? null : d))}
        onDrop={onDropOnRow(item)}>
        <div className={`${styles.cell} ${styles.edge}`} style={{ background: group.color, opacity: isSub ? 0.5 : 1 }} />
        <div className={`${styles.cell} ${styles.checkCell}`}>
          <Checkbox checked={selected.indexOf(item.id) >= 0} onChange={() => toggleSelected(item.id)} ariaLabel={`Select ${item.title}`} />
        </div>
        <TitleCell
          item={item} board={board} canEdit={rights.canEdit} isSub={isSub}
          subCount={subsByParent[item.id] || 0} expanded={expanded.indexOf(item.id) >= 0}
          updateCount={props.updateCounts[item.id] || 0}
          onToggle={() => toggleExpanded(item.id)} onOpen={() => props.onOpenItem(item.id)}
          onRename={t => actions.renameItem(item, t)} onMenu={t => itemMenu(item, t)}
          draggable={canDrag} onDragStart={e => { e.dataTransfer.setData('text/plain', String(item.id)); e.dataTransfer.effectAllowed = 'move'; setDragId(item.id); }}
        />
        {cols.map(col => (
          <div key={col.id} className={styles.cell} style={col.type === 'status' ? { borderRight: '1px solid var(--wb-bg)' } : undefined}>
            <Cell column={col} value={item.values[col.id]} canEdit={rights.canEdit} canManage={rights.canManage} accent={group.color}
              done={done} onChange={v => actions.updateCell(item, col.id, v)} onEditLabels={props.onEditLabels} onAddOption={actions.addOption} />
          </div>
        ))}
        <div className={styles.cell} />
      </div>
    );
  };

  const renderGroup = (group: IBoardGroup, index: number): JSX.Element => {
    const rows: IGroupRows[] = grouped[group.id] || [];
    const collapsed = view.collapsedGroupIds.indexOf(group.id) >= 0;
    const groupItemIds = rows.map(r => r.item.id);
    const allSelected = groupItemIds.length > 0 && groupItemIds.every(id => selected.indexOf(id) >= 0);
    return (
      <section key={group.id} className={styles.group} aria-label={group.title}>
        <div className={styles.groupHeader}>
          <button type="button" className={styles.iconBtn} onClick={() => toggleCollapsed(group.id)} aria-expanded={!collapsed}
            aria-label={collapsed ? `Expand ${group.title}` : `Collapse ${group.title}`} style={{ color: group.color }}>
            <Icon iconName={collapsed ? 'ChevronRight' : 'ChevronDown'} />
          </button>
          <GroupTitle group={group} canEdit={rights.canManage} onRename={t => { actions.updateGroup(group.id, { title: t }).catch(() => undefined); }} />
          <span className={styles.muted + ' ' + styles.small}>{rows.length} {rows.length === 1 ? 'item' : 'items'}</span>
          {rights.canManage && (
            <button type="button" className={styles.iconBtn} onClick={e => groupMenu(group, index, e.currentTarget)} aria-label={`Actions for ${group.title}`}>
              <Icon iconName="More" />
            </button>
          )}
        </div>
        {!collapsed && (
          <div className={styles.grid} style={{ gridTemplateColumns: template }} role="table" aria-label={`${group.title} items`}>
            <div className={`${styles.hcell} ${styles.edge}`} style={{ background: group.color, borderTopLeftRadius: 4 }} />
            <div className={`${styles.hcell} ${styles.checkCell}`}>
              <Checkbox checked={allSelected} onChange={() => setSelected(s => (allSelected ? s.filter(id => groupItemIds.indexOf(id) < 0) : s.concat(groupItemIds.filter(id => s.indexOf(id) < 0))))}
                ariaLabel={`Select all in ${group.title}`} />
            </div>
            <div className={`${styles.hcell} ${styles.hcellTitle}`}>
              <span className={styles.colName} onClick={() => props.setView({ ...view, sort: view.sort && view.sort.columnId === 'title' && view.sort.dir === 'asc' ? { columnId: 'title', dir: 'desc' } : { columnId: 'title', dir: 'asc' } })}>
                Item {view.sort && view.sort.columnId === 'title' && <Icon iconName={view.sort.dir === 'asc' ? 'SortUp' : 'SortDown'} />}
              </span>
            </div>
            {cols.map(col => (
              <div key={col.id} className={styles.hcell}>
                <span className={styles.colName} title={col.title}>{col.title}</span>
                {view.sort && view.sort.columnId === col.id && <Icon iconName={view.sort.dir === 'asc' ? 'SortUp' : 'SortDown'} />}
                <button type="button" className={`${styles.iconBtn} ${styles.colMenuBtn}`} onClick={e => columnMenu(col, e.currentTarget)} aria-label={`Column options for ${col.title}`}>
                  <Icon iconName="ChevronDown" style={{ fontSize: 10 }} />
                </button>
              </div>
            ))}
            <div className={styles.hcell}>
              {rights.canManage && (
                <button type="button" className={styles.iconBtn} onClick={e => addColumnMenu(e.currentTarget)} aria-label="Add column" title="Add column">
                  <Icon iconName="Add" />
                </button>
              )}
            </div>
            {rows.map(r => {
              const out = [renderRow(group, r.item, false)];
              if (expanded.indexOf(r.item.id) >= 0 || (r.subitems.length > 0 && (view.search || view.personId !== null) && subsByParent[r.item.id])) {
                r.subitems.forEach(s => out.push(renderRow(group, s, true)));
                if (rights.canAdd) {
                  out.push(
                    <AddSubRow key={'addsub' + r.item.id} cols={cols} autoFocus={addingSubFor === r.item.id}
                      onAdd={t => { setAddingSubFor(null); actions.addItem(group.id, t, { parentId: r.item.id }).catch(() => undefined); }} />
                  );
                }
              }
              return out;
            })}
            {rights.canAdd && (
              <AddRow cols={cols} placeholder="+ Add item" onAdd={t => { actions.addItem(group.id, t).catch(() => undefined); }}
                onDrop={e => { e.preventDefault(); if (dragId !== null) { actions.moveItem(dragId, group.id, null); } setDragId(null); setDragOverId(null); }} />
            )}
            <SummaryRow cols={cols} items={rows.map(r => r.item)} />
          </div>
        )}
      </section>
    );
  };

  const hasAny = Object.keys(grouped).length > 0;
  const filtering = !!(view.search || view.personId !== null || view.dateFilter || Object.keys(view.labelFilters).some(k => view.labelFilters[k].length > 0));

  return (
    <div className={styles.boardBody}>
      {!hasAny && filtering && <div className={styles.center}><span className={styles.muted}>No items match your search or filters.</span></div>}
      {board.config.groups.map((g, i) => (filtering && !grouped[g.id] ? null : renderGroup(g, i)))}
      {rights.canManage && !filtering && (
        <div className={styles.addGroupBtn}>
          <DefaultButton iconProps={{ iconName: 'Add' }} text="Add new group" onClick={() => { actions.addGroup().catch(() => undefined); }} />
        </div>
      )}

      {selected.length > 0 && (
        <div className={styles.bulkBar} role="region" aria-label="Selected items">
          <strong>{selected.length} selected</strong>
          {rights.canEdit && (
            <DefaultButton text="Move to group" iconProps={{ iconName: 'MoveToFolder' }}
              menuProps={{ items: board.config.groups.map(g => ({ key: g.id, text: g.title, onClick: () => { actions.moveItemsToGroup(selected, g.id).catch(() => undefined); setSelected([]); } })) }} />
          )}
          {rights.canDelete && (
            <DefaultButton text="Delete" iconProps={{ iconName: 'Delete' }} onClick={() => {
              confirm({ title: `Delete ${selected.length} item(s)?`, message: 'They go to the site recycle bin, with their subitems.', confirmText: 'Delete', danger: true })
                .then(ok => { if (ok) { actions.deleteItems(selected).catch(() => undefined); setSelected([]); } }).catch(() => undefined);
            }} />
          )}
          <button type="button" className={styles.iconBtn} onClick={() => setSelected([])} aria-label="Clear selection"><Icon iconName="Cancel" /></button>
        </div>
      )}

      {menu && (
        <ContextualMenu target={menu.target} items={menu.items} onDismiss={() => setMenu(null)} directionalHint={DirectionalHint.bottomLeftEdge} />
      )}
      {colorFor && (
        <Callout target={colorFor.target} onDismiss={() => setColorFor(null)} setInitialFocus>
          <div className={styles.popup}>
            <strong>Group colour</strong>
            <Swatches value={colorFor.group.color} onChange={c => { actions.updateGroup(colorFor.group.id, { color: c }).catch(() => undefined); setColorFor(null); }} />
          </div>
        </Callout>
      )}
      {renameCol && (
        <Dialog hidden={false} onDismiss={() => setRenameCol(null)} dialogContentProps={{ title: 'Rename column' }}>
          <TextField label="Column name" value={renameCol.title} autoFocus onChange={(_, v) => setRenameCol({ ...renameCol, title: v || '' })}
            onKeyDown={e => { if (e.key === 'Enter' && renameCol.title.trim()) { actions.renameColumn(renameCol.column.id, renameCol.title.trim()).catch(() => undefined); setRenameCol(null); } }} />
          <DialogFooter>
            <PrimaryButton text="Save" disabled={!renameCol.title.trim()} onClick={() => { actions.renameColumn(renameCol.column.id, renameCol.title.trim()).catch(() => undefined); setRenameCol(null); }} />
            <DefaultButton text="Cancel" onClick={() => setRenameCol(null)} />
          </DialogFooter>
        </Dialog>
      )}
      {confirmDialog}
    </div>
  );
}

function AddSubRow(props: { cols: IBoardColumn[]; autoFocus: boolean; onAdd: (title: string) => void }): JSX.Element {
  const [value, setValue] = React.useState('');
  const submit = (): void => {
    const t = value.trim();
    if (t) {
      props.onAdd(t);
      setValue('');
    }
  };
  return (
    <div style={{ display: 'contents' }}>
      <div className={`${styles.cell} ${styles.edge}`} />
      <div className={styles.cell} />
      <div className={`${styles.cell} ${styles.addRow}`} style={{ gridColumn: `span ${props.cols.length + 2}`, paddingLeft: 32 }}>
        <input className={styles.addInput} placeholder="+ Add subitem" value={value} autoFocus={props.autoFocus} onChange={e => setValue(e.target.value)}
          onKeyDown={e => { if (e.key === 'Enter') { submit(); } else if (e.key === 'Escape') { setValue(''); } }} onBlur={submit} aria-label="Add subitem" />
      </div>
    </div>
  );
}

function GroupTitle(props: { group: IBoardGroup; canEdit: boolean; onRename: (t: string) => void }): JSX.Element {
  const [draft, setDraft] = React.useState(props.group.title);
  React.useEffect(() => setDraft(props.group.title), [props.group.title]);
  const commit = (): void => {
    const t = draft.trim();
    if (t && t !== props.group.title) {
      props.onRename(t);
    } else {
      setDraft(props.group.title);
    }
  };
  return (
    <input className={styles.groupTitle} style={{ color: props.group.color, width: Math.max(80, draft.length * 9 + 16) }} value={draft}
      disabled={!props.canEdit} onChange={e => setDraft(e.target.value)} onBlur={commit}
      onKeyDown={e => { if (e.key === 'Enter') { (e.target as HTMLInputElement).blur(); } else if (e.key === 'Escape') { setDraft(props.group.title); } }}
      aria-label="Group name" />
  );
}
