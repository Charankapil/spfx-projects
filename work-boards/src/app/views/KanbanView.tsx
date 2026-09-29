import * as React from 'react';
import { Icon, Dropdown } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { IBoard, IBoardColumn, IWorkItem, IPerson, ILabel } from '../../models/types';
import { IBoardRights } from '../../services/permissions';
import { IBoardActions } from '../board/useBoard';
import { IViewState, visibleRows, dueOf, isDone, labelFor } from '../../engine/viewQuery';
import { Avatars, LabelChip } from '../common/Common';
import { textOn } from '../../models/colors';
import { F } from '../../engine/fieldMap';
import { relativeDay, todayIso } from '../../engine/dates';

export interface IKanbanViewProps {
  board: IBoard;
  items: IWorkItem[];
  rights: IBoardRights;
  view: IViewState;
  actions: IBoardActions;
  updateCounts: { [itemId: number]: number };
  onOpenItem: (itemId: number) => void;
}

const EMPTY_LANE: ILabel = { id: '__empty', text: '', color: '#c4c4c4' };

export function KanbanView(props: IKanbanViewProps): JSX.Element {
  const { board, rights, actions } = props;
  const statusCols = board.config.columns.filter(c => c.type === 'status');
  const laneCol: IBoardColumn | undefined = statusCols.filter(c => c.id === board.config.kanbanColumnId)[0] || statusCols[0];
  const [dragId, setDragId] = React.useState<number | null>(null);
  const [overLane, setOverLane] = React.useState<string | null>(null);
  const [adding, setAdding] = React.useState<{ lane: string; title: string } | null>(null);

  const grouped = React.useMemo(() => visibleRows(board, props.items, props.view), [board, props.items, props.view]);
  const topItems: { item: IWorkItem; subs: IWorkItem[] }[] = [];
  board.config.groups.forEach(g => (grouped[g.id] || []).forEach(r => topItems.push({ item: r.item, subs: props.items.filter(i => i.parentId === r.item.id) })));

  if (!laneCol) {
    return (
      <div className={styles.center}>
        <span className={styles.muted}>Kanban lanes come from a Status column. Add a Status column to this board to use Kanban.</span>
        {rights.canManage && (
          <button type="button" className={styles.linkBtn} onClick={() => { actions.addColumn('status', 'Status').catch(() => undefined); }}>Add a Status column</button>
        )}
      </div>
    );
  }

  const lanes: ILabel[] = (laneCol.labels || []).concat([EMPTY_LANE]);
  const peopleCol = board.config.columns.filter(c => c.field === F.Owner)[0] || board.config.columns.filter(c => c.type === 'people')[0];
  const otherStatus = statusCols.filter(c => c.id !== laneCol.id);
  const today = todayIso();

  const drop = (lane: ILabel): ((e: React.DragEvent) => void) => (e: React.DragEvent): void => {
    e.preventDefault();
    setOverLane(null);
    const id = dragId;
    setDragId(null);
    if (id === null) {
      return;
    }
    const item = props.items.filter(i => i.id === id)[0];
    const newValue = lane.id === EMPTY_LANE.id ? null : lane.text;
    if (item && (item.values[laneCol.id] || null) !== newValue) {
      actions.updateCell(item, laneCol.id, newValue);
    }
  };

  const addCard = (lane: ILabel): void => {
    if (!adding || !adding.title.trim()) {
      setAdding(null);
      return;
    }
    const firstGroup = board.config.groups[0];
    actions.addItem(firstGroup.id, adding.title.trim(), { values: lane.id === EMPTY_LANE.id ? {} : { [laneCol.id]: lane.text } }).catch(() => undefined);
    setAdding(null);
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0 }}>
      {statusCols.length > 1 && (
        <div className={styles.toolbar} style={{ paddingTop: 0 }}>
          <Dropdown
            label="Lanes by"
            styles={{ root: { width: 220 } }}
            selectedKey={laneCol.id}
            options={statusCols.map(c => ({ key: c.id, text: c.title }))}
            disabled={!rights.canManage}
            onChange={(_, o) => { if (o) { actions.updateConfig(cfg => ({ ...cfg, kanbanColumnId: String(o.key) })).catch(() => undefined); } }}
          />
        </div>
      )}
      <div className={styles.kanban}>
        {lanes.map(lane => {
          const cards = topItems.filter(t => (lane.id === EMPTY_LANE.id ? !t.item.values[laneCol.id] || !(laneCol.labels || []).some(l => l.text === t.item.values[laneCol.id]) : t.item.values[laneCol.id] === lane.text));
          if (lane.id === EMPTY_LANE.id && cards.length === 0 && dragId === null) {
            return null;
          }
          const headBg = lane.id === EMPTY_LANE.id ? 'var(--wb-border-light)' : lane.color;
          return (
            <div
              key={lane.id}
              className={`${styles.lane} ${overLane === lane.id ? styles.laneOver : ''}`}
              onDragOver={e => { if (dragId !== null) { e.preventDefault(); setOverLane(lane.id); } }}
              onDragLeave={() => setOverLane(o => (o === lane.id ? null : o))}
              onDrop={drop(lane)}
              aria-label={`${lane.text || 'No ' + laneCol.title}, ${cards.length} items`}
              role="region"
            >
              <div className={styles.laneHead} style={{ background: headBg, color: lane.id === EMPTY_LANE.id ? 'var(--wb-text)' : textOn(lane.color) }}>
                <span>{lane.text || `No ${laneCol.title.toLowerCase()}`}</span>
                <span>{cards.length}</span>
              </div>
              <div className={styles.laneBody}>
                {cards.map(({ item, subs }) => {
                  const due = dueOf(board, item);
                  const done = isDone(board, item);
                  const subsDone = subs.filter(s => isDone(board, s)).length;
                  const updates = props.updateCounts[item.id] || 0;
                  return (
                    <button
                      key={item.id}
                      type="button"
                      className={`${styles.kcard} ${dragId === item.id ? styles.kcardDragging : ''}`}
                      draggable={rights.canEdit}
                      onDragStart={e => { e.dataTransfer.setData('text/plain', String(item.id)); e.dataTransfer.effectAllowed = 'move'; setDragId(item.id); }}
                      onDragEnd={() => { setDragId(null); setOverLane(null); }}
                      onClick={() => props.onOpenItem(item.id)}
                    >
                      <strong>{item.title || 'Untitled'}</strong>
                      <div className={styles.row}>
                        {otherStatus.map(c => <LabelChip key={c.id} label={labelFor(c, item.values[c.id] as string)} />)}
                      </div>
                      {subs.length > 0 && (
                        <div>
                          <div className={styles.progress}><span style={{ width: `${Math.round((subsDone / subs.length) * 100)}%` }} /></div>
                          <span className={`${styles.small} ${styles.muted}`}>{subsDone} of {subs.length} subitems done</span>
                        </div>
                      )}
                      <div className={styles.row}>
                        {peopleCol && <Avatars people={(item.values[peopleCol.id] as IPerson[]) || []} size={22} />}
                        <span className={styles.spacer} />
                        {updates > 0 && <span className={`${styles.small} ${styles.muted}`}><Icon iconName="Comment" /> {updates}</span>}
                        {due && <span className={`${styles.small} ${!done && due < today ? styles.overdue : styles.muted}`}>{relativeDay(due)}</span>}
                      </div>
                    </button>
                  );
                })}
                {rights.canAdd && lane.id !== EMPTY_LANE.id && (
                  adding && adding.lane === lane.id ? (
                    <input className={styles.inlineInput} style={{ height: 34, borderRadius: 6 }} autoFocus placeholder="Item name"
                      value={adding.title} onChange={e => setAdding({ lane: lane.id, title: e.target.value })}
                      onBlur={() => addCard(lane)} onKeyDown={e => { if (e.key === 'Enter') { addCard(lane); } else if (e.key === 'Escape') { setAdding(null); } }}
                      aria-label={`New item in ${lane.text}`} />
                  ) : (
                    <button type="button" className={styles.linkBtn} onClick={() => setAdding({ lane: lane.id, title: '' })}>+ Add item</button>
                  )
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}
