import * as React from 'react';
import { Dropdown, DefaultButton, Icon } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { IBoard, IBoardColumn, IWorkItem, ITimelineValue, IBoardGroup } from '../../models/types';
import { IBoardRights } from '../../services/permissions';
import { IBoardActions } from '../board/useBoard';
import { IViewState, visibleRows, labelFor } from '../../engine/viewQuery';
import { addDays, diffDays, todayIso, isoToDate, MONTH_NAMES, formatRange } from '../../engine/dates';
import { textOn } from '../../models/colors';

export interface ITimelineViewProps {
  board: IBoard;
  items: IWorkItem[];
  rights: IBoardRights;
  view: IViewState;
  actions: IBoardActions;
  onOpenItem: (itemId: number) => void;
}

type Scale = 'day' | 'week' | 'month';
const PX: { [s: string]: number } = { day: 40, week: 18, month: 5 };
const LABEL_W = 260;

interface IDrag {
  itemId: number;
  mode: 'move' | 'start' | 'end';
  startX: number;
  delta: number;
}

function range(col: IBoardColumn, item: IWorkItem): { start: string | null; end: string | null } {
  const v = item.values[col.id];
  if (col.type === 'timeline') {
    const tv = (v as ITimelineValue) || { start: null, end: null };
    return { start: tv.start || tv.end, end: tv.end || tv.start };
  }
  return { start: (v as string) || null, end: (v as string) || null };
}

export function TimelineView(props: ITimelineViewProps): JSX.Element {
  const { board, rights, actions } = props;
  const dateCols = board.config.columns.filter(c => c.type === 'timeline' || c.type === 'date');
  const col = dateCols.filter(c => c.id === board.config.timelineColumnId)[0] || dateCols[0];
  const [scale, setScale] = React.useState<Scale>('week');
  const [drag, setDrag] = React.useState<IDrag | null>(null);
  const dragRef = React.useRef<IDrag | null>(null);
  dragRef.current = drag;
  const scrollRef = React.useRef<HTMLDivElement>(null);
  const statusCol = board.config.columns.filter(c => c.id === board.config.kanbanColumnId)[0] || board.config.columns.filter(c => c.type === 'status')[0];

  const grouped = React.useMemo(() => visibleRows(board, props.items, props.view), [board, props.items, props.view]);
  const today = todayIso();

  // Visible date range: a little before the earliest start to a little after the latest end.
  const bounds = React.useMemo(() => {
    let min = addDays(today, -14);
    let max = addDays(today, 42);
    if (col) {
      props.items.forEach(i => {
        const r = range(col, i);
        if (r.start && r.start < min) {
          min = r.start;
        }
        if (r.end && r.end > max) {
          max = r.end;
        }
      });
    }
    // Start on a Monday so week ticks line up.
    const d = isoToDate(addDays(min, -7));
    const dow = (d.getDay() + 6) % 7;
    const start = addDays(addDays(min, -7), -dow);
    return { start, days: diffDays(start, addDays(max, 21)) + 1 };
  }, [props.items, col]);

  const px = PX[scale];
  const width = bounds.days * px;
  const x = (iso: string): number => diffDays(bounds.start, iso) * px;

  // Scroll so today is in view on first render and when the scale changes.
  React.useEffect(() => {
    if (scrollRef.current) {
      scrollRef.current.scrollLeft = Math.max(0, x(today) - 3 * px * 7 / (scale === 'month' ? 1 : 3));
    }
  }, [scale, bounds.start]);

  React.useEffect(() => {
    if (!drag) {
      return undefined;
    }
    const move = (e: PointerEvent): void => {
      const d = dragRef.current;
      if (d) {
        setDrag({ ...d, delta: Math.round((e.clientX - d.startX) / px) });
      }
    };
    const up = (): void => {
      const d = dragRef.current;
      setDrag(null);
      if (!d || d.delta === 0 || !col) {
        return;
      }
      const item = props.items.filter(i => i.id === d.itemId)[0];
      if (!item) {
        return;
      }
      const r = range(col, item);
      if (!r.start || !r.end) {
        return;
      }
      let start = r.start;
      let end = r.end;
      if (d.mode === 'move') {
        start = addDays(start, d.delta);
        end = addDays(end, d.delta);
      } else if (d.mode === 'start') {
        start = addDays(start, d.delta);
        if (start > end) {
          start = end;
        }
      } else {
        end = addDays(end, d.delta);
        if (end < start) {
          end = start;
        }
      }
      actions.updateCell(item, col.id, col.type === 'timeline' ? { start, end } : start);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
    return () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
    };
  }, [drag !== null, px, col]);

  if (!col) {
    return (
      <div className={styles.center}>
        <span className={styles.muted}>The Timeline view shows a Timeline or Date column. This board has neither yet.</span>
        {rights.canManage && <DefaultButton text="Add a Timeline column" iconProps={{ iconName: 'Timeline' }} onClick={() => { actions.addColumn('timeline', 'Timeline').catch(() => undefined); }} />}
      </div>
    );
  }

  // Header ticks.
  const ticks: { left: number; label: string }[] = [];
  const weekends: number[] = [];
  for (let i = 0; i < bounds.days; i++) {
    const iso = addDays(bounds.start, i);
    const d = isoToDate(iso);
    if (scale !== 'month' && (d.getDay() === 0 || d.getDay() === 6)) {
      weekends.push(i * px);
    }
    if (scale === 'day') {
      ticks.push({ left: i * px, label: (d.getDate() === 1 || i === 0 ? MONTH_NAMES[d.getMonth()] + ' ' : '') + d.getDate() });
    } else if (scale === 'week' && d.getDay() === 1) {
      ticks.push({ left: i * px, label: d.getDate() + ' ' + MONTH_NAMES[d.getMonth()] });
    } else if (scale === 'month' && d.getDate() === 1) {
      ticks.push({ left: i * px, label: MONTH_NAMES[d.getMonth()] + ' ' + d.getFullYear() });
    }
  }

  const barFor = (item: IWorkItem, group: IBoardGroup): JSX.Element | null => {
    const r = range(col, item);
    if (!r.start || !r.end) {
      return null;
    }
    let start = r.start;
    let end = r.end;
    if (drag && drag.itemId === item.id) {
      if (drag.mode === 'move') {
        start = addDays(start, drag.delta);
        end = addDays(end, drag.delta);
      } else if (drag.mode === 'start') {
        start = addDays(start, drag.delta);
      } else {
        end = addDays(end, drag.delta);
      }
      if (end < start) {
        end = start;
      }
    }
    const label = statusCol ? labelFor(statusCol, item.values[statusCol.id] as string) : null;
    const color = label ? label.color : group.color;
    const left = x(start);
    const w = (diffDays(start, end) + 1) * px;
    const begin = (mode: IDrag['mode']): ((e: React.PointerEvent) => void) => (e: React.PointerEvent): void => {
      if (!rights.canEdit) {
        return;
      }
      e.stopPropagation();
      e.preventDefault();
      setDrag({ itemId: item.id, mode, startX: e.clientX, delta: 0 });
    };
    const title = `${item.title}: ${formatRange(start, end)}${label ? ' · ' + label.text : ''}`;
    if (start === end && px < 20) {
      return (
        <div className={styles.ganttMilestone} style={{ left: left + px / 2 - 9, background: color }} title={title}
          onPointerDown={begin('move')} onDoubleClick={() => props.onOpenItem(item.id)} role="img" aria-label={title} />
      );
    }
    return (
      <div className={styles.ganttBar} style={{ left, width: Math.max(w, 8), background: color, color: textOn(color) }} title={title}
        onPointerDown={begin('move')} onDoubleClick={() => props.onOpenItem(item.id)} role="img" aria-label={title}>
        {w > 60 ? item.title : ''}
        {rights.canEdit && <span className={styles.ganttHandle} style={{ left: 0 }} onPointerDown={begin('start')} />}
        {rights.canEdit && <span className={styles.ganttHandle} style={{ right: 0 }} onPointerDown={begin('end')} />}
      </div>
    );
  };

  const schedule = (item: IWorkItem): ((e: React.MouseEvent<HTMLDivElement>) => void) => (e: React.MouseEvent<HTMLDivElement>): void => {
    if (!rights.canEdit || range(col, item).start) {
      return;
    }
    const rect = e.currentTarget.getBoundingClientRect();
    const day = addDays(bounds.start, Math.floor((e.clientX - rect.left) / px));
    actions.updateCell(item, col.id, col.type === 'timeline' ? { start: day, end: addDays(day, scale === 'day' ? 0 : 4) } : day);
  };

  const gridCols = `${LABEL_W}px ${width}px`;
  const todayLeft = x(today) + px / 2;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0 }}>
      <div className={styles.toolbar} style={{ paddingTop: 0 }}>
        {dateCols.length > 1 && (
          <Dropdown styles={{ root: { width: 200 } }} ariaLabel="Column shown" selectedKey={col.id}
            options={dateCols.map(c => ({ key: c.id, text: c.title }))} disabled={!rights.canManage}
            onChange={(_, o) => { if (o) { actions.updateConfig(cfg => ({ ...cfg, timelineColumnId: String(o.key) })).catch(() => undefined); } }} />
        )}
        {(['day', 'week', 'month'] as Scale[]).map(s => (
          <DefaultButton key={s} text={s === 'day' ? 'Days' : s === 'week' ? 'Weeks' : 'Months'} checked={scale === s} onClick={() => setScale(s)}
            styles={{ rootChecked: { background: 'var(--wb-selected)' } }} />
        ))}
        <DefaultButton text="Today" iconProps={{ iconName: 'GotoToday' }} onClick={() => { if (scrollRef.current) { scrollRef.current.scrollLeft = Math.max(0, x(today) - 200); } }} />
        <span className={`${styles.small} ${styles.muted}`}>
          {rights.canEdit ? 'Drag a bar to move it, drag its ends to change dates. Click an empty row to schedule an item.' : ''}
        </span>
      </div>
      <div className={styles.gantt} ref={scrollRef}>
        <div className={styles.ganttGrid} style={{ gridTemplateColumns: gridCols }}>
          <div className={`${styles.ganttLabel} ${styles.ganttGroupLabel} ${styles.ganttHeader}`} style={{ zIndex: 4 }}>Item</div>
          <div className={`${styles.ganttTrack} ${styles.ganttHeader}`}>
            {ticks.map(t => <div key={t.left} className={styles.ganttTick} style={{ left: t.left }}>{t.label}</div>)}
          </div>
          {board.config.groups.map(group => {
            const rows = grouped[group.id] || [];
            if (rows.length === 0) {
              return null;
            }
            return (
              <React.Fragment key={group.id}>
                <div className={`${styles.ganttLabel} ${styles.ganttGroupLabel}`} style={{ color: group.color }}>
                  <Icon iconName="ChevronDown" style={{ fontSize: 10 }} /> {group.title}
                </div>
                <div className={styles.ganttTrack} style={{ background: 'var(--wb-surface)' }} />
                {rows.map(({ item }) => (
                  <React.Fragment key={item.id}>
                    <div className={styles.ganttLabel} onClick={() => props.onOpenItem(item.id)} role="button" tabIndex={0}
                      onKeyDown={e => { if (e.key === 'Enter') { props.onOpenItem(item.id); } }}>
                      <span className={styles.boardDot} style={{ background: group.color }} />
                      <span className={styles.ellipsis}>{item.title || 'Untitled'}</span>
                    </div>
                    <div className={styles.ganttTrack} onClick={schedule(item)} style={{ cursor: rights.canEdit && !range(col, item).start ? 'copy' : 'default' }}>
                      {weekends.map(w => <div key={w} className={styles.ganttWeekend} style={{ left: w, width: px }} />)}
                      {barFor(item, group)}
                    </div>
                  </React.Fragment>
                ))}
              </React.Fragment>
            );
          })}
          <div className={styles.ganttToday} style={{ left: LABEL_W + todayLeft }} aria-hidden="true" />
        </div>
      </div>
    </div>
  );
}
