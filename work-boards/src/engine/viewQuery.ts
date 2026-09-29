import { IBoard, IBoardColumn, IWorkItem, IPerson, ITimelineValue, ILabel } from '../models/types';
import { cellText } from './fieldMap';
import { todayIso, addDays } from './dates';

/** What the toolbar controls: search, person filter, label filters, sort and hidden columns. */
export interface IViewState {
  search: string;
  personId: number | null;
  /** columnId -> label texts to keep ("" means empty). */
  labelFilters: { [columnId: string]: string[] };
  dateFilter: '' | 'overdue' | 'thisweek' | 'nodate';
  sort: { columnId: string; dir: 'asc' | 'desc' } | null;
  hiddenColumnIds: string[];
  collapsedGroupIds: string[];
}

export const EMPTY_VIEW: IViewState = {
  search: '',
  personId: null,
  labelFilters: {},
  dateFilter: '',
  sort: null,
  hiddenColumnIds: [],
  collapsedGroupIds: []
};

export function activeFilterCount(v: IViewState): number {
  let n = 0;
  Object.keys(v.labelFilters).forEach(k => {
    if (v.labelFilters[k].length > 0) {
      n++;
    }
  });
  if (v.dateFilter) {
    n++;
  }
  return n;
}

/** The main date for an item: a Date column, or the end of a Timeline column. */
export function dueOf(board: IBoard, item: IWorkItem): string | null {
  const col = board.config.columns.filter(c => c.id === board.config.timelineColumnId)[0]
    || board.config.columns.filter(c => c.type === 'date' || c.type === 'timeline')[0];
  if (!col) {
    return null;
  }
  const v = item.values[col.id];
  if (col.type === 'timeline') {
    return v ? (v as ITimelineValue).end || (v as ITimelineValue).start : null;
  }
  return (v as string) || null;
}

export function isDone(board: IBoard, item: IWorkItem): boolean {
  const col = board.config.columns.filter(c => c.id === board.config.kanbanColumnId)[0]
    || board.config.columns.filter(c => c.type === 'status')[0];
  if (!col) {
    return false;
  }
  const label = labelFor(col, item.values[col.id] as string);
  return !!(label && label.isDone);
}

export function labelFor(col: IBoardColumn, text: string | null | undefined): ILabel | null {
  if (!text) {
    return null;
  }
  return (col.labels || []).filter(l => l.text === text)[0] || { id: '', text, color: '#c4c4c4' };
}

function matches(board: IBoard, item: IWorkItem, v: IViewState, today: string): boolean {
  const cols = board.config.columns;
  if (v.search) {
    const q = v.search.toLowerCase();
    const hay = [item.title, board.key + '-' + item.id].concat(cols.map(c => cellText(c, item.values[c.id]))).join(' ').toLowerCase();
    if (hay.indexOf(q) < 0) {
      return false;
    }
  }
  if (v.personId !== null) {
    const hasPerson = cols.some(c => c.type === 'people' && ((item.values[c.id] as IPerson[]) || []).some(p => p.id === v.personId));
    if (!hasPerson) {
      return false;
    }
  }
  for (const colId of Object.keys(v.labelFilters)) {
    const wanted = v.labelFilters[colId];
    if (!wanted || wanted.length === 0) {
      continue;
    }
    const col = cols.filter(c => c.id === colId)[0];
    if (!col) {
      continue;
    }
    const raw = item.values[col.id];
    const values: string[] = col.type === 'dropdown' ? ((raw as string[]) || []) : raw ? [String(raw)] : [];
    const ok = values.length === 0 ? wanted.indexOf('') >= 0 : values.some(x => wanted.indexOf(x) >= 0);
    if (!ok) {
      return false;
    }
  }
  if (v.dateFilter) {
    const due = dueOf(board, item);
    if (v.dateFilter === 'nodate' && due) {
      return false;
    }
    if (v.dateFilter === 'overdue' && (!due || due >= today || isDone(board, item))) {
      return false;
    }
    if (v.dateFilter === 'thisweek' && (!due || due < today || due > addDays(today, 7))) {
      return false;
    }
  }
  return true;
}

function sortValue(col: IBoardColumn, item: IWorkItem): string | number {
  const v = item.values[col.id];
  if (col.type === 'number') {
    return typeof v === 'number' ? v : Number.NEGATIVE_INFINITY;
  }
  if (col.type === 'status') {
    const idx = (col.labels || []).map(l => l.text).indexOf(v as string);
    return idx < 0 ? 999 : idx;
  }
  if (col.type === 'timeline') {
    const tv = v as ITimelineValue;
    return (tv && (tv.start || tv.end)) || '9999';
  }
  if (col.type === 'date') {
    return (v as string) || '9999';
  }
  return cellText(col, v).toLowerCase();
}

export function compareItems(board: IBoard, v: IViewState): (a: IWorkItem, b: IWorkItem) => number {
  const col = v.sort ? board.config.columns.filter(c => c.id === (v.sort as { columnId: string }).columnId)[0] : null;
  const byTitle = v.sort && v.sort.columnId === 'title';
  const dir = v.sort && v.sort.dir === 'desc' ? -1 : 1;
  return (a, b) => {
    if (byTitle) {
      return dir * a.title.localeCompare(b.title);
    }
    if (col) {
      const x = sortValue(col, a);
      const y = sortValue(col, b);
      if (x < y) {
        return -dir;
      }
      if (x > y) {
        return dir;
      }
    }
    return a.sortOrder - b.sortOrder || a.id - b.id;
  };
}

export interface IGroupRows {
  item: IWorkItem;
  subitems: IWorkItem[];
}

/**
 * Items of each group after filters and sort. A parent shows when it or one of its
 * subitems matches; subitems keep their own order.
 */
export function visibleRows(board: IBoard, items: IWorkItem[], v: IViewState, now: Date = new Date()): { [groupId: string]: IGroupRows[] } {
  const today = todayIso(now);
  const cmp = compareItems(board, v);
  const subsByParent: { [id: number]: IWorkItem[] } = {};
  items.forEach(i => {
    if (i.parentId) {
      (subsByParent[i.parentId] = subsByParent[i.parentId] || []).push(i);
    }
  });
  const filtering = !!(v.search || v.personId !== null || v.dateFilter || activeFilterCount(v) > 0);
  const out: { [groupId: string]: IGroupRows[] } = {};
  const knownGroups = board.config.groups.map(g => g.id);
  const fallbackGroup = knownGroups[0];
  items
    .filter(i => !i.parentId)
    .sort(cmp)
    .forEach(item => {
      const subs = (subsByParent[item.id] || []).slice().sort((a, b) => a.sortOrder - b.sortOrder || a.id - b.id);
      const selfMatch = !filtering || matches(board, item, v, today);
      const subMatches = filtering ? subs.filter(s => matches(board, s, v, today)) : subs;
      if (!selfMatch && subMatches.length === 0) {
        return;
      }
      const groupId = knownGroups.indexOf(item.groupId) >= 0 ? item.groupId : fallbackGroup;
      (out[groupId] = out[groupId] || []).push({ item, subitems: selfMatch ? subs : subMatches });
    });
  return out;
}

/** People who appear in any People column, for the person filter. */
export function peopleOnBoard(board: IBoard, items: IWorkItem[]): IPerson[] {
  const seen: { [id: number]: IPerson } = {};
  const peopleCols = board.config.columns.filter(c => c.type === 'people');
  items.forEach(i => {
    peopleCols.forEach(c => {
      ((i.values[c.id] as IPerson[]) || []).forEach(p => {
        seen[p.id] = p;
      });
    });
  });
  return Object.keys(seen).map(k => seen[Number(k)]).sort((a, b) => a.title.localeCompare(b.title));
}

/** Summary for a column across items: label counts, number sum, date range or checked count. */
export type ColumnSummary =
  | { kind: 'labels'; parts: { label: ILabel; count: number }[]; total: number }
  | { kind: 'sum'; value: number }
  | { kind: 'range'; start: string | null; end: string | null }
  | { kind: 'checked'; checked: number; total: number }
  | { kind: 'none' };

export function summarize(col: IBoardColumn, items: IWorkItem[]): ColumnSummary {
  if (col.type === 'status') {
    const counts: { [text: string]: number } = {};
    items.forEach(i => {
      const t = (i.values[col.id] as string) || '';
      counts[t] = (counts[t] || 0) + 1;
    });
    const parts: { label: ILabel; count: number }[] = [];
    (col.labels || []).forEach(l => {
      if (counts[l.text]) {
        parts.push({ label: l, count: counts[l.text] });
        delete counts[l.text];
      }
    });
    Object.keys(counts).forEach(t => {
      parts.push({ label: t ? { id: '', text: t, color: '#c4c4c4' } : { id: '', text: 'Empty', color: '#e6e9ef' }, count: counts[t] });
    });
    return { kind: 'labels', parts, total: items.length };
  }
  if (col.type === 'number') {
    let sum = 0;
    items.forEach(i => {
      const n = i.values[col.id];
      if (typeof n === 'number') {
        sum += n;
      }
    });
    return { kind: 'sum', value: sum };
  }
  if (col.type === 'date' || col.type === 'timeline') {
    let start: string | null = null;
    let end: string | null = null;
    items.forEach(i => {
      const v = i.values[col.id];
      const s = col.type === 'date' ? (v as string) : v ? (v as ITimelineValue).start : null;
      const e = col.type === 'date' ? (v as string) : v ? (v as ITimelineValue).end : null;
      if (s && (!start || s < start)) {
        start = s;
      }
      if (e && (!end || e > end)) {
        end = e;
      }
    });
    return { kind: 'range', start, end };
  }
  if (col.type === 'checkbox') {
    return { kind: 'checked', checked: items.filter(i => i.values[col.id] === true).length, total: items.length };
  }
  return { kind: 'none' };
}
