import { IBoard, IWorkItem } from '../models/types';
import { cellText } from './fieldMap';

export function csvEscape(value: string): string {
  if (/[",\n\r]/.test(value)) {
    return '"' + value.replace(/"/g, '""') + '"';
  }
  return value;
}

/**
 * Board as CSV, in board order: group, item, parent (for subitems), then every visible column.
 * Starts with a BOM so Excel opens UTF-8 correctly.
 */
export function boardToCsv(board: IBoard, items: IWorkItem[]): string {
  const cols = board.config.columns.filter(c => !c.hidden);
  const header = ['Group', 'Item', 'Parent item', 'ID'].concat(cols.map(c => c.title));
  const lines: string[] = [header.map(csvEscape).join(',')];
  const byId: { [id: number]: IWorkItem } = {};
  items.forEach(i => {
    byId[i.id] = i;
  });
  const bySort = (a: IWorkItem, b: IWorkItem): number => a.sortOrder - b.sortOrder || a.id - b.id;
  board.config.groups.forEach(group => {
    const top = items.filter(i => i.groupId === group.id && !i.parentId).sort(bySort);
    top.forEach(item => {
      const rows = [item].concat(items.filter(s => s.parentId === item.id).sort(bySort));
      rows.forEach(row => {
        const parent = row.parentId ? byId[row.parentId] : null;
        const cells = [group.title, row.title, parent ? parent.title : '', board.key + '-' + row.id]
          .concat(cols.map(c => cellText(c, row.values[c.id])));
        lines.push(cells.map(csvEscape).join(','));
      });
    });
  });
  return '﻿' + lines.join('\r\n');
}
