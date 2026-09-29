import { IActivityEntry, IBoardColumn, IBoardConfig } from '../models/types';
import { F } from './fieldMap';
import { formatDay } from './dates';

/* eslint-disable @typescript-eslint/no-explicit-any */

/** One row from /items(id)/versions. Field values are keyed by (encoded) internal name. */
export interface IVersionRow {
  VersionLabel: string;
  Created: string;
  Editor?: { LookupValue?: string; Title?: string };
  [field: string]: any;
}

/** The versions endpoint encodes "_" in internal names as "_x005f_". */
export function versionField(row: IVersionRow, internalName: string): any {
  if (Object.prototype.hasOwnProperty.call(row, internalName)) {
    return row[internalName];
  }
  const encoded = internalName.replace(/_/g, '_x005f_');
  if (Object.prototype.hasOwnProperty.call(row, encoded)) {
    return row[encoded];
  }
  if (Object.prototype.hasOwnProperty.call(row, 'OData_' + internalName)) {
    return row['OData_' + internalName];
  }
  return undefined;
}

function lookupText(v: any): string {
  if (v === null || v === undefined) {
    return '';
  }
  if (Array.isArray(v)) {
    return v.map(lookupText).filter(s => s.length > 0).join(', ');
  }
  if (Array.isArray(v.results)) {
    return lookupText(v.results);
  }
  if (typeof v === 'object') {
    return String(v.LookupValue || v.Title || v.Email || '');
  }
  return String(v);
}

function dateText(v: any): string {
  if (!v) {
    return '';
  }
  const s = String(v);
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? formatDay(s.substring(0, 10)) : s;
}

export function versionValueText(col: IBoardColumn, row: IVersionRow): string {
  const raw = versionField(row, col.field);
  switch (col.type) {
    case 'people':
    case 'dropdown':
      return lookupText(raw);
    case 'checkbox':
      return raw === true || raw === 'true' || raw === 1 ? 'Checked' : 'Unchecked';
    case 'date':
      return dateText(raw);
    case 'timeline': {
      const start = dateText(raw);
      const end = dateText(col.fieldEnd ? versionField(row, col.fieldEnd) : '');
      return start || end ? `${start} – ${end}` : '';
    }
    default:
      return raw === null || raw === undefined ? '' : String(raw);
  }
}

/**
 * Turn version history into "who changed what from X to Y" entries, newest first.
 * `versions` may come in any order; they are sorted by version number.
 */
export function diffVersions(config: IBoardConfig, versions: IVersionRow[]): IActivityEntry[] {
  const sorted = versions.slice().sort((a, b) => parseFloat(a.VersionLabel) - parseFloat(b.VersionLabel));
  const groupTitle = (id: string): string => {
    const g = config.groups.filter(x => x.id === id)[0];
    return g ? g.title : id;
  };
  const entries: IActivityEntry[] = [];
  for (let i = 0; i < sorted.length; i++) {
    const cur = sorted[i];
    const who = (cur.Editor && (cur.Editor.LookupValue || cur.Editor.Title)) || 'Someone';
    if (i === 0) {
      entries.push({ version: cur.VersionLabel, when: cur.Created, who, created: true, changes: [] });
      continue;
    }
    const prev = sorted[i - 1];
    const changes: { column: string; from: string; to: string }[] = [];
    const title = { a: String(versionField(prev, 'Title') || ''), b: String(versionField(cur, 'Title') || '') };
    if (title.a !== title.b) {
      changes.push({ column: 'Item', from: title.a, to: title.b });
    }
    const g = { a: String(versionField(prev, F.GroupId) || ''), b: String(versionField(cur, F.GroupId) || '') };
    if (g.a !== g.b) {
      changes.push({ column: 'Group', from: groupTitle(g.a), to: groupTitle(g.b) });
    }
    config.columns.forEach(col => {
      const a = versionValueText(col, prev);
      const b = versionValueText(col, cur);
      if (a !== b) {
        changes.push({ column: col.title, from: a, to: b });
      }
    });
    if (changes.length > 0) {
      entries.push({ version: cur.VersionLabel, when: cur.Created, who, created: false, changes });
    }
  }
  return entries.reverse();
}
