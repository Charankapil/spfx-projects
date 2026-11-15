/**
 * What changed between two reports (rows in REPORT_HEADERS order): access that is new,
 * access that is gone, and access whose level changed. A row is identified by who, where
 * and how; its level and "own permissions" flag are what can change.
 */

export type ChangeKind = 'Added' | 'Removed' | 'Changed';

export interface IChange {
  kind: ChangeKind;
  /** The current row (Added / Changed) or the earlier one (Removed). */
  row: string[];
  /** Changed only: the earlier access level. */
  was?: string;
}

const WHO = 0;
const SITE = 1;
const OBJECT = 3;
const NAME = 4;
const URL = 5;
const LEVEL = 6;
const HOW = 7;

function keyOf(r: string[]): string {
  return [r[WHO], r[SITE], r[OBJECT], r[URL] || r[NAME], r[HOW]].map((c) => (c || '').toLowerCase()).join('\u0001');
}

export function compareReports(earlier: string[][], current: string[][]): IChange[] {
  const before: { [key: string]: string[][] } = {};
  earlier.forEach((r) => {
    const k = keyOf(r);
    (before[k] = before[k] || []).push(r);
  });
  const out: IChange[] = [];
  current.forEach((r) => {
    const same = before[keyOf(r)];
    if (!same || !same.length) {
      out.push({ kind: 'Added', row: r });
      return;
    }
    // Several grants can share a key (two custom levels through one group): pair an equal level first.
    let at = -1;
    same.forEach((s, i) => {
      if (at < 0 && s[LEVEL] === r[LEVEL]) {
        at = i;
      }
    });
    const old = same.splice(at < 0 ? 0 : at, 1)[0];
    if (old[LEVEL] !== r[LEVEL]) {
      out.push({ kind: 'Changed', row: r, was: old[LEVEL] });
    }
  });
  Object.keys(before).forEach((k) => before[k].forEach((r) => out.push({ kind: 'Removed', row: r })));
  const order: { [k: string]: number } = { Added: 0, Changed: 1, Removed: 2 };
  return out
    .map((c, n) => ({ c, n }))
    .sort((a, b) => order[a.c.kind] - order[b.c.kind] || a.n - b.n)
    .map((o) => o.c);
}

export const CHANGE_HEADERS = ['Change', 'Who', 'Site URL', 'Site title', 'Object', 'Name', 'URL', 'Access', 'Was', 'How'];

export function changeRows(changes: IChange[]): string[][] {
  return changes.map((c) => [c.kind, c.row[0], c.row[1], c.row[2], c.row[3], c.row[4], c.row[5], c.row[6], c.was || '', c.row[7]]);
}
