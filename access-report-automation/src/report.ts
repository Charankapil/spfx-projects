import { gunzipSync, gzipSync } from 'node:zlib';
import { toCsv } from './csv.js';
import { LEVELS } from './permissions.js';
import { AccessRow } from './scanner.js';
import { buildXlsx, Cell, Sheet } from './xlsx.js';

/** The monthly deliverables: an Excel workbook, CSVs of the same data, and the changes since the previous report. */

export interface PrincipalInfo {
  key: string;
  label: string;
  kind: string;
}

export interface FailedSite {
  url: string;
  title: string;
  error: string;
}

export interface Snapshot {
  v: 1;
  period: string;
  generatedAt: string;
  principals: PrincipalInfo[];
  failedSites: string[];
  rows: AccessRow[];
}

export interface Change {
  change: 'Added' | 'Removed' | 'More access' | 'Less access' | 'Changed';
  row: AccessRow;
  before?: AccessRow;
}

export function packSnapshot(s: Snapshot): Uint8Array {
  return gzipSync(Buffer.from(JSON.stringify(s), 'utf8'));
}

export function unpackSnapshot(data: Uint8Array): Snapshot {
  const s = JSON.parse(gunzipSync(data).toString('utf8')) as Snapshot;
  if (!s || s.v !== 1 || !Array.isArray(s.rows) || !Array.isArray(s.principals)) {
    throw new Error('Not a snapshot this version understands.');
  }
  return s;
}

const OBJECT_ORDER: Record<string, number> = { Site: 0, Subsite: 1, Library: 2, List: 3 };

export function sortRows(rows: AccessRow[], labels: Map<string, string>): AccessRow[] {
  const cmp = (a: string, b: string): number => (a < b ? -1 : a > b ? 1 : 0);
  return rows.slice().sort(
    (a, b) =>
      cmp((labels.get(a.p) || a.p).toLowerCase(), (labels.get(b.p) || b.p).toLowerCase()) ||
      cmp(a.siteUrl.toLowerCase(), b.siteUrl.toLowerCase()) ||
      (OBJECT_ORDER[a.object] ?? 9) - (OBJECT_ORDER[b.object] ?? 9) ||
      cmp(a.url.toLowerCase(), b.url.toLowerCase()) ||
      cmp(a.title.toLowerCase(), b.title.toLowerCase())
  );
}

/** A renamed library keeps its address, so the title is not part of what identifies a row. */
const rowKey = (r: AccessRow): string => `${r.p}\n${r.object}\n${r.url.toLowerCase()}`;

/**
 * What changed since the previous report. Only people and groups that were in both reports
 * are compared, and sites that could not be read (now or then) are left out: an unread site
 * is not a removed permission.
 */
export function diffSnapshots(previous: Snapshot, current: Snapshot): Change[] {
  const inBoth = new Set(current.principals.map((p) => p.key).filter((k) => previous.principals.some((p) => p.key === k)));
  const unread = new Set(previous.failedSites.concat(current.failedSites).map((u) => u.toLowerCase()));
  const usable = (r: AccessRow): boolean => inBoth.has(r.p) && !unread.has(r.siteUrl.toLowerCase());
  const before = new Map<string, AccessRow>();
  previous.rows.filter(usable).forEach((r) => before.set(rowKey(r), r));
  const changes: Change[] = [];
  const seen = new Set<string>();
  current.rows.filter(usable).forEach((r) => {
    const k = rowKey(r);
    seen.add(k);
    const old = before.get(k);
    if (!old) {
      changes.push({ change: 'Added', row: r });
    } else if (old.access !== r.access) {
      const up = LEVELS.indexOf(r.access) > LEVELS.indexOf(old.access);
      changes.push({ change: LEVELS.indexOf(r.access) < 0 || LEVELS.indexOf(old.access) < 0 ? 'Changed' : up ? 'More access' : 'Less access', row: r, before: old });
    } else if (old.levels !== r.levels || old.how !== r.how) {
      changes.push({ change: 'Changed', row: r, before: old });
    }
  });
  before.forEach((old, k) => {
    if (!seen.has(k)) {
      changes.push({ change: 'Removed', row: old, before: old });
    }
  });
  return changes;
}

export const ACCESS_HEADERS = ['Who', 'Site URL', 'Site title', 'Object', 'Name', 'URL', 'Access', 'Permission levels', 'How', 'Own permissions', 'Note'];
export const CHANGE_HEADERS = ['Change', 'Who', 'Site URL', 'Site title', 'Object', 'Name', 'URL', 'Access before', 'Access now', 'Permission levels before', 'Permission levels now', 'How before', 'How now'];
export const ISSUE_HEADERS = ['Kind', 'What', 'Title', 'Problem'];

export interface ReportInput {
  name: string;
  period: string;
  generatedAt: Date;
  principals: PrincipalInfo[];
  rows: AccessRow[];
  failedSites: FailedSite[];
  unresolved: Array<{ input: string; type: string; reason: string }>;
  warnings: string[];
  sitesTotal: number;
  previous?: Snapshot;
  scope: string;
  /** Rows the workbook may hold before it points to the CSV instead. */
  maxSheetRows?: number;
}

export interface ReportFiles {
  workbook: Buffer;
  accessCsv: string;
  changesCsv?: string;
  issuesCsv?: string;
  snapshot: Snapshot;
  changes: Change[];
  counts: { rows: number; sitesWithAccess: number; added: number; removed: number; changed: number };
}

export function buildReport(input: ReportInput): ReportFiles {
  const labels = new Map(input.principals.map((p) => [p.key, p.label] as [string, string]));
  const who = (key: string): string => labels.get(key) || key;
  const rows = sortRows(input.rows, labels);
  const snapshot: Snapshot = {
    v: 1,
    period: input.period,
    generatedAt: input.generatedAt.toISOString(),
    principals: input.principals,
    failedSites: input.failedSites.map((f) => f.url),
    rows
  };
  const changes = input.previous ? diffSnapshots(input.previous, snapshot) : [];
  const order: Record<string, number> = { Added: 0, 'More access': 1, 'Less access': 2, Changed: 3, Removed: 4 };
  changes.sort((a, b) => order[a.change] - order[b.change] || (who(a.row.p) < who(b.row.p) ? -1 : who(a.row.p) > who(b.row.p) ? 1 : 0));

  const accessCells = rows.map((r): Cell[] => [who(r.p), r.siteUrl, r.siteTitle, r.object, r.title, r.url, r.access, r.levels, r.how, r.object === 'Site' ? '' : r.unique ? 'Yes' : 'No', r.note]);
  const changeCells = changes.map((c): Cell[] => {
    const now = c.change === 'Removed' ? undefined : c.row;
    const old = c.before;
    return [c.change, who(c.row.p), c.row.siteUrl, c.row.siteTitle, c.row.object, c.row.title, c.row.url, old ? old.access : '', now ? now.access : '', old ? old.levels : '', now ? now.levels : '', old ? old.how : '', now ? now.how : ''];
  });
  const issueCells: Cell[][] = input.unresolved
    .map((u): Cell[] => ['Not in the directory', u.input, u.type, u.reason])
    .concat(input.failedSites.map((f): Cell[] => ['Site not read', f.url, f.title, f.error]));

  // ---- summary ----
  const count = (pred: (r: AccessRow) => boolean): number => rows.reduce((n, r) => n + (pred(r) ? 1 : 0), 0);
  const firstFor = new Set(input.principals.filter((p) => !input.previous || !input.previous.principals.some((q) => q.key === p.key)).map((p) => p.key));
  const summary: Cell[][] = [];
  const bold: number[] = [];
  const title = (t: string): void => {
    bold.push(summary.length);
    summary.push([t]);
  };
  title(`${input.name} ${input.period}`);
  summary.push(['Generated (UTC)', input.generatedAt.toISOString().replace('T', ' ').substring(0, 16)]);
  summary.push(['Compared with', input.previous ? `the ${input.previous.period} report` : 'nothing (first report)']);
  summary.push(['Site collections checked', input.sitesTotal - input.failedSites.length]);
  summary.push(['Site collections not read', input.failedSites.length]);
  summary.push(['People and groups', input.principals.length]);
  summary.push(['Places with access', rows.length]);
  summary.push(['Scope', input.scope]);
  summary.push([]);
  bold.push(summary.length);
  summary.push(['Who', 'Sites', 'Subsites', 'Libraries and lists', 'With own permissions', 'Full control', 'Highest access', 'Added', 'More access', 'Less access', 'Removed', 'Unconfirmed']);
  input.principals
    .slice()
    .sort((a, b) => (a.label.toLowerCase() < b.label.toLowerCase() ? -1 : 1))
    .forEach((p) => {
      const mine = (r: AccessRow): boolean => r.p === p.key;
      const highest = rows.filter(mine).reduce((best, r) => (LEVELS.indexOf(r.access) > LEVELS.indexOf(best) ? r.access : best), '');
      const ch = (kind: Change['change']): Cell => (firstFor.has(p.key) ? (input.previous ? 'new' : '') : changes.filter((c) => c.row.p === p.key && c.change === kind).length);
      summary.push([
        p.label,
        count((r) => mine(r) && r.object === 'Site'),
        count((r) => mine(r) && r.object === 'Subsite'),
        count((r) => mine(r) && (r.object === 'Library' || r.object === 'List')),
        count((r) => mine(r) && r.object !== 'Site' && r.unique),
        count((r) => mine(r) && r.access === 'Full control'),
        highest || 'No access found',
        ch('Added'),
        ch('More access'),
        ch('Less access'),
        ch('Removed'),
        count((r) => mine(r) && !!r.note)
      ]);
    });
  const notes = input.warnings.slice();
  if (input.unresolved.length) {
    notes.push(`${input.unresolved.length} entr${input.unresolved.length === 1 ? 'y' : 'ies'} in the list of people and groups could not be found in the directory (see the "Not read" sheet).`);
  }
  if (input.failedSites.length) {
    notes.push(`${input.failedSites.length} site collection${input.failedSites.length === 1 ? '' : 's'} could not be read (see the "Not read" sheet). Access there is unknown, and they are left out of the changes.`);
  }
  const cap = input.maxSheetRows ?? 500000;
  if (rows.length > cap) {
    notes.push(`The Access sheet shows the first ${cap.toLocaleString('en-US')} of ${rows.length.toLocaleString('en-US')} rows. The CSV file in the csv folder has all of them.`);
  }
  notes.push('Access is reported for sites, subsites and libraries. Folders, files and items with their own permissions (including sharing links) are not listed.');
  summary.push([]);
  title('Notes');
  notes.forEach((n) => summary.push([n]));

  const sheets: Sheet[] = [
    { name: 'Summary', widths: [44, 12, 12, 18, 20, 12, 18, 10, 12, 12, 10, 13], rows: summary, bold },
    { name: 'Access', widths: [38, 48, 28, 10, 28, 56, 14, 28, 44, 16, 40], rows: [ACCESS_HEADERS as Cell[]].concat(accessCells.slice(0, cap)), header: true },
    { name: 'Changes', widths: [13, 38, 48, 28, 10, 28, 56, 14, 14, 26, 26, 40, 40], rows: [CHANGE_HEADERS as Cell[]].concat(changeCells.slice(0, cap)), header: true },
    { name: 'Not read', widths: [22, 56, 30, 70], rows: [ISSUE_HEADERS as Cell[]].concat(issueCells), header: true }
  ];

  return {
    workbook: buildXlsx(sheets),
    accessCsv: toCsv(ACCESS_HEADERS, accessCells),
    changesCsv: input.previous ? toCsv(CHANGE_HEADERS, changeCells) : undefined,
    issuesCsv: issueCells.length ? toCsv(ISSUE_HEADERS, issueCells) : undefined,
    snapshot,
    changes,
    counts: {
      rows: rows.length,
      sitesWithAccess: new Set(rows.map((r) => r.siteUrl.toLowerCase())).size,
      added: changes.filter((c) => c.change === 'Added').length,
      removed: changes.filter((c) => c.change === 'Removed').length,
      changed: changes.filter((c) => c.change !== 'Added' && c.change !== 'Removed').length
    }
  };
}
