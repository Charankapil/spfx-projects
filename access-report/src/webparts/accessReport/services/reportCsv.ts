/**
 * The report file format, written identically by the web part and by the Power Automate flow:
 * UTF-8 with a byte-order mark (so Excel opens non-English text), CRLF line ends, and every
 * cell quoted. Cells that start with = + - @ (or a tab / carriage return) get a leading
 * apostrophe: site and library titles are user-controlled, and without it a title such as
 * =HYPERLINK(...) would run as a formula when the file is opened in Excel (CSV injection).
 */

export const REPORT_HEADERS = ['Who', 'Site URL', 'Site title', 'Object', 'Name', 'URL', 'Access', 'How', 'Own permissions'];
export const UNREAD_HEADERS = ['Who', 'Site URL', 'Site title', 'Step', 'Status', 'Error'];
export const BOM = '﻿';

export function reportCell(value: unknown): string {
  let s = value === null || value === undefined ? '' : String(value);
  if (/^[=+\-@\t\r]/.test(s)) {
    s = "'" + s;
  }
  return '"' + s.replace(/"/g, '""') + '"';
}

export function reportLine(cells: unknown[]): string {
  return cells.map(reportCell).join(',');
}

/** Header + rows, each line ended by CRLF, behind a byte-order mark. */
export function toReportCsv(headers: string[], rows: unknown[][]): string {
  return BOM + [headers].concat(rows as string[][]).map((r) => reportLine(r) + '\r\n').join('');
}

/** yyyy-MM-dd-HHmm in UTC: sorts by time, and is what the flow stamps its files with. */
export function fileStamp(d: Date): string {
  const p = (n: number): string => (n < 10 ? '0' : '') + n;
  return `${d.getUTCFullYear()}-${p(d.getUTCMonth() + 1)}-${p(d.getUTCDate())}-${p(d.getUTCHours())}${p(d.getUTCMinutes())}`;
}

export const REPORT_PREFIX = 'access-report-';
export const UNREAD_SUFFIX = '-unread';

export function reportFileName(stamp: string): string {
  return `${REPORT_PREFIX}${stamp}.csv`;
}

export function unreadFileName(stamp: string): string {
  return `${REPORT_PREFIX}${stamp}${UNREAD_SUFFIX}.csv`;
}

/** The time in a report's file name, or undefined when it is not one of ours. */
export function stampOf(fileName: string): { stamp: string; unread: boolean; at: Date } | undefined {
  const m = /^access-report-(\d{4})-(\d{2})-(\d{2})-(\d{2})(\d{2})(-unread)?\.csv$/i.exec(fileName);
  if (!m) {
    return undefined;
  }
  const at = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5])));
  return isNaN(at.getTime()) ? undefined : { stamp: `${m[1]}-${m[2]}-${m[3]}-${m[4]}${m[5]}`, unread: !!m[6], at };
}

// ---- reading ---------------------------------------------------------------------------

export function stripBom(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.substring(1) : text;
}

/** Small RFC 4180 reader: quoted fields, doubled quotes, line breaks inside quotes, CRLF / LF. */
export function parseCsv(text: string): string[][] {
  const src = stripBom(text);
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  const endRow = (): void => {
    row.push(field);
    field = '';
    if (row.length > 1 || row[0] !== '') {
      rows.push(row);
    }
    row = [];
  };
  for (let i = 0; i < src.length; i++) {
    const c = src.charAt(i);
    if (inQuotes) {
      if (c === '"') {
        if (src.charAt(i + 1) === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += c;
      }
    } else if (c === '"' && field === '') {
      inQuotes = true;
    } else if (c === ',') {
      row.push(field);
      field = '';
    } else if (c === '\n' || c === '\r') {
      if (c === '\r' && src.charAt(i + 1) === '\n') {
        i++;
      }
      endRow();
    } else {
      field += c;
    }
  }
  if (field !== '' || row.length > 0) {
    endRow();
  }
  return rows;
}

/** The rows of a report file (header dropped, the formula guard's apostrophe removed). */
export function parseReport(text: string): string[][] {
  const rows = parseCsv(text);
  const unguard = (s: string): string => (/^'[=+\-@\t\r]/.test(s) ? s.substring(1) : s);
  return rows.slice(1).map((r) => r.map(unguard));
}

export function downloadText(fileName: string, text: string): void {
  const blob = new Blob([text], { type: 'text/csv;charset=utf-8' });
  downloadBlob(fileName, blob);
}

export function downloadBlob(fileName: string, blob: Blob): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
