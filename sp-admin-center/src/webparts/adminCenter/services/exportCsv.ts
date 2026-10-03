/**
 * Browser-side CSV download (UTF-8 with BOM so Excel opens non-English text).
 * Cells that start with = + - @ are prefixed with an apostrophe: SharePoint
 * titles are user-controlled, and without that a title such as =HYPERLINK(...)
 * would run as a formula when the file is opened in Excel (CSV injection).
 */
export function csvCell(value: unknown): string {
  let s = value === null || value === undefined ? '' : value instanceof Date ? value.toISOString() : String(value);
  if (/^[=+\-@\t\r]/.test(s)) {
    s = "'" + s;
  }
  return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

export function toCsv(headers: string[], rows: unknown[][]): string {
  return [headers].concat(rows as string[][]).map((r) => r.map(csvCell).join(',')).join('\r\n');
}

export function downloadCsv(fileName: string, headers: string[], rows: unknown[][]): void {
  const blob = new Blob(['﻿' + toCsv(headers, rows)], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
