/** CSV in and out. Written for Excel: UTF-8 with a byte-order mark, CRLF, and no cell Excel would run as a formula. */

export function csvCell(v: unknown): string {
  let s = v === undefined || v === null ? '' : String(v);
  // A cell starting with = + - @ (or tab / CR) is a formula to Excel: a site title must never be one.
  if (/^[=+\-@\t\r]/.test(s)) {
    s = "'" + s;
  }
  return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

export function toCsv(headers: string[], rows: unknown[][]): string {
  const lines = [headers.map(csvCell).join(',')];
  for (const r of rows) {
    lines.push(r.map(csvCell).join(','));
  }
  return '﻿' + lines.join('\r\n') + '\r\n';
}

/** RFC 4180 reader: quoted fields, doubled quotes, line breaks in quotes, CRLF / LF, BOM, and , ; or tab as the delimiter. */
export function parseCsv(text: string): string[][] {
  const src = text.charCodeAt(0) === 0xfeff ? text.substring(1) : text;
  const firstLine = src.split(/\r?\n/, 1)[0] || '';
  const counts = [',', ';', '\t'].map((d) => ({ d, n: firstLine.split(d).length - 1 })).sort((a, b) => b.n - a.n);
  const d = counts[0].n > 0 ? counts[0].d : ',';
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  const endRow = (): void => {
    row.push(field);
    field = '';
    if (row.length > 1 || row[0].trim() !== '') {
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
    } else if (c === d) {
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
