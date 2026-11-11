/**
 * A value starting with = + - @ would run as a formula when the file is
 * opened in Excel, so those are prefixed with an apostrophe.
 */
export function csvEscape(value: string): string {
  const safe = /^[=+\-@\t\r]/.test(value) ? `'${value}` : value;
  if (/[",\n\r]/.test(safe)) {
    return `"${safe.replace(/"/g, '""')}"`;
  }
  return safe;
}

export function toCsv(header: string[], rows: Array<Array<string | number | undefined>>): string {
  const lines = [header.map(csvEscape).join(',')];
  for (const row of rows) {
    lines.push(
      row
        // Numbers are written as they are: "-5" is a value, not a formula.
        .map((cell) => (typeof cell === 'number' ? String(cell) : csvEscape(cell === undefined || cell === null ? '' : String(cell))))
        .join(',')
    );
  }
  return lines.join('\r\n');
}

/** Builds the CSV in memory and hands it to the browser: no server round trip. */
export function downloadCsv(fileNameStem: string, header: string[], rows: Array<Array<string | number | undefined>>): void {
  // The BOM makes Excel read the file as UTF-8.
  const blob = new Blob(['﻿' + toCsv(header, rows)], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${fileNameStem}-${new Date().toISOString().substring(0, 19).replace(/[:T]/g, '-')}.csv`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
