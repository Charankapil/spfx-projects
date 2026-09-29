import { IRunReport, KIND_LABELS } from '../models/IRunReport';
import { ObjectStatus } from '../models/IUniqueObject';

export const STATUS_LABELS: { [status in ObjectStatus]: string } = {
  found: 'Unique permissions (not restored)',
  restored: 'Inheritance restored',
  failed: 'Failed',
  excluded: 'Excluded',
  skipped: 'Not processed'
};

function csvEscape(value: string): string {
  // File and folder names are user content, so a name starting with = + - @
  // would otherwise run as a formula in Excel.
  const safe = /^[=+\-@\t\r]/.test(value) ? `'${value}` : value;
  if (/[",\n\r]/.test(safe)) {
    return `"${safe.replace(/"/g, '""')}"`;
  }
  return safe;
}

/** One row per object that had unique permissions, then any scan errors. */
export function reportToCsv(report: IRunReport): string {
  const header = [
    'Status',
    'Type',
    'Name',
    'Path',
    'Site',
    'List',
    'Item ID',
    'Folder depth',
    'Permissions before restore',
    'Restored at',
    'Message'
  ];
  const line = (values: string[]): string => values.map(csvEscape).join(',');
  const lines: string[] = [line(header)];
  for (const o of report.objects) {
    lines.push(
      line([
        STATUS_LABELS[o.status],
        KIND_LABELS[o.kind],
        o.name,
        o.path,
        o.webUrl,
        o.listTitle || '',
        o.itemId !== undefined ? String(o.itemId) : '',
        String(o.depth),
        o.previousPermissions || '',
        o.restoredAt || '',
        o.message || ''
      ])
    );
  }
  for (const error of report.scanErrors) {
    lines.push(line(['Scan error', '', '', '', '', '', '', '', '', '', error]));
  }
  return lines.join('\r\n');
}

export function downloadCsv(report: IRunReport): void {
  // The BOM makes Excel open the file as UTF-8, so non-ASCII names survive.
  const blob = new Blob(['﻿', reportToCsv(report)], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `reinherit-${report.mode}-${report.startedAt.replace(/[:.]/g, '-')}.csv`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
