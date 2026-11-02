/** A table ready for export: column headers plus rows of primitive cell values. */
export type CellValue = string | number | boolean | undefined;

export interface IExportColumn {
  header: string;
  /** Approximate character width used for Excel column sizing. */
  width?: number;
  /** Marks a column as a date (ISO string) so Excel stores a real date. */
  isDate?: boolean;
  /** Left out of PDF tables (IDs and long detail columns); kept in Excel and CSV. */
  pdfHidden?: boolean;
}

export interface IExportTable {
  name: string;
  columns: IExportColumn[];
  rows: CellValue[][];
}

export interface IExportSection {
  title: string;
  /** Headline numbers printed above the table in PDF exports. */
  summary: { label: string; value: string }[];
  tables: IExportTable[];
  /** Chart snapshots (PNG data URLs) included in PDF exports. */
  images?: { title: string; dataUrl: string; width: number; height: number }[];
}

/** Characters that make spreadsheet apps evaluate a cell as a formula (CSV injection). */
const FORMULA_START = /^[=+\-@\t\r]/;

export function safeText(v: string): string {
  return FORMULA_START.test(v) ? `'${v}` : v;
}

function csvCell(v: CellValue): string {
  if (v === undefined || v === null) return '';
  if (typeof v === 'number') return isFinite(v) ? String(v) : '';
  if (typeof v === 'boolean') return v ? 'Yes' : 'No';
  const s = safeText(String(v));
  return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

export function toCsv(table: IExportTable): string {
  const lines = [table.columns.map((c) => csvCell(c.header)).join(',')];
  table.rows.forEach((r) => lines.push(r.map(csvCell).join(',')));
  // BOM so Excel opens UTF-8 correctly.
  return '﻿' + lines.join('\r\n');
}

export function fileStamp(d: Date = new Date()): string {
  const p = (n: number): string => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
}

export function safeFileName(name: string): string {
  return name.replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, '-').replace(/-+/g, '-').slice(0, 80);
}

export function downloadBlob(blob: Blob, fileName: string): void {
  const nav = window.navigator as Navigator & { msSaveOrOpenBlob?: (b: Blob, n: string) => void };
  if (nav.msSaveOrOpenBlob) {
    nav.msSaveOrOpenBlob(blob, fileName);
    return;
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  a.rel = 'noopener';
  a.style.display = 'none';
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 1000);
}

export function downloadCsv(table: IExportTable, baseName?: string): void {
  const csv = toCsv(table);
  downloadBlob(new Blob([csv], { type: 'text/csv;charset=utf-8' }), `${safeFileName(baseName || table.name)}_${fileStamp()}.csv`);
}

export function downloadJson(data: unknown, baseName: string): void {
  downloadBlob(new Blob([JSON.stringify(data, undefined, 2)], { type: 'application/json' }), `${safeFileName(baseName)}_${fileStamp()}.json`);
}
