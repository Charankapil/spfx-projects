import { downloadBlob, fileStamp, IExportSection, IExportTable, safeFileName } from './tabular';

/** Excel sheet names: max 31 chars, unique, without []:*?/\ */
export function sheetName(name: string, used: Record<string, boolean>): string {
  const base = name.replace(/[[\]:*?/\\]/g, ' ').trim().slice(0, 31) || 'Sheet';
  let candidate = base;
  let i = 2;
  while (used[candidate.toLowerCase()]) {
    const suffix = ` (${i++})`;
    candidate = base.slice(0, 31 - suffix.length) + suffix;
  }
  used[candidate.toLowerCase()] = true;
  return candidate;
}

/**
 * Builds a multi-sheet .xlsx workbook. ExcelJS is loaded on demand in its own
 * bundle chunk so it does not slow down the dashboard's first load.
 */
export async function downloadWorkbook(title: string, sections: IExportSection[], tenantLabel: string): Promise<void> {
  const mod = await import(/* webpackChunkName: 'lens-exceljs' */ 'exceljs');
  // ExcelJS is CommonJS: bundlers expose it either as named exports or under `default`.
  const ExcelJS = (mod as unknown as { Workbook?: unknown }).Workbook ? mod : (mod as unknown as { default: typeof mod }).default;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Platform Lens';
  wb.created = new Date();
  const used: Record<string, boolean> = {};

  // Summary sheet with every section's headline numbers.
  const summary = wb.addWorksheet(sheetName('Summary', used));
  summary.columns = [{ header: 'Section', width: 32 }, { header: 'Measure', width: 40 }, { header: 'Value', width: 22 }];
  summary.addRow([title, 'Tenant', tenantLabel]);
  summary.addRow([title, 'Exported', new Date().toLocaleString()]);
  sections.forEach((s) => s.summary.forEach((k) => summary.addRow([s.title, k.label, k.value])));
  summary.getRow(1).font = { bold: true };

  const addTable = (t: IExportTable): void => {
    const ws = wb.addWorksheet(sheetName(t.name, used), { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = t.columns.map((c) => ({ header: c.header, width: c.width || Math.min(48, Math.max(10, c.header.length + 4)) }));
    t.rows.forEach((r) => {
      ws.addRow(
        r.map((v, i) => {
          if (v === undefined || v === null) return undefined;
          if (t.columns[i] && t.columns[i].isDate && typeof v === 'string') {
            const d = Date.parse(v);
            return isNaN(d) ? v : new Date(d);
          }
          // xlsx stores strings as text (never formulas), so no injection escaping is needed here.
          return v;
        })
      );
    });
    t.columns.forEach((c, i) => {
      if (c.isDate) ws.getColumn(i + 1).numFmt = 'yyyy-mm-dd';
    });
    const header = ws.getRow(1);
    header.font = { bold: true };
    if (t.rows.length) ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: t.columns.length } };
  };
  sections.forEach((s) => s.tables.forEach(addTable));

  const buffer = await wb.xlsx.writeBuffer();
  downloadBlob(
    new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
    `${safeFileName(title)}_${fileStamp()}.xlsx`
  );
}
