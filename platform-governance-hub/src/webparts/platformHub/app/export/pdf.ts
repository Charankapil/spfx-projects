import { CellValue, downloadBlob, fileStamp, IExportSection, safeFileName } from './tabular';

/** Rows per table in a PDF; the full data belongs in the Excel/CSV exports. */
export const PDF_ROW_LIMIT = 500;

function cellText(v: CellValue): string {
  if (v === undefined || v === null) return '';
  if (typeof v === 'boolean') return v ? 'Yes' : 'No';
  if (typeof v === 'number') return v.toLocaleString('en-US');
  return String(v);
}

/**
 * Builds a landscape A4 report pack: cover, then for each section its headline
 * numbers, chart snapshots and tables. jsPDF is loaded on demand.
 */
export async function downloadPdf(title: string, sections: IExportSection[], tenantLabel: string): Promise<void> {
  const [{ jsPDF }, autoTableModule] = await Promise.all([
    import(/* webpackChunkName: 'pgh-jspdf' */ 'jspdf'),
    import(/* webpackChunkName: 'pgh-jspdf' */ 'jspdf-autotable')
  ]);
  const autoTable = autoTableModule.autoTable || autoTableModule.default;
  if (typeof jsPDF !== 'function' || typeof autoTable !== 'function') throw new Error('The PDF library could not be loaded.');
  const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4', compress: true });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const margin = 36;

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(24);
  doc.text(title, margin, 96);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(12);
  doc.text(`Tenant: ${tenantLabel}`, margin, 124);
  doc.text(`Generated: ${new Date().toLocaleString()}`, margin, 142);
  doc.text(`Sections: ${sections.map((s) => s.title).join(', ')}`, margin, 160, { maxWidth: pageW - margin * 2 });

  sections.forEach((section) => {
    doc.addPage();
    let y = margin + 8;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.text(section.title, margin, y);
    y += 18;
    if (section.summary.length) {
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(10);
      const colW = (pageW - margin * 2) / Math.min(4, section.summary.length);
      section.summary.forEach((k, i) => {
        const col = i % 4;
        const row = Math.floor(i / 4);
        doc.setTextColor(90);
        doc.text(k.label, margin + col * colW, y + row * 30);
        doc.setTextColor(20);
        doc.setFont('helvetica', 'bold');
        doc.text(k.value, margin + col * colW, y + row * 30 + 13);
        doc.setFont('helvetica', 'normal');
      });
      y += Math.ceil(section.summary.length / 4) * 30 + 6;
    }
    (section.images || []).forEach((img) => {
      const w = Math.min(pageW - margin * 2, img.width);
      const h = (img.height / img.width) * w;
      if (y + h + 20 > pageH - margin) {
        doc.addPage();
        y = margin;
      }
      doc.setFontSize(10);
      doc.setTextColor(60);
      doc.text(img.title, margin, y + 10);
      doc.addImage(img.dataUrl, 'PNG', margin, y + 16, w, h);
      y += h + 26;
    });
    doc.setTextColor(20);
    section.tables.forEach((t) => {
      if (y > pageH - margin - 60) {
        doc.addPage();
        y = margin;
      }
      doc.setFontSize(11);
      doc.setFont('helvetica', 'bold');
      const truncated = t.rows.length > PDF_ROW_LIMIT;
      const keep = t.columns.map((c, i) => (c.pdfHidden ? -1 : i)).filter((i) => i >= 0);
      doc.text(`${t.name}${truncated ? ` (first ${PDF_ROW_LIMIT} of ${t.rows.length} rows; see the Excel export for all)` : ` (${t.rows.length} rows)`}`, margin, y + 10);
      doc.setFont('helvetica', 'normal');
      autoTable(doc, {
        startY: y + 16,
        margin: { left: margin, right: margin },
        head: [keep.map((i) => t.columns[i].header)],
        body: t.rows.slice(0, PDF_ROW_LIMIT).map((r) => keep.map((i) => (t.columns[i].isDate && typeof r[i] === 'string' ? String(r[i]).slice(0, 10) : cellText(r[i])))),
        styles: { fontSize: 7.5, cellPadding: 3, overflow: 'linebreak' },
        headStyles: { fillColor: [15, 107, 110], textColor: 255 },
        alternateRowStyles: { fillColor: [245, 247, 248] }
      });
      const last = (doc as unknown as { lastAutoTable?: { finalY: number } }).lastAutoTable;
      y = (last ? last.finalY : y + 40) + 18;
    });
  });

  const pages = doc.getNumberOfPages();
  for (let i = 1; i <= pages; i++) {
    doc.setPage(i);
    doc.setFontSize(8);
    doc.setTextColor(120);
    doc.text(`${title} · ${tenantLabel} · page ${i} of ${pages}`, margin, pageH - 16);
  }
  const blob = doc.output('blob');
  downloadBlob(blob, `${safeFileName(title)}_${fileStamp()}.pdf`);
}
