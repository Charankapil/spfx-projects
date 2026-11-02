import * as React from 'react';
import { DefaultButton, IContextualMenuItem } from '@fluentui/react';
import { PageId, useApp } from '../AppContext';
import { IExportSection, downloadCsv } from '../../export/tabular';
import { describeError } from '../../http/ApiClient';
import { svgToPngDataUrl } from '../../export/svgToPng';
import { pageDef, loadPrices } from '../exportHelpers';

export type ExportFormat = 'xlsx' | 'pdf' | 'csv';

/** Captures the charts currently rendered inside `root` for the PDF export. */
export async function captureCharts(root: HTMLElement | undefined, background: string): Promise<IExportSection['images']> {
  if (!root) return [];
  const cards = Array.prototype.slice.call(root.querySelectorAll('[data-chart-card]')) as HTMLElement[];
  const images: NonNullable<IExportSection['images']> = [];
  for (const card of cards) {
    const svg = card.querySelector('svg[data-chart]');
    if (!(svg instanceof SVGSVGElement)) continue;
    try {
      const img = await svgToPngDataUrl(svg, background);
      images.push({ title: card.getAttribute('data-chart-card') || '', ...img });
    } catch {
      /* a chart that cannot be rasterised is skipped, not fatal */
    }
  }
  return images;
}

export async function runExport(format: ExportFormat, title: string, sections: IExportSection[], tenantLabel: string): Promise<void> {
  if (format === 'xlsx') {
    const m = await import(/* webpackChunkName: 'pgh-export-xlsx' */ '../../export/excel');
    await m.downloadWorkbook(title, sections, tenantLabel);
  } else if (format === 'pdf') {
    const m = await import(/* webpackChunkName: 'pgh-export-pdf' */ '../../export/pdf');
    await m.downloadPdf(title, sections, tenantLabel);
  } else {
    sections.forEach((s) => s.tables.forEach((t) => downloadCsv(t, `${s.title}-${t.name}`)));
  }
}

/** "Export" button for a single page: Excel workbook, PDF with chart snapshots, or CSV. */
export const ExportMenu: React.FC<{ page: PageId; rootRef: React.RefObject<HTMLElement> }> = ({ page, rootRef }) => {
  const { provider, envFilter, notify, theme } = useApp();
  const [busy, setBusy] = React.useState(false);
  const def = pageDef(page);
  const go = async (format: ExportFormat): Promise<void> => {
    if (!def || !def.build) return;
    setBusy(true);
    try {
      const section = await def.build(provider, envFilter, loadPrices());
      if (format === 'pdf') section.images = await captureCharts(rootRef.current || undefined, theme.semanticColors.bodyBackground);
      await runExport(format, def.title, [section], provider.tenantLabel);
      notify(`${def.title} exported.`, 'success');
    } catch (e) {
      notify(`Export failed: ${describeError(e)}`, 'error');
    } finally {
      setBusy(false);
    }
  };
  const items: IContextualMenuItem[] = [
    { key: 'xlsx', text: 'Excel workbook (.xlsx)', iconProps: { iconName: 'ExcelDocument' }, onClick: () => { go('xlsx').catch(() => undefined); } },
    { key: 'pdf', text: 'PDF report with charts', iconProps: { iconName: 'PDF' }, onClick: () => { go('pdf').catch(() => undefined); } },
    { key: 'csv', text: 'CSV (one file per table)', iconProps: { iconName: 'TextDocument' }, onClick: () => { go('csv').catch(() => undefined); } }
  ];
  return <DefaultButton text={busy ? 'Exporting…' : 'Export'} iconProps={{ iconName: 'Download' }} disabled={busy} menuProps={{ items }} />;
};
