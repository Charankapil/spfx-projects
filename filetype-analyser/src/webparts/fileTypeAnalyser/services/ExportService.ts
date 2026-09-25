import { ILibraryNode } from '../models/ILibraryNode';
import { ISiteCollectionOverview } from '../models/ISiteCollectionOverview';
import { IWebNode } from '../models/IWebNode';

function csvEscape(value: string): string {
  // Library titles now round-trip through a shared file in Site Assets, so a
  // title starting with = + - @ would otherwise run as a formula in Excel.
  const safe = /^[=+\-@\t\r]/.test(value) ? `'${value}` : value;
  if (/[",\n\r]/.test(safe)) {
    return `"${safe.replace(/"/g, '""')}"`;
  }
  return safe;
}

function collectWebErrors(web: IWebNode, out: IWebNode[]): void {
  if (web.error) {
    out.push(web);
  }
  for (const child of web.webs) {
    collectWebErrors(child, out);
  }
}

function flattenLibraries(web: IWebNode, out: { webTitle: string; webUrl: string; library: ILibraryNode }[]): void {
  for (const library of web.libraries) {
    out.push({ webTitle: web.title, webUrl: web.url, library });
  }
  for (const child of web.webs) {
    flattenLibraries(child, out);
  }
}

/**
 * Builds a CSV with one row per (web, library, file type) and triggers a
 * client-side download. No server round trip - everything is already in
 * memory from the scan.
 */
export function exportOverviewToCsv(overview: ISiteCollectionOverview): void {
  const rows: { webTitle: string; webUrl: string; library: ILibraryNode }[] = [];
  flattenLibraries(overview.rootWeb, rows);

  const header = [
    'Site',
    'Web',
    'Library',
    'File Type',
    'File Count',
    'Library Item Count',
    'Scan Error'
  ];
  const lines: string[] = [header.map(csvEscape).join(',')];
  const line = (values: string[]): string => values.map(csvEscape).join(',');

  for (const row of rows) {
    const lib = row.library;
    if (lib.fileTypes.length === 0) {
      lines.push(line([overview.siteUrl, row.webUrl, lib.title, '', '0', String(lib.itemCount), lib.error || '']));
      continue;
    }
    for (const stat of lib.fileTypes) {
      lines.push(
        line([
          overview.siteUrl,
          row.webUrl,
          lib.title,
          stat.extension,
          String(stat.count),
          String(lib.itemCount),
          lib.error || ''
        ])
      );
    }
  }

  const failedWebs: IWebNode[] = [];
  collectWebErrors(overview.rootWeb, failedWebs);
  for (const web of failedWebs) {
    lines.push(line([overview.siteUrl, web.url, '', '', '', '', web.error || '']));
  }

  const csvContent = lines.join('\r\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  const timestamp = overview.scanCompletedAt
    ? overview.scanCompletedAt.toISOString().replace(/[:.]/g, '-')
    : new Date().toISOString().replace(/[:.]/g, '-');
  link.href = url;
  link.download = `file-type-overview-${timestamp}.csv`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
