import { ILibraryNode } from '../models/ILibraryNode';
import { ISiteCollectionOverview } from '../models/ISiteCollectionOverview';
import { IWebNode } from '../models/IWebNode';

function csvEscape(value: string): string {
  if (/[",\n]/.test(value)) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
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

  const header = ['Site', 'Web', 'Library', 'File Type', 'File Count', 'Library Item Count', 'Scan Error'];
  const lines: string[] = [header.map(csvEscape).join(',')];

  for (const row of rows) {
    if (row.library.fileTypes.length === 0) {
      lines.push(
        [
          overview.siteUrl,
          row.webUrl,
          row.library.title,
          '',
          '0',
          String(row.library.itemCount),
          row.library.error || ''
        ]
          .map((v) => csvEscape(String(v)))
          .join(',')
      );
      continue;
    }
    for (const stat of row.library.fileTypes) {
      lines.push(
        [
          overview.siteUrl,
          row.webUrl,
          row.library.title,
          stat.extension,
          String(stat.count),
          String(row.library.itemCount),
          row.library.error || ''
        ]
          .map((v) => csvEscape(String(v)))
          .join(',')
      );
    }
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
