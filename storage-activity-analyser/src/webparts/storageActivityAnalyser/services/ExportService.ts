import { IScanResult } from '../models/IScanResult';
import { AGE_BANDS, bandTotals, describeAge, newestFileAge, splitByThreshold, thresholdLabel } from './activity';

function csvEscape(value: string): string {
  // Titles and file names come from the site, so a leading = + - @ must not run as a formula in Excel.
  const safe = /^[=+\-@\t\r]/.test(value) ? `'${value}` : value;
  if (/[",\n\r]/.test(safe)) {
    return `"${safe.replace(/"/g, '""')}"`;
  }
  return safe;
}

function download(lines: string[][], fileName: string): void {
  // The BOM makes Excel open the file as UTF-8, so non-English titles survive.
  const csv = '﻿' + lines.map((l) => l.map(csvEscape).join(',')).join('\r\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function stamp(result: IScanResult): string {
  return result.scanCompletedAt.substring(0, 10);
}

/** One row per library with the active / inactive split for the chosen threshold and bytes per age band. */
export function exportLibrariesCsv(result: IScanResult, thresholdMonths: number): void {
  const label = thresholdLabel(thresholdMonths);
  const header = [
    'Site',
    'Site URL',
    'Library',
    'Library URL',
    'Files',
    'Size of current files (bytes)',
    `Active files (changed in last ${label})`,
    'Active size (bytes)',
    `Inactive files (not changed for ${label}+)`,
    'Inactive size (bytes)',
    'Inactive share of size (%)',
    'Newest file changed',
    'Library size incl. version history (bytes)',
    ...AGE_BANDS.map((b) => `Size ${b.label} (bytes)`),
    'Error'
  ];
  const lines: string[][] = [header];
  for (const lib of result.libraries) {
    const split = splitByThreshold(lib.histogram, thresholdMonths);
    const newest = newestFileAge(lib.histogram);
    lines.push([
      lib.webTitle,
      lib.webUrl,
      lib.title,
      lib.url,
      String(split.totalFiles),
      String(split.totalBytes),
      String(split.activeFiles),
      String(split.activeBytes),
      String(split.inactiveFiles),
      String(split.inactiveBytes),
      split.totalBytes ? ((split.inactiveBytes / split.totalBytes) * 100).toFixed(1) : '0',
      newest === undefined ? '' : describeAge(newest),
      lib.metrics ? String(lib.metrics.totalSize) : '',
      ...AGE_BANDS.map((b) => String(bandTotals(lib.histogram, b).bytes)),
      lib.error
        ? `${lib.partial ? 'Partly read: ' : ''}${lib.error}`
        : lib.unreadItems
        ? `${lib.unreadItems} of ${lib.itemCount} items not visible to the person who ran the scan`
        : ''
    ]);
  }
  for (const web of result.webs) {
    if (web.error) {
      lines.push([web.title, web.url, '', '', '', '', '', '', '', '', '', '', '', ...AGE_BANDS.map(() => ''), web.error]);
    }
  }
  download(lines, `storage-activity-libraries-${stamp(result)}.csv`);
}

export function exportLargestFilesCsv(result: IScanResult, thresholdMonths: number): void {
  const origin = new URL(result.rootUrl).origin;
  const lines: string[][] = [['File', 'Site', 'Library', 'Size (bytes)', 'Last modified', 'URL']];
  for (const file of result.largestOldFiles) {
    if (file.ageMonths >= thresholdMonths) {
      lines.push([
        file.name,
        file.webTitle,
        file.libraryTitle,
        String(file.bytes),
        file.modified,
        `${origin}${file.serverRelativeUrl}`
      ]);
    }
  }
  download(lines, `storage-activity-largest-inactive-files-${stamp(result)}.csv`);
}
