import { IScanResult } from '../models/IScanResult';
import { ageInMonths } from './activity';

export interface IFolderRow {
  key: string;
  libraryKey: string;
  libraryTitle: string;
  webTitle: string;
  libraryUrl: string;
  path: string;
  depth: number;
  files: number;
  bytes: number;
  versionBytes: number;
  totalSize: number;
  lastModified?: string;
  /** Whole months since anything in the folder changed, measured from when the figures were taken. */
  ageMonths?: number;
  /** Nothing in the folder has changed for the threshold or longer. */
  inactive: boolean;
  childrenUnavailable: boolean;
}

/** Every folder in the result's folder breakdowns, biggest first. */
export function folderRows(result: IScanResult, thresholdMonths: number): IFolderRow[] {
  const rows: IFolderRow[] = [];
  for (const lib of result.libraries) {
    if (!lib.folders) {
      continue;
    }
    const measuredAt = new Date(lib.foldersAt || result.scanStartedAt);
    const libraryKey = `${lib.webUrl}|${lib.id}`;
    for (const f of lib.folders) {
      const changed = f.lastModified ? new Date(f.lastModified) : undefined;
      const age = changed && !isNaN(changed.getTime()) ? ageInMonths(changed, measuredAt) : undefined;
      rows.push({
        key: `${libraryKey}|${f.path}`,
        libraryKey,
        libraryTitle: lib.title,
        webTitle: lib.webTitle,
        libraryUrl: lib.url,
        path: f.path,
        depth: f.depth,
        files: f.files,
        bytes: f.bytes,
        versionBytes: Math.max(0, f.totalSize - f.bytes),
        totalSize: f.totalSize,
        lastModified: f.lastModified,
        ageMonths: age,
        inactive: age !== undefined && age >= thresholdMonths && f.files > 0,
        childrenUnavailable: !!f.childrenUnavailable
      });
    }
  }
  rows.sort((a, b) => b.totalSize - a.totalSize);
  return rows;
}

/**
 * Inactive folders without an inactive parent, so nested folders are not
 * counted twice. Exact: SharePoint reports the newest change in each
 * folder, so every file in these folders is at least that old.
 */
export function inactiveFolders(rows: IFolderRow[]): { folders: IFolderRow[]; files: number; bytes: number; totalSize: number } {
  const inactive = rows.filter((r) => r.inactive);
  const top = inactive.filter(
    (r) => !inactive.some((p) => p !== r && p.libraryKey === r.libraryKey && r.path.indexOf(p.path + '/') === 0)
  );
  return {
    folders: top,
    files: top.reduce((s, r) => s + r.files, 0),
    bytes: top.reduce((s, r) => s + r.bytes, 0),
    totalSize: top.reduce((s, r) => s + r.totalSize, 0)
  };
}
