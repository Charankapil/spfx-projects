import { IRestoreOptions } from './IRestoreOptions';
import { ObjectKind, IUniqueObject } from './IUniqueObject';

export interface IRunStats {
  websChecked: number;
  listsChecked: number;
  itemsChecked: number;
  uniqueFound: number;
  restored: number;
  failed: number;
  excluded: number;
  skipped: number;
}

export interface IScopeRef {
  kind: string;
  title: string;
  url: string;
}

export type RunMode = 'scan' | 'restore';

/** One scan (and, if it got that far, its restore) - saved and reported on. */
export interface IRunReport {
  id: string;
  siteUrl: string;
  mode: RunMode;
  runBy: string;
  startedAt: string;
  scanCompletedAt?: string;
  restoreStartedAt?: string;
  completedAt?: string;
  cancelled?: boolean;
  scopes: IScopeRef[];
  options: IRestoreOptions;
  stats: IRunStats;
  /** Everything with unique permissions that was found in scope. */
  objects: IUniqueObject[];
  /** Parts of the scope that could not be read (403 on a subsite and so on). */
  scanErrors: string[];
}

/** Kept in the report index so the history list doesn't download every report. */
export interface IReportSummary {
  id: string;
  fileName: string;
  mode: RunMode;
  runBy: string;
  startedAt: string;
  completedAt?: string;
  cancelled?: boolean;
  scopes: IScopeRef[];
  stats: IRunStats;
}

export const EMPTY_STATS: IRunStats = {
  websChecked: 0,
  listsChecked: 0,
  itemsChecked: 0,
  uniqueFound: 0,
  restored: 0,
  failed: 0,
  excluded: 0,
  skipped: 0
};

export const KIND_LABELS: { [kind in ObjectKind]: string } = {
  web: 'Site',
  library: 'Library',
  list: 'List',
  folder: 'Folder',
  file: 'File',
  item: 'List item'
};
