/** How a site's ownership was found (the permission model assigns owners through Entra ID groups). */
export type OwnershipKind = 'group' | 'people' | 'single' | 'none' | 'unknown';

/** Everything measured for one site. Counts, not percentages, so sites add up into hubs. */
export interface ISiteMetrics {
  hub: string;
  title: string;
  url: string;
  /** Items in the site's document libraries (including folders), from the library counts. */
  documents: number;
  /** Documents actually read; less than `documents` when the per-site limit was reached. */
  scanned: number;
  sampled: boolean;
  /** Per required field: how many scanned documents have a value. */
  fieldFilled: { [field: string]: number };
  /** Required fields that no library in the site has as a column. */
  missingColumns: string[];
  /** Scanned documents with every required field filled in. */
  complete: number;
  labelled: number;
  stale: number;
  staleBytes: number;
  duplicates: number;
  duplicateBytes: number;
  storageBytes?: number;
  ownership: OwnershipKind;
  owners?: number;
  error?: string;
}

/** Sites summed up (a hub, or everything in scope). */
export interface IAggregate {
  sites: number;
  scannedSites: number;
  documents: number;
  scanned: number;
  complete: number;
  labelled: number;
  stale: number;
  staleBytes: number;
  duplicates: number;
  duplicateBytes: number;
  storageBytes: number;
  storageKnown: boolean;
  ownersOk: number;
  ownersKnown: number;
  fieldFilled: { [field: string]: number };
  /** Per required field: in how many sites it is missing as a column. */
  missingColumnSites: { [field: string]: number };
}

export interface IScorecardRun {
  version: 1;
  runAt: string;
  requiredFields: string[];
  staleYears: number;
  sites: ISiteMetrics[];
}

export interface ISnapshot extends IScorecardRun {
  id?: number;
  savedBy?: string;
}

export interface IScanProgress {
  done: number;
  total: number;
  current: string;
}
