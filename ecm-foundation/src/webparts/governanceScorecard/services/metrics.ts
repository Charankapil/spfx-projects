import { IDocumentRecord, IMetadataField } from '../../metadataExplorer/models/IMetadata';
import { IAggregate, ISiteMetrics, OwnershipKind } from '../models/IScorecard';

/**
 * Score weights. Metadata is the biggest lever in the strategy (search,
 * retention and AI readiness all depend on it), then retention labels, then
 * freshness (stale content to archive), then accountable ownership.
 */
export const WEIGHTS = { metadata: 40, labelled: 30, fresh: 20, ownership: 10 };

export type Status = 'good' | 'warning' | 'critical';

export const STATUS_THRESHOLDS = { good: 80, warning: 60 };

export function statusOf(score: number | undefined): Status | undefined {
  if (score === undefined) {
    return undefined;
  }
  return score >= STATUS_THRESHOLDS.good ? 'good' : score >= STATUS_THRESHOLDS.warning ? 'warning' : 'critical';
}

export function pct(part: number, whole: number): number | undefined {
  return whole > 0 ? (part / whole) * 100 : undefined;
}

function ownershipOk(kind: OwnershipKind): boolean | undefined {
  if (kind === 'unknown') {
    return undefined;
  }
  return kind === 'group' || kind === 'people';
}

/**
 * Measures one site from its scanned documents. A required field is matched
 * to a column by display name or internal name, case-insensitively; a site
 * without that column counts every document as missing it (that is the gap).
 */
export function measureSite(
  base: Pick<ISiteMetrics, 'hub' | 'title' | 'url'>,
  documents: IDocumentRecord[],
  fields: IMetadataField[],
  libraryItemCount: number,
  sampled: boolean,
  requiredFields: string[],
  staleYears: number,
  now: Date
): ISiteMetrics {
  const keyFor: { [required: string]: string | undefined } = {};
  const missingColumns: string[] = [];
  requiredFields.forEach((required) => {
    const r = required.toLowerCase();
    const field = fields.filter((f) => f.title.toLowerCase() === r || f.key.toLowerCase() === r)[0];
    keyFor[required] = field?.key;
    if (!field) {
      missingColumns.push(required);
    }
  });
  const labelKey = '_ComplianceTag';
  const staleBefore = new Date(now.getTime());
  staleBefore.setFullYear(staleBefore.getFullYear() - staleYears);

  const fieldFilled: { [field: string]: number } = {};
  requiredFields.forEach((f) => (fieldFilled[f] = 0));
  let complete = 0;
  let labelled = 0;
  let stale = 0;
  let staleBytes = 0;
  const seen: { [key: string]: boolean } = {};
  let duplicates = 0;
  let duplicateBytes = 0;

  documents.forEach((doc) => {
    let all = true;
    requiredFields.forEach((required) => {
      const key = keyFor[required];
      const values = key ? doc.values[key] : undefined;
      if (values && values.length) {
        fieldFilled[required]++;
      } else {
        all = false;
      }
    });
    if (all) {
      complete++;
    }
    const label = doc.values[labelKey];
    if (label && label.length) {
      labelled++;
    }
    if (doc.modified && doc.modified < staleBefore) {
      stale++;
      staleBytes += doc.size || 0;
    }
    // Same name and same size anywhere in the site: very likely a copy.
    if (doc.size) {
      const key = `${doc.name.toLowerCase()}|${doc.size}`;
      if (seen[key]) {
        duplicates++;
        duplicateBytes += doc.size;
      } else {
        seen[key] = true;
      }
    }
  });

  return {
    ...base,
    documents: Math.max(libraryItemCount, documents.length),
    scanned: documents.length,
    sampled,
    fieldFilled,
    missingColumns,
    complete,
    labelled,
    stale,
    staleBytes,
    duplicates,
    duplicateBytes,
    ownership: 'unknown'
  };
}

export function aggregate(sites: ISiteMetrics[], requiredFields: string[]): IAggregate {
  const total: IAggregate = {
    sites: sites.length,
    scannedSites: 0,
    documents: 0,
    scanned: 0,
    complete: 0,
    labelled: 0,
    stale: 0,
    staleBytes: 0,
    duplicates: 0,
    duplicateBytes: 0,
    storageBytes: 0,
    storageKnown: false,
    ownersOk: 0,
    ownersKnown: 0,
    fieldFilled: {},
    missingColumnSites: {}
  };
  requiredFields.forEach((f) => {
    total.fieldFilled[f] = 0;
    total.missingColumnSites[f] = 0;
  });
  sites.forEach((s) => {
    if (s.error) {
      return;
    }
    total.scannedSites++;
    total.documents += s.documents;
    total.scanned += s.scanned;
    total.complete += s.complete;
    total.labelled += s.labelled;
    total.stale += s.stale;
    total.staleBytes += s.staleBytes;
    total.duplicates += s.duplicates;
    total.duplicateBytes += s.duplicateBytes;
    if (typeof s.storageBytes === 'number') {
      total.storageBytes += s.storageBytes;
      total.storageKnown = true;
    }
    const ok = ownershipOk(s.ownership);
    if (ok !== undefined) {
      total.ownersKnown++;
      total.ownersOk += ok ? 1 : 0;
    }
    requiredFields.forEach((f) => {
      total.fieldFilled[f] += s.fieldFilled[f] || 0;
      if (s.missingColumns.indexOf(f) >= 0) {
        total.missingColumnSites[f]++;
      }
    });
  });
  return total;
}

export interface IScoreParts {
  metadata?: number;
  labelled?: number;
  stale?: number;
  ownership?: number;
  score?: number;
}

/** The headline percentages and the weighted score; parts without data are left out of the weighting. */
export function scoreOf(a: IAggregate): IScoreParts {
  const metadata = pct(a.complete, a.scanned);
  const labelled = pct(a.labelled, a.scanned);
  const stale = pct(a.stale, a.scanned);
  const ownership = pct(a.ownersOk, a.ownersKnown);
  const parts: [number | undefined, number][] = [
    [metadata, WEIGHTS.metadata],
    [labelled, WEIGHTS.labelled],
    [stale === undefined ? undefined : 100 - stale, WEIGHTS.fresh],
    [ownership, WEIGHTS.ownership]
  ];
  let sum = 0;
  let weight = 0;
  parts.forEach(([value, w]) => {
    if (value !== undefined) {
      sum += value * w;
      weight += w;
    }
  });
  // Without documents there is nothing to score: ownership alone would mislead.
  const score = a.scanned > 0 && weight ? sum / weight : undefined;
  return { metadata, labelled, stale, ownership, score };
}

export function siteScore(site: ISiteMetrics, requiredFields: string[]): IScoreParts {
  return scoreOf(aggregate([site], requiredFields));
}

export function formatBytes(bytes: number): string {
  if (!bytes) {
    return '0 B';
  }
  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.min(units.length - 1, Math.floor(Math.log(bytes) / Math.log(1024)));
  const value = bytes / Math.pow(1024, i);
  return `${value >= 100 || i === 0 ? Math.round(value) : value.toFixed(1)} ${units[i]}`;
}

export function formatPct(value: number | undefined): string {
  return value === undefined ? '–' : `${Math.round(value)}%`;
}
