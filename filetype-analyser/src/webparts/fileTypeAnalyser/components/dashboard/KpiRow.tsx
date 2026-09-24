import * as React from 'react';

import { ISiteCollectionOverview } from '../../models/ISiteCollectionOverview';
import { formatBytes } from '../../services/formatBytes';
import styles from './Dashboard.module.scss';
import { formatCompact, formatShortDate } from './format';

interface IKpi {
  label: string;
  value: string;
  delta?: string;
}

/**
 * Change against the previous scan, in neutral ink: more files or more
 * storage is not good or bad in itself, so the delta carries no status colour.
 */
function delta(
  current: number,
  previous: number | undefined,
  since: Date | undefined,
  format: (n: number) => string
): string | undefined {
  if (previous === undefined || !since) {
    return undefined;
  }
  const diff = current - previous;
  const when = formatShortDate(since);
  if (diff === 0) {
    return `No change since ${when}`;
  }
  const sign = diff > 0 ? '+' : '−';
  return `${sign}${format(Math.abs(diff))} since ${when}`;
}

export const KpiRow: React.FC<{ overview: ISiteCollectionOverview }> = ({ overview }) => {
  const prev = overview.previous;
  const since = prev ? prev.scanCompletedAt : undefined;

  const kpis: IKpi[] = [
    {
      label: 'Site storage used',
      value: overview.storage.available ? formatBytes(overview.storage.usedBytes) : 'n/a',
      delta: overview.storage.available
        ? delta(overview.storage.usedBytes, prev ? prev.storageBytes : undefined, since, formatBytes)
        : undefined
    },
    {
      label: 'Files',
      value: formatCompact(overview.totalFiles),
      delta: delta(overview.totalFiles, prev ? prev.totalFiles : undefined, since, formatCompact)
    },
    {
      label: 'Libraries',
      value: formatCompact(overview.totalLibraries),
      delta: delta(overview.totalLibraries, prev ? prev.totalLibraries : undefined, since, formatCompact)
    },
    {
      label: 'Sites (incl. root)',
      value: formatCompact(overview.totalWebs),
      delta: delta(overview.totalWebs, prev ? prev.totalWebs : undefined, since, formatCompact)
    },
    {
      label: 'File types',
      value: formatCompact(overview.totalFileTypeStats.length),
      delta: delta(
        overview.totalFileTypeStats.length,
        prev ? prev.distinctTypes : undefined,
        since,
        formatCompact
      )
    }
  ];

  return (
    <div className={styles.kpiRow}>
      {kpis.map((k) => (
        <div key={k.label} className={styles.kpi}>
          <div className={styles.kpiLabel}>{k.label}</div>
          <div className={styles.kpiValue} title={k.value}>
            {k.value}
          </div>
          {k.delta && (
            <div className={styles.kpiDelta} title={k.delta}>
              {k.delta}
            </div>
          )}
        </div>
      ))}
    </div>
  );
};
