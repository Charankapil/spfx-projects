import * as React from 'react';
import { Icon } from '@fluentui/react';

import { ISiteCollectionOverview } from '../../models/ISiteCollectionOverview';
import { formatBytes } from '../../services/formatBytes';
import styles from './Dashboard.module.scss';
import { formatCompact, formatShortDate } from './format';
import { useCountUp, useEntrance } from './motion';

interface IKpiProps {
  label: string;
  icon: string;
  value: number | undefined;
  previous: number | undefined;
  since: Date | undefined;
  format: (n: number) => string;
}

/**
 * Change against the previous scan, in neutral ink: more files or more
 * storage is not good or bad in itself, so the delta carries no status colour.
 */
function delta(current: number, previous: number | undefined, since: Date | undefined, format: (n: number) => string): string | undefined {
  if (previous === undefined || !since) {
    return undefined;
  }
  const diff = current - previous;
  const when = formatShortDate(since);
  if (diff === 0) {
    return `No change since ${when}`;
  }
  return `${diff > 0 ? '+' : '−'}${format(Math.abs(diff))} since ${when}`;
}

const Kpi: React.FC<IKpiProps> = ({ label, icon, value, previous, since, format }) => {
  const animated = useCountUp(value || 0);
  const entering = useEntrance(value);
  const shown = value === undefined ? 'n/a' : format(animated);
  const change = value === undefined ? undefined : delta(value, previous, since, format);
  return (
    <div className={entering ? `${styles.kpi} ${styles.enter}` : styles.kpi}>
      <div className={styles.kpiHead}>
        <span className={styles.kpiIcon}>
          <Icon iconName={icon} />
        </span>
        <span className={styles.kpiLabel}>{label}</span>
      </div>
      <div className={styles.kpiValue} title={value === undefined ? shown : format(value)}>
        {shown}
      </div>
      {change && (
        <div className={styles.kpiDelta} title={change}>
          {change}
        </div>
      )}
    </div>
  );
};

export const KpiRow: React.FC<{ overview: ISiteCollectionOverview }> = ({ overview }) => {
  const prev = overview.previous;
  const since = prev ? prev.scanCompletedAt : undefined;

  return (
    <div className={styles.kpiRow}>
      <Kpi
        label="Files"
        icon="Page"
        value={overview.totalFiles}
        previous={prev ? prev.totalFiles : undefined}
        since={since}
        format={formatCompact}
      />
      <Kpi
        label="File types"
        icon="Tag"
        value={overview.totalFileTypeStats.length}
        previous={prev ? prev.distinctTypes : undefined}
        since={since}
        format={formatCompact}
      />
      <Kpi
        label="Libraries"
        icon="DocLibrary"
        value={overview.totalLibraries}
        previous={prev ? prev.totalLibraries : undefined}
        since={since}
        format={formatCompact}
      />
      <Kpi
        label="Sites (incl. root)"
        icon="SharepointLogo"
        value={overview.totalWebs}
        previous={prev ? prev.totalWebs : undefined}
        since={since}
        format={formatCompact}
      />
      <Kpi
        label="Site storage used"
        icon="Database"
        value={overview.storage.available ? overview.storage.usedBytes : undefined}
        previous={prev ? prev.storageBytes : undefined}
        since={since}
        format={formatBytes}
      />
    </div>
  );
};
