import * as React from 'react';
import { useMemo, useState } from 'react';

import { IAgeHistogram } from '../../models/IScanResult';
import { AGE_BANDS, bandTotals, thresholdLabel } from '../../services/activity';
import { formatBytes } from '../../services/formatBytes';
import styles from './Dashboard.module.scss';
import { formatCompact, formatPercent } from './format';

interface IAgeChartProps {
  histogram: IAgeHistogram;
  thresholdMonths: number;
}

/**
 * Storage by time since last modified. One bar per age band, coloured by
 * whether that band is active or inactive for the chosen threshold, so the
 * split in the summary can be seen building up across ages.
 */
export const AgeChart: React.FC<IAgeChartProps> = ({ histogram, thresholdMonths }) => {
  const [hovered, setHovered] = useState<number | undefined>(undefined);

  const bands = useMemo(() => {
    const rows = AGE_BANDS.map((band) => ({ band, ...bandTotals(histogram, band) }));
    const totalBytes = rows.reduce((sum, r) => sum + r.bytes, 0);
    const maxBytes = rows.reduce((max, r) => Math.max(max, r.bytes), 0);
    return { rows, totalBytes, maxBytes };
  }, [histogram]);

  return (
    <section className={styles.card} aria-label="Storage by time since last modified">
      <div className={styles.cardHeader}>
        <div>
          <h3 className={styles.cardTitle}>Storage by time since last modified</h3>
          <div className={styles.cardSubtitle}>
            Bars to the right of the line hold files not modified for {thresholdLabel(thresholdMonths)} or more.
          </div>
        </div>
        <div className={styles.legend}>
          <span className={styles.legendItem}>
            <span className={`${styles.swatch} ${styles.active}`} /> Active
          </span>
          <span className={styles.legendItem}>
            <span className={`${styles.swatch} ${styles.inactive}`} /> Inactive
          </span>
        </div>
      </div>

      <div className={styles.chart}>
        {bands.rows.map((row, i) => {
          const inactive = row.band.from >= thresholdMonths;
          const heightPct = bands.maxBytes ? (row.bytes / bands.maxBytes) * 100 : 0;
          const isBoundary = row.band.from === thresholdMonths;
          return (
            <div
              key={row.band.label}
              className={`${styles.column} ${isBoundary ? styles.boundary : ''}`}
              onMouseEnter={() => setHovered(i)}
              onMouseLeave={() => setHovered(undefined)}
              onFocus={() => setHovered(i)}
              onBlur={() => setHovered(undefined)}
              tabIndex={0}
              aria-label={`${row.band.label}: ${formatBytes(row.bytes)}, ${row.files.toLocaleString()} files, ${
                inactive ? 'inactive' : 'active'
              }`}
            >
              <div className={styles.barArea}>
                <span className={styles.barValue}>{row.bytes > 0 ? formatBytes(row.bytes) : ''}</span>
                <span
                  className={`${styles.bar} ${inactive ? styles.inactive : styles.active} ${
                    hovered !== undefined && hovered !== i ? styles.dimmed : ''
                  }`}
                  style={{ height: `${row.bytes > 0 ? Math.max(heightPct, 1.5) : 0}%` }}
                />
                {hovered === i && (
                  <div className={styles.tooltip} role="tooltip">
                    <div className={styles.tooltipTitle}>Last modified: {row.band.label}</div>
                    <div>
                      <strong>{formatBytes(row.bytes)}</strong> · {formatPercent(row.bytes, bands.totalBytes)} of file storage
                    </div>
                    <div>{formatCompact(row.files)} files</div>
                    <div className={styles.tooltipMuted}>{inactive ? 'Inactive' : 'Active'}</div>
                  </div>
                )}
              </div>
              <span className={styles.axisLabel}>{row.band.label}</span>
            </div>
          );
        })}
      </div>
    </section>
  );
};
