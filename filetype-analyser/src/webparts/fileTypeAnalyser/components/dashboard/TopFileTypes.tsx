import * as React from 'react';

import { ISiteCollectionOverview } from '../../models/ISiteCollectionOverview';
import styles from './Dashboard.module.scss';
import { categoryOf, OTHER_CATEGORY } from './fileTypeCategories';
import { formatPercent } from './format';
import { formatBytes } from '../../services/formatBytes';

const TOP_N = 12;

interface IRow {
  key: string;
  label: string;
  count: number;
  color: string;
  hint: string;
}

/** Exact per-extension counts - the readable companion to the treemap. */
export const TopFileTypes: React.FC<{ overview: ISiteCollectionOverview }> = ({ overview }) => {
  const stats = overview.totalFileTypeStats;
  const total = stats.reduce((sum, s) => sum + s.count, 0);

  const rows: IRow[] = stats.slice(0, TOP_N).map((s) => {
    const cat = categoryOf(s.extension);
    const size = typeof s.estimatedBytes === 'number' ? `, ~${formatBytes(s.estimatedBytes)} estimated` : '';
    return { key: s.extension, label: `.${s.extension}`, count: s.count, color: cat.color, hint: `${cat.label}${size}` };
  });
  const rest = stats.slice(TOP_N);
  if (rest.length > 0) {
    rows.push({
      key: '__other',
      label: `${rest.length} more`,
      count: rest.reduce((sum, s) => sum + s.count, 0),
      color: OTHER_CATEGORY.color,
      hint: rest.map((s) => `.${s.extension}`).join(', ')
    });
  }
  const max = rows.reduce((m, r) => Math.max(m, r.count), 0);

  return (
    <div className={styles.panel}>
      <h3 className={styles.panelTitle}>File types</h3>
      <p className={styles.panelCaption}>Number of files across the site collection</p>
      {rows.length === 0 ? (
        <div className={styles.empty}>No files found.</div>
      ) : (
        <ul className={styles.typeList}>
          {rows.map((r) => (
            <li
              key={r.key}
              className={styles.typeRow}
              title={`${r.label} — ${r.hint}: ${r.count.toLocaleString()} files (${formatPercent(r.count, total)})`}
            >
              <span className={styles.typeName}>{r.label}</span>
              <span className={styles.barTrack}>
                <span
                  className={styles.bar}
                  style={{ display: 'block', width: `${max ? (r.count / max) * 100 : 0}%`, background: r.color }}
                />
              </span>
              <span className={styles.typeValue}>
                {r.count.toLocaleString()} &middot; {formatPercent(r.count, total)}
              </span>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
};
