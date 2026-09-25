import * as React from 'react';
import { useState } from 'react';

import { ISiteCollectionOverview } from '../../models/ISiteCollectionOverview';
import styles from './Dashboard.module.scss';
import { categoryOf } from './fileTypeCategories';
import { formatPercent } from './format';

const COLLAPSED_ROWS = 12;

/** Exact per-extension file counts - every type, most common first. */
export const TopFileTypes: React.FC<{ overview: ISiteCollectionOverview }> = ({ overview }) => {
  const [showAll, setShowAll] = useState(false);
  const stats = overview.totalFileTypeStats;
  const total = stats.reduce((sum, s) => sum + s.count, 0);
  const max = stats.reduce((m, s) => Math.max(m, s.count), 0);
  const rows = showAll ? stats : stats.slice(0, COLLAPSED_ROWS);

  return (
    <div className={styles.panel}>
      <h3 className={styles.panelTitle}>File types</h3>
      <p className={styles.panelCaption}>
        Number of files of each type across the site collection
        {overview.typeListCapped &&
          '. Some libraries only reported their 10 most common types, so rarer types there are not counted.'}
      </p>
      {stats.length === 0 ? (
        <div className={styles.empty}>No files found.</div>
      ) : (
        <>
          <ul className={styles.typeList}>
            {rows.map((s) => {
              const cat = categoryOf(s.extension);
              return (
                <li
                  key={s.extension}
                  className={styles.typeRow}
                  title={`.${s.extension} — ${cat.label}: ${s.count.toLocaleString()} files (${formatPercent(s.count, total)})`}
                >
                  <span className={styles.typeName}>.{s.extension}</span>
                  <span className={styles.barTrack}>
                    <span
                      className={styles.bar}
                      style={{ display: 'block', width: `${max ? (s.count / max) * 100 : 0}%`, background: cat.color }}
                    />
                  </span>
                  <span className={styles.typeValue}>
                    {s.count.toLocaleString()} &middot; {formatPercent(s.count, total)}
                  </span>
                </li>
              );
            })}
          </ul>
          {stats.length > COLLAPSED_ROWS && (
            <button type="button" className={styles.showAll} onClick={() => setShowAll(!showAll)}>
              {showAll ? 'Show top 12' : `Show all ${stats.length} types`}
            </button>
          )}
        </>
      )}
    </div>
  );
};
