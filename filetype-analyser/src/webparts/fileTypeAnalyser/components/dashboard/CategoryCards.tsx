import * as React from 'react';
import { Icon } from '@fluentui/react';

import { ISiteCollectionOverview } from '../../models/ISiteCollectionOverview';
import styles from './Dashboard.module.scss';
import { totalsByCategory } from './fileTypeCategories';
import { formatCompact, formatPercent } from './format';
import { useEntrance } from './motion';

/**
 * "What's in this site collection": one card per category with its share
 * of all files. Doubles as the colour key for the treemap below it.
 */
export const CategoryCards: React.FC<{ overview: ISiteCollectionOverview }> = ({ overview }) => {
  const entering = useEntrance(overview);
  const totals = totalsByCategory(overview.totalFileTypeStats);
  const all = totals.reduce((sum, t) => sum + t.count, 0);
  if (totals.length === 0) {
    return null;
  }

  return (
    <div className={styles.panel}>
      <h3 className={styles.panelTitle}>What&rsquo;s in this site collection</h3>
      <p className={styles.panelCaption}>Files grouped by kind; colours match the treemap</p>
      <div className={styles.categoryGrid}>
        {totals.map((t, i) => {
          const share = all ? t.count / all : 0;
          const topTypes = t.extensions
            .slice()
            .sort((a, b) => b.count - a.count)
            .slice(0, 3)
            .map((e) => `.${e.extension}`)
            .join(' · ');
          return (
            <div
              key={t.category.key}
              className={entering ? `${styles.categoryCard} ${styles.enter}` : styles.categoryCard}
              style={entering ? { animationDelay: `${Math.min(i, 8) * 40}ms` } : undefined}
              title={`${t.category.label}: ${t.count.toLocaleString()} files (${formatPercent(t.count, all)})`}
            >
              <div className={styles.categoryHead}>
                <span className={styles.categoryIcon} style={{ background: t.category.color, color: t.category.ink }}>
                  <Icon iconName={t.category.icon} />
                </span>
                <span className={styles.categoryName}>{t.category.label}</span>
              </div>
              <div className={styles.categoryCount}>
                {formatCompact(t.count)}
                <span className={styles.categoryShare}>{formatPercent(t.count, all)}</span>
              </div>
              <div className={styles.categoryTrack}>
                <div
                  className={styles.categoryFill}
                  style={{ width: `${Math.max(share * 100, share > 0 ? 2 : 0)}%`, background: t.category.color }}
                />
              </div>
              <div className={styles.categoryTypes}>{topTypes}</div>
            </div>
          );
        })}
      </div>
    </div>
  );
};
