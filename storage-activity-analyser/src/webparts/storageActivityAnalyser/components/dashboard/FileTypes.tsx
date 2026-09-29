import * as React from 'react';
import { useMemo, useState } from 'react';
import { DefaultButton, Icon, Pivot, PivotItem } from '@fluentui/react';

import { IScanResult } from '../../models/IScanResult';
import { IFileTypeSplit, splitFileType, thresholdLabel, totalFileTypes } from '../../services/activity';
import { exportFileTypesCsv } from '../../services/ExportService';
import { formatBytes } from '../../services/formatBytes';
import styles from './Dashboard.module.scss';
import { categoryOf, CATEGORIES, IFileCategory, OTHER_CATEGORY } from './fileTypeCategories';
import { formatCompact, formatPercent } from './format';

interface IFileTypesProps {
  result: IScanResult;
  thresholdMonths: number;
}

interface IRow {
  key: string;
  label: string;
  icon: string;
  detail: string;
  split: IFileTypeSplit;
}

type SortKey = 'totalBytes' | 'inactiveBytes' | 'totalFiles';

const COLLAPSED_ROWS = 12;

function typeLabel(extension: string): string {
  return extension.charAt(0) === '(' ? extension : `.${extension}`;
}

function sumSplits(extension: string, splits: IFileTypeSplit[]): IFileTypeSplit {
  const total: IFileTypeSplit = {
    extension,
    activeFiles: 0,
    activeBytes: 0,
    inactiveFiles: 0,
    inactiveBytes: 0,
    totalFiles: 0,
    totalBytes: 0
  };
  for (const s of splits) {
    total.activeFiles += s.activeFiles;
    total.activeBytes += s.activeBytes;
    total.inactiveFiles += s.inactiveFiles;
    total.inactiveBytes += s.inactiveBytes;
    total.totalFiles += s.totalFiles;
    total.totalBytes += s.totalBytes;
  }
  return total;
}

/** What kinds of file take the storage, split into active and inactive. */
export const FileTypes: React.FC<IFileTypesProps> = ({ result, thresholdMonths }) => {
  const [view, setView] = useState<'type' | 'category'>('type');
  const [sortKey, setSortKey] = useState<SortKey>('totalBytes');
  const [showAll, setShowAll] = useState(false);

  const splits = useMemo(() => {
    const stats = totalFileTypes(result.libraries);
    return stats ? stats.map((s) => splitFileType(s, thresholdMonths)) : undefined;
  }, [result, thresholdMonths]);

  const rows: IRow[] = useMemo(() => {
    if (!splits) {
      return [];
    }
    if (view === 'type') {
      return splits.map((s) => {
        const category = categoryOf(s.extension);
        return { key: s.extension, label: typeLabel(s.extension), icon: category.icon, detail: category.label, split: s };
      });
    }
    const groups: { [key: string]: IFileTypeSplit[] } = {};
    for (const s of splits) {
      const key = categoryOf(s.extension).key;
      (groups[key] = groups[key] || []).push(s);
    }
    return CATEGORIES.concat([OTHER_CATEGORY])
      .filter((c: IFileCategory) => groups[c.key])
      .map((c: IFileCategory) => {
        const members = groups[c.key].slice().sort((a, b) => b.totalBytes - a.totalBytes);
        const top = members.slice(0, 4).map((m) => typeLabel(m.extension)).join(', ');
        return {
          key: c.key,
          label: c.label,
          icon: c.icon,
          detail: members.length > 4 ? `${top} and ${members.length - 4} more` : top,
          split: sumSplits(c.key, members)
        };
      });
  }, [splits, view]);

  const sorted = useMemo(() => rows.slice().sort((a, b) => b.split[sortKey] - a.split[sortKey]), [rows, sortKey]);

  if (!splits) {
    return (
      <section className={styles.card} aria-label="Storage by file type">
        <h3 className={styles.cardTitle}>Storage by file type</h3>
        <div className={styles.cardSubtitle}>This scan was saved by an earlier version. Run a new scan to see file types.</div>
      </section>
    );
  }

  const totalBytes = rows.reduce((sum, r) => sum + r.split.totalBytes, 0);
  const totalInactive = rows.reduce((sum, r) => sum + r.split.inactiveBytes, 0);
  const maxBytes = sorted.reduce((max, r) => Math.max(max, r.split.totalBytes), 0);
  const biggestInactive = rows.slice().sort((a, b) => b.split.inactiveBytes - a.split.inactiveBytes)[0];
  const visible = showAll ? sorted : sorted.slice(0, COLLAPSED_ROWS);
  const period = thresholdLabel(thresholdMonths);

  const header = (key: SortKey, label: string): JSX.Element => (
    <th className={styles.num} aria-sort={sortKey === key ? 'descending' : 'none'}>
      <button className={styles.sortButton} onClick={() => setSortKey(key)}>
        {label}
        {sortKey === key && <Icon iconName="SortDown" className={styles.sortIcon} />}
      </button>
    </th>
  );

  return (
    <section className={styles.card} aria-label="Storage by file type">
      <div className={styles.cardHeader}>
        <div>
          <h3 className={styles.cardTitle}>Storage by file type</h3>
          <div className={styles.cardSubtitle}>
            Current file versions. The orange part of each bar is files not modified for {period} or more.
          </div>
          {biggestInactive && biggestInactive.split.inactiveBytes > 0 && (
            <div className={styles.highlight}>
              <Icon iconName="Lightbulb" className={styles.insightIcon} />
              <span>
                <strong>{biggestInactive.label}</strong> {view === 'type' ? 'files hold' : 'holds'} the most inactive
                storage: {formatBytes(biggestInactive.split.inactiveBytes)} (
                {formatPercent(biggestInactive.split.inactiveBytes, totalInactive)} of all inactive storage).
              </span>
            </div>
          )}
        </div>
        <div className={styles.cardActions}>
          <Pivot
            selectedKey={view}
            onLinkClick={(item) => {
              if (item) {
                setView(item.props.itemKey === 'category' ? 'category' : 'type');
                setShowAll(false);
              }
            }}
            headersOnly
          >
            <PivotItem headerText="By file type" itemKey="type" />
            <PivotItem headerText="By category" itemKey="category" />
          </Pivot>
          <DefaultButton
            text="Export"
            iconProps={{ iconName: 'Download' }}
            onClick={() => exportFileTypesCsv(result, thresholdMonths)}
          />
        </div>
      </div>

      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th>{view === 'type' ? 'File type' : 'Category'}</th>
              <th className={styles.barHeader}>
                <span className={styles.legend}>
                  <span className={styles.legendItem}>
                    <span className={`${styles.swatch} ${styles.inactive}`} /> Inactive
                  </span>
                  <span className={styles.legendItem}>
                    <span className={`${styles.swatch} ${styles.active}`} /> Active
                  </span>
                </span>
              </th>
              {header('totalBytes', 'Size')}
              {header('inactiveBytes', 'Inactive')}
              <th className={styles.num}>Share of storage</th>
              {header('totalFiles', 'Files')}
            </tr>
          </thead>
          <tbody>
            {visible.map((row) => (
              <tr key={row.key}>
                <td>
                  <div className={styles.nameCell}>
                    <Icon iconName={row.icon} className={styles.typeIcon} />
                    <span className={styles.typeName}>{row.label}</span>
                  </div>
                  <div className={styles.subtle}>{row.detail}</div>
                </td>
                <td className={styles.barCell}>
                  <span
                    className={styles.typeBar}
                    style={{ width: `${maxBytes ? Math.max(1, (row.split.totalBytes / maxBytes) * 100) : 0}%` }}
                    title={`Inactive ${formatBytes(row.split.inactiveBytes)} · active ${formatBytes(row.split.activeBytes)}`}
                  >
                    {row.split.inactiveBytes > 0 && (
                      <span className={`${styles.miniSegment} ${styles.inactive}`} style={{ flexGrow: row.split.inactiveBytes }} />
                    )}
                    {row.split.activeBytes > 0 && (
                      <span className={`${styles.miniSegment} ${styles.active}`} style={{ flexGrow: row.split.activeBytes }} />
                    )}
                  </span>
                </td>
                <td className={styles.num}>{formatBytes(row.split.totalBytes)}</td>
                <td className={styles.num}>
                  <strong>{formatBytes(row.split.inactiveBytes)}</strong>
                  <div className={styles.subtle}>{formatPercent(row.split.inactiveBytes, row.split.totalBytes)} of type</div>
                </td>
                <td className={styles.num}>{formatPercent(row.split.totalBytes, totalBytes)}</td>
                <td className={styles.num}>{formatCompact(row.split.totalFiles)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {sorted.length > COLLAPSED_ROWS && (
        <DefaultButton
          className={styles.showAll}
          text={showAll ? 'Show fewer' : `Show all ${sorted.length}`}
          onClick={() => setShowAll(!showAll)}
        />
      )}
    </section>
  );
};
