import * as React from 'react';
import { useMemo, useState } from 'react';
import { DefaultButton, Icon, Pivot, PivotItem, TooltipHost } from '@fluentui/react';
import * as strings from 'StoragePulseWebPartStrings';

import { IAgeHistogram, IScanResult } from '../../models/IScanResult';
import { addToHistogram, emptyHistogram, newestFileAge, splitByThreshold } from '../../services/activity';
import { formatBytes } from '../../services/formatBytes';
import { safeHref } from '../../services/safeData';
import { describeAge, format, thresholdLabel } from '../text';
import styles from './Dashboard.module.scss';
import { formatCompact, formatPercent } from './format';

interface ILocationTableProps {
  result: IScanResult;
  thresholdMonths: number;
}

interface IRow {
  key: string;
  title: string;
  subtitle: string;
  url: string;
  files: number;
  totalBytes: number;
  activeBytes: number;
  inactiveBytes: number;
  inactiveShare: number;
  newest?: number;
  versionBytes?: number;
  error?: string;
}

type SortKey = 'title' | 'files' | 'totalBytes' | 'inactiveBytes' | 'inactiveShare' | 'newest';

const COLLAPSED_ROWS = 12;

function buildRow(
  key: string,
  title: string,
  subtitle: string,
  url: string,
  histogram: IAgeHistogram,
  thresholdMonths: number,
  versionBytes: number | undefined,
  error: string | undefined
): IRow {
  const split = splitByThreshold(histogram, thresholdMonths);
  return {
    key,
    title,
    subtitle,
    url,
    files: split.totalFiles,
    totalBytes: split.totalBytes,
    activeBytes: split.activeBytes,
    inactiveBytes: split.inactiveBytes,
    inactiveShare: split.totalBytes ? split.inactiveBytes / split.totalBytes : 0,
    newest: newestFileAge(histogram),
    versionBytes,
    error
  };
}

export const LocationTable: React.FC<ILocationTableProps> = ({ result, thresholdMonths }) => {
  const [view, setView] = useState<'library' | 'site'>('library');
  const [sortKey, setSortKey] = useState<SortKey>('inactiveBytes');
  const [descending, setDescending] = useState(true);
  const [showAll, setShowAll] = useState(false);
  const hasSites = result.webs.length > 1;
  const origin = useMemo(() => new URL(result.rootUrl).origin, [result]);

  const rows = useMemo(() => {
    if (view === 'library') {
      return result.libraries.map((lib) =>
        buildRow(
          `${lib.webUrl}|${lib.id}`,
          lib.title,
          lib.webTitle,
          lib.url,
          lib.histogram,
          thresholdMonths,
          lib.metrics ? Math.max(0, lib.metrics.totalSize - lib.metrics.fileStreamSize) : undefined,
          lib.error
            ? lib.partial
              ? format(strings.PartlyRead, { error: lib.error })
              : lib.error
            : lib.unreadItems
            ? format(strings.HiddenItems, {
                hidden: lib.unreadItems.toLocaleString(),
                total: lib.itemCount.toLocaleString()
              })
            : undefined
        )
      );
    }
    return result.webs.map((web) => {
      const libs = result.libraries.filter((l) => l.webUrl === web.url);
      const histogram = emptyHistogram();
      let versionBytes: number | undefined;
      for (const lib of libs) {
        addToHistogram(histogram, lib.histogram);
        if (lib.metrics) {
          versionBytes = (versionBytes || 0) + Math.max(0, lib.metrics.totalSize - lib.metrics.fileStreamSize);
        }
      }
      const errors = [
        web.error,
        ...libs.filter((l) => l.error).map((l) => `${l.title}: ${l.error}`),
        ...libs
          .filter((l) => !l.error && l.unreadItems)
          .map((l) => format(strings.HiddenItemsShort, { title: l.title, count: (l.unreadItems || 0).toLocaleString() }))
      ].filter(Boolean);
      return buildRow(
        web.url,
        web.title,
        libs.length === 1 ? strings.OneLibrary : format(strings.LibrariesCount, { count: libs.length }),
        web.url,
        histogram,
        thresholdMonths,
        versionBytes,
        errors.length ? errors.join(' | ') : undefined
      );
    });
  }, [result, thresholdMonths, view]);

  const sorted = useMemo(() => {
    const copy = rows.slice();
    copy.sort((a, b) => {
      let cmp: number;
      if (sortKey === 'title') {
        cmp = a.title.localeCompare(b.title);
      } else if (sortKey === 'newest') {
        // Libraries with no files sort as the oldest.
        cmp = (a.newest === undefined ? Infinity : a.newest) - (b.newest === undefined ? Infinity : b.newest);
      } else {
        cmp = a[sortKey] - b[sortKey];
      }
      return descending ? -cmp : cmp;
    });
    return copy;
  }, [rows, sortKey, descending]);

  const hasVersions = rows.some((r) => r.versionBytes !== undefined);
  const visible = showAll ? sorted : sorted.slice(0, COLLAPSED_ROWS);

  const header = (key: SortKey, label: string, numeric = true): JSX.Element => {
    const active = sortKey === key;
    return (
      <th
        className={numeric ? styles.num : undefined}
        aria-sort={active ? (descending ? 'descending' : 'ascending') : 'none'}
      >
        <button
          className={styles.sortButton}
          onClick={() => {
            if (active) {
              setDescending(!descending);
            } else {
              setSortKey(key);
              setDescending(key !== 'title');
            }
          }}
        >
          {label}
          {active && <Icon iconName={descending ? 'SortDown' : 'SortUp'} className={styles.sortIcon} />}
        </button>
      </th>
    );
  };

  return (
    <section className={styles.card} aria-label={strings.WhereTitle}>
      <div className={styles.cardHeader}>
        <div>
          <h3 className={styles.cardTitle}>{strings.WhereTitle}</h3>
          <div className={styles.cardSubtitle}>{format(strings.WhereSubtitle, { period: thresholdLabel(thresholdMonths) })}</div>
        </div>
        {hasSites && (
          <Pivot
            selectedKey={view}
            onLinkClick={(item) => item && setView(item.props.itemKey === 'site' ? 'site' : 'library')}
            headersOnly
          >
            <PivotItem headerText={strings.ByLibrary} itemKey="library" />
            <PivotItem headerText={strings.BySite} itemKey="site" />
          </Pivot>
        )}
      </div>

      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <thead>
            <tr>
              {header('title', view === 'library' ? strings.ColLibrary : strings.ColSite, false)}
              {header('files', strings.ColFiles)}
              {header('totalBytes', strings.ColSize)}
              {header('inactiveBytes', strings.ColInactive)}
              {header('inactiveShare', strings.ColInactiveShare)}
              {header('newest', strings.ColLastChange)}
              {hasVersions && <th className={styles.num}>{strings.ColVersions}</th>}
            </tr>
          </thead>
          <tbody>
            {visible.map((row) => {
              const dormant = row.newest !== undefined && row.newest >= thresholdMonths;
              const href = safeHref(row.url, origin);
              return (
                <tr key={row.key}>
                  <td>
                    <div className={styles.nameCell}>
                      {href ? (
                        <a href={href} target="_blank" rel="noopener noreferrer" className={styles.nameLink}>
                          {row.title}
                        </a>
                      ) : (
                        <span className={styles.nameText}>{row.title}</span>
                      )}
                      {dormant && <span className={styles.badge}>{strings.Dormant}</span>}
                      {row.error && (
                        <TooltipHost content={row.error}>
                          <span className={styles.warnIcon} tabIndex={0} role="img" aria-label={row.error}>
                            <Icon iconName="Warning" />
                          </span>
                        </TooltipHost>
                      )}
                    </div>
                    <div className={styles.subtle}>{row.subtitle}</div>
                  </td>
                  <td className={styles.num}>{formatCompact(row.files)}</td>
                  <td className={styles.num}>{formatBytes(row.totalBytes)}</td>
                  <td className={styles.num}>
                    <strong>{formatBytes(row.inactiveBytes)}</strong>
                  </td>
                  <td className={styles.num}>
                    <div className={styles.shareCell}>
                      <span className={styles.miniBar} aria-hidden="true">
                        {row.inactiveBytes > 0 && (
                          <span className={`${styles.miniSegment} ${styles.inactive}`} style={{ flexGrow: row.inactiveBytes }} />
                        )}
                        {row.activeBytes > 0 && (
                          <span className={`${styles.miniSegment} ${styles.active}`} style={{ flexGrow: row.activeBytes }} />
                        )}
                      </span>
                      <span className={styles.shareValue}>
                        {row.totalBytes ? formatPercent(row.inactiveBytes, row.totalBytes) : '–'}
                      </span>
                    </div>
                  </td>
                  <td className={styles.num}>{row.newest === undefined ? strings.NoFiles : describeAge(row.newest)}</td>
                  {hasVersions && (
                    <td className={styles.num}>{row.versionBytes === undefined ? '–' : formatBytes(row.versionBytes)}</td>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {sorted.length > COLLAPSED_ROWS && (
        <DefaultButton
          className={styles.showAll}
          text={showAll ? strings.ShowFewer : format(strings.ShowAll, { count: sorted.length })}
          onClick={() => setShowAll(!showAll)}
        />
      )}
    </section>
  );
};
