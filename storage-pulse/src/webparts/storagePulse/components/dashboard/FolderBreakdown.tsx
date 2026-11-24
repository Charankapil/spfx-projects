import * as React from 'react';
import { useMemo, useState } from 'react';
import { DefaultButton, Icon, TooltipHost } from '@fluentui/react';
import * as strings from 'StoragePulseWebPartStrings';

import { IScanResult } from '../../models/IScanResult';
import { exportFoldersCsv } from '../../services/ExportService';
import { folderRows, inactiveFolders } from '../../services/folders';
import { formatBytes } from '../../services/formatBytes';
import { encodePath, safeHref } from '../../services/safeData';
import { describeAge, format, thresholdLabel } from '../text';
import styles from './Dashboard.module.scss';
import { formatCompact } from './format';

interface IFolderBreakdownProps {
  result: IScanResult;
  thresholdMonths: number;
}

const COLLAPSED_ROWS = 12;
const MAX_ROWS = 100;

/**
 * The biggest folders of large libraries, straight from SharePoint's storage
 * metrics: exact sizes, version history and the date of the last change in
 * each, so whole inactive folders stand out without reading their files.
 */
export const FolderBreakdown: React.FC<IFolderBreakdownProps> = ({ result, thresholdMonths }) => {
  const [showAll, setShowAll] = useState(false);
  const origin = useMemo(() => new URL(result.rootUrl).origin, [result]);
  const rows = useMemo(() => folderRows(result, thresholdMonths), [result, thresholdMonths]);
  const inactive = useMemo(() => inactiveFolders(rows), [rows]);
  const notes = useMemo(
    () => result.libraries.filter((l) => l.foldersNote && l.folders).map((l) => `${l.title}: ${l.foldersNote}`),
    [result]
  );

  if (rows.length === 0) {
    return null;
  }
  const period = thresholdLabel(thresholdMonths);
  const manyLibraries = new Set(rows.map((r) => r.libraryKey)).size > 1;
  const visible = rows.slice(0, showAll ? MAX_ROWS : COLLAPSED_ROWS);

  return (
    <section className={styles.card} aria-label={strings.FoldersTitle}>
      <div className={styles.cardHeader}>
        <div>
          <h3 className={styles.cardTitle}>{strings.FoldersTitle}</h3>
          <div className={styles.cardSubtitle}>{strings.FoldersSubtitle}</div>
        </div>
        <DefaultButton
          text={strings.ExportList}
          iconProps={{ iconName: 'Download' }}
          onClick={() => exportFoldersCsv(result, thresholdMonths)}
        />
      </div>
      {inactive.folders.length > 0 && (
        <div className={styles.note} role="note">
          <Icon iconName="Archive" className={styles.insightIcon} />
          <span>
            {format(strings.FoldersInactive, {
              count: inactive.folders.length,
              period,
              size: formatBytes(inactive.bytes),
              total: formatBytes(inactive.totalSize),
              files: formatCompact(inactive.files)
            })}
          </span>
        </div>
      )}
      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th>{strings.ColFolder}</th>
              <th className={styles.num}>{strings.ColFiles}</th>
              <th className={styles.num}>{strings.ColSize}</th>
              <th className={styles.num}>{strings.ColVersions}</th>
              <th className={styles.num}>{strings.ColLastChange}</th>
            </tr>
          </thead>
          <tbody>
            {visible.map((row) => {
              const href = safeHref(`${row.libraryUrl}/${encodePath(row.path)}`, origin);
              return (
                <tr key={row.key}>
                  <td>
                    <div className={styles.nameCell} style={{ paddingLeft: (row.depth - 1) * 12 }}>
                      {href ? (
                        <a href={href} target="_blank" rel="noopener noreferrer" className={styles.nameLink}>
                          {row.path}
                        </a>
                      ) : (
                        <span className={styles.nameText}>{row.path}</span>
                      )}
                      {row.inactive && <span className={styles.badge}>{strings.Dormant}</span>}
                      {row.childrenUnavailable && (
                        <TooltipHost content={strings.FolderNotOpened}>
                          <span className={styles.warnIcon} tabIndex={0} role="img" aria-label={strings.FolderNotOpened}>
                            <Icon iconName="Info" />
                          </span>
                        </TooltipHost>
                      )}
                    </div>
                    {manyLibraries && (
                      <div className={styles.subtle}>
                        {row.webTitle} › {row.libraryTitle}
                      </div>
                    )}
                  </td>
                  <td className={styles.num}>{formatCompact(row.files)}</td>
                  <td className={styles.num}>
                    <strong>{formatBytes(row.bytes)}</strong>
                  </td>
                  <td className={styles.num}>{formatBytes(row.versionBytes)}</td>
                  <td className={styles.num}>{row.ageMonths === undefined ? '–' : describeAge(row.ageMonths)}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {notes.length > 0 && <p className={styles.footnote}>{notes.join(' · ')}</p>}
      {rows.length > COLLAPSED_ROWS && (
        <DefaultButton
          className={styles.showAll}
          text={showAll ? strings.ShowFewer : format(strings.ShowTop, { count: Math.min(rows.length, MAX_ROWS) })}
          onClick={() => setShowAll(!showAll)}
        />
      )}
    </section>
  );
};
