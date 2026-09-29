import * as React from 'react';
import { useMemo, useState } from 'react';
import { DefaultButton } from '@fluentui/react';

import { IScanResult } from '../../models/IScanResult';
import { describeAge, thresholdLabel } from '../../services/activity';
import { exportLargestFilesCsv } from '../../services/ExportService';
import { formatBytes } from '../../services/formatBytes';
import styles from './Dashboard.module.scss';
import { formatDay } from './format';

interface ILargestFilesProps {
  result: IScanResult;
  thresholdMonths: number;
}

const COLLAPSED_ROWS = 10;

/** The biggest inactive files: usually the quickest wins when freeing space. */
export const LargestFiles: React.FC<ILargestFilesProps> = ({ result, thresholdMonths }) => {
  const [showAll, setShowAll] = useState(false);
  const origin = useMemo(() => new URL(result.rootUrl).origin, [result]);
  const files = useMemo(
    () => result.largestOldFiles.filter((f) => f.ageMonths >= thresholdMonths),
    [result, thresholdMonths]
  );

  if (files.length === 0) {
    return null;
  }
  const visible = showAll ? files.slice(0, 50) : files.slice(0, COLLAPSED_ROWS);

  return (
    <section className={styles.card} aria-label="Largest inactive files">
      <div className={styles.cardHeader}>
        <div>
          <h3 className={styles.cardTitle}>Largest inactive files</h3>
          <div className={styles.cardSubtitle}>
            Files not modified for {thresholdLabel(thresholdMonths)} or more, biggest first.
          </div>
        </div>
        <DefaultButton
          text="Export list"
          iconProps={{ iconName: 'Download' }}
          onClick={() => exportLargestFilesCsv(result, thresholdMonths)}
        />
      </div>
      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th>File</th>
              <th className={styles.num}>Size</th>
              <th className={styles.num}>Last modified</th>
            </tr>
          </thead>
          <tbody>
            {visible.map((file) => {
              const modified = file.modified ? new Date(file.modified) : undefined;
              return (
                <tr key={file.serverRelativeUrl}>
                  <td>
                    <a
                      href={`${origin}${file.serverRelativeUrl}`}
                      target="_blank"
                      rel="noopener noreferrer"
                      className={styles.nameLink}
                    >
                      {file.name}
                    </a>
                    <div className={styles.subtle}>
                      {file.webTitle} › {file.libraryTitle}
                    </div>
                  </td>
                  <td className={styles.num}>
                    <strong>{formatBytes(file.bytes)}</strong>
                  </td>
                  <td className={styles.num}>
                    {modified && !isNaN(modified.getTime()) ? formatDay(modified) : ''}
                    <div className={styles.subtle}>{describeAge(file.ageMonths)}</div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {files.length > COLLAPSED_ROWS && (
        <DefaultButton
          className={styles.showAll}
          text={showAll ? 'Show fewer' : `Show top ${Math.min(files.length, 50)}`}
          onClick={() => setShowAll(!showAll)}
        />
      )}
    </section>
  );
};
