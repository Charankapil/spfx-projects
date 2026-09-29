import * as React from 'react';
import { useMemo } from 'react';
import { Icon } from '@fluentui/react';
import * as strings from 'StoragePulseWebPartStrings';

import { IScanResult } from '../../models/IScanResult';
import { splitByThreshold, totalHistogram } from '../../services/activity';
import { formatBytes } from '../../services/formatBytes';
import { format, thresholdLabel } from '../text';
import { AgeChart } from './AgeChart';
import styles from './Dashboard.module.scss';
import { FileTypes } from './FileTypes';
import { LargestFiles } from './LargestFiles';
import { LocationTable } from './LocationTable';
import { ScanIssues } from './ScanIssues';
import { SplitSummary } from './SplitSummary';

export interface IDashboardProps {
  result: IScanResult;
  thresholdMonths: number;
}

export const Dashboard: React.FC<IDashboardProps> = ({ result, thresholdMonths }) => {
  const histogram = useMemo(() => totalHistogram(result.libraries), [result]);
  const split = useMemo(() => splitByThreshold(histogram, thresholdMonths), [histogram, thresholdMonths]);
  const quick = useMemo(() => {
    const measured = result.libraries.filter((l) => l.measuredAsWhole);
    return { count: measured.length, bytes: measured.reduce((sum, l) => sum + l.bytes, 0) };
  }, [result]);

  return (
    <div className={styles.dashboard}>
      <SplitSummary result={result} split={split} thresholdMonths={thresholdMonths} />
      {quick.count > 0 && result.quickAfterMonths !== undefined && (
        <div className={styles.note} role="note">
          <Icon iconName="Info" className={styles.insightIcon} />
          <span>
            {format(strings.QuickNote, {
              count: quick.count,
              period: thresholdLabel(result.quickAfterMonths),
              size: formatBytes(quick.bytes)
            })}
          </span>
        </div>
      )}
      <ScanIssues result={result} />
      <AgeChart histogram={histogram} thresholdMonths={thresholdMonths} />
      <FileTypes result={result} thresholdMonths={thresholdMonths} />
      <LocationTable result={result} thresholdMonths={thresholdMonths} />
      <LargestFiles result={result} thresholdMonths={thresholdMonths} />
      {result.paging && (
        <p className={styles.footnote}>{format(strings.PagingNote, { count: result.paging.libraries })}</p>
      )}
      <p className={styles.footnote}>{strings.Footnote}</p>
    </div>
  );
};
