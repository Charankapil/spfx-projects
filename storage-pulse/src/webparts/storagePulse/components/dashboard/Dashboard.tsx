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
import { FolderBreakdown } from './FolderBreakdown';
import { LargestFiles } from './LargestFiles';
import { LocationTable } from './LocationTable';
import { ScanIssues } from './ScanIssues';
import { SplitSummary } from './SplitSummary';

export interface IDashboardProps {
  result: IScanResult;
  thresholdMonths: number;
  /** Where the folder breakdowns come from: the whole result, including libraries not shown yet (default `result`). */
  foldersFrom?: IScanResult;
}

export const Dashboard: React.FC<IDashboardProps> = ({ result, thresholdMonths, foldersFrom }) => {
  const histogram = useMemo(() => totalHistogram(result.libraries), [result]);
  const split = useMemo(() => splitByThreshold(histogram, thresholdMonths), [histogram, thresholdMonths]);
  const quick = useMemo(() => {
    const measured = result.libraries.filter((l) => l.measuredAsWhole);
    return { count: measured.length, bytes: measured.reduce((sum, l) => sum + l.bytes, 0) };
  }, [result]);
  const estimated = useMemo(() => {
    const libs = result.libraries.filter((l) => l.fromEstimate);
    const coverages = libs.map((l) => (l.estimate ? l.estimate.indexCoverage : undefined)).filter((c): c is number => c !== undefined);
    return {
      count: libs.length,
      files: libs.reduce((sum, l) => sum + l.files, 0),
      lowestCoverage: coverages.length ? Math.min(...coverages) : undefined
    };
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
      {estimated.count > 0 && (
        <div className={styles.note} role="note">
          <Icon iconName="Lightning" className={styles.insightIcon} />
          <span>
            {format(strings.EstimateNote, { count: estimated.count, files: estimated.files.toLocaleString() })}
            {estimated.lowestCoverage !== undefined &&
              estimated.lowestCoverage < 0.95 &&
              ` ${format(strings.EstimateIndexGap, { percent: `${Math.round(estimated.lowestCoverage * 100)}%` })}`}
          </span>
        </div>
      )}
      <ScanIssues result={result} />
      <AgeChart histogram={histogram} thresholdMonths={thresholdMonths} />
      <FileTypes result={result} thresholdMonths={thresholdMonths} />
      <LocationTable result={result} thresholdMonths={thresholdMonths} />
      <FolderBreakdown result={foldersFrom || result} thresholdMonths={thresholdMonths} />
      <LargestFiles result={result} thresholdMonths={thresholdMonths} />
      {result.paging && (
        <p className={styles.footnote}>{format(strings.PagingNote, { count: result.paging.libraries })}</p>
      )}
      <p className={styles.footnote}>{strings.Footnote}</p>
    </div>
  );
};
