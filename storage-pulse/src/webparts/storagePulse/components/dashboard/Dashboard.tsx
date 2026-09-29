import * as React from 'react';
import { useMemo } from 'react';
import * as strings from 'StoragePulseWebPartStrings';

import { IScanResult } from '../../models/IScanResult';
import { splitByThreshold, totalHistogram } from '../../services/activity';
import { AgeChart } from './AgeChart';
import styles from './Dashboard.module.scss';
import { FileTypes } from './FileTypes';
import { LargestFiles } from './LargestFiles';
import { LocationTable } from './LocationTable';
import { SplitSummary } from './SplitSummary';

export interface IDashboardProps {
  result: IScanResult;
  thresholdMonths: number;
}

export const Dashboard: React.FC<IDashboardProps> = ({ result, thresholdMonths }) => {
  const histogram = useMemo(() => totalHistogram(result.libraries), [result]);
  const split = useMemo(() => splitByThreshold(histogram, thresholdMonths), [histogram, thresholdMonths]);

  return (
    <div className={styles.dashboard}>
      <SplitSummary result={result} split={split} thresholdMonths={thresholdMonths} />
      <AgeChart histogram={histogram} thresholdMonths={thresholdMonths} />
      <FileTypes result={result} thresholdMonths={thresholdMonths} />
      <LocationTable result={result} thresholdMonths={thresholdMonths} />
      <LargestFiles result={result} thresholdMonths={thresholdMonths} />
      <p className={styles.footnote}>{strings.Footnote}</p>
    </div>
  );
};
