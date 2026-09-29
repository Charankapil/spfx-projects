import * as React from 'react';
import { useMemo } from 'react';

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
      <p className={styles.footnote}>
        Sizes are the current version of each file, from the libraries the person who ran the scan can open.
        A file counts as active when it was modified within the chosen period, measured from the scan date.
        Site storage used also includes version history, the recycle bins and list data, so it is larger than
        the file total.
      </p>
    </div>
  );
};
