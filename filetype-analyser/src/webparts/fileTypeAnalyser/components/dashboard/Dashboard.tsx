import * as React from 'react';

import { ISiteCollectionOverview } from '../../models/ISiteCollectionOverview';
import styles from './Dashboard.module.scss';
import { KpiRow } from './KpiRow';
import { TopFileTypes } from './TopFileTypes';
import { Treemap } from './Treemap';

export const Dashboard: React.FC<{ overview: ISiteCollectionOverview }> = ({ overview }) => (
  <div className={styles.dashboard}>
    <KpiRow overview={overview} />
    <div className={styles.panels}>
      <Treemap overview={overview} />
      <TopFileTypes overview={overview} />
    </div>
  </div>
);
