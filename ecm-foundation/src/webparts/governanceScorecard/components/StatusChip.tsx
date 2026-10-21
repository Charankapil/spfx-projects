import * as React from 'react';

import { Status } from '../services/metrics';
import styles from './GovernanceScorecard.module.scss';
import { STATUS_COLOR } from './vizTokens';

export const STATUS_LABEL: { [s in Status]: string } = {
  good: 'On track',
  warning: 'Needs attention',
  critical: 'At risk'
};

/** Status never relies on colour alone: an icon shape and a label go with it. */
export const StatusIcon = (props: { status: Status }): JSX.Element => {
  const color = STATUS_COLOR[props.status];
  return (
    <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden="true">
      <circle cx="8" cy="8" r="8" fill={color} />
      {props.status === 'good' && <path d="M4.5 8.2l2.2 2.2 4.8-4.8" fill="none" stroke="#fff" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />}
      {props.status === 'warning' && <path d="M8 4.2v4.6M8 11.2v.4" fill="none" stroke="#1E2227" strokeWidth="1.8" strokeLinecap="round" />}
      {props.status === 'critical' && <path d="M5.5 5.5l5 5M10.5 5.5l-5 5" fill="none" stroke="#fff" strokeWidth="1.8" strokeLinecap="round" />}
    </svg>
  );
};

export const StatusChip = (props: { status?: Status }): JSX.Element =>
  props.status ? (
    <span className={styles.statusChip}>
      <StatusIcon status={props.status} />
      {STATUS_LABEL[props.status]}
    </span>
  ) : (
    <span className={styles.statusChipNone}>No data</span>
  );
