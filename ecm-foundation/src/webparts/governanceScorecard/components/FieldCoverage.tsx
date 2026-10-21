import * as React from 'react';

import { IAggregate } from '../models/IScorecard';
import { pct } from '../services/metrics';
import styles from './GovernanceScorecard.module.scss';
import { SERIES } from './vizTokens';

/** One series, one colour: coverage per required field, weakest first, value at the bar tip. */
export const FieldCoverage = (props: { total: IAggregate; requiredFields: string[] }): JSX.Element => {
  const rows = props.requiredFields
    .map((field) => ({ field, value: pct(props.total.fieldFilled[field] || 0, props.total.scanned), missingIn: props.total.missingColumnSites[field] || 0 }))
    .sort((a, b) => (a.value ?? -1) - (b.value ?? -1));
  return (
    <section className={styles.card} aria-labelledby="ecm-field-coverage">
      <h3 id="ecm-field-coverage" className={styles.cardTitle}>
        Required metadata, field by field
      </h3>
      <p className={styles.cardSub}>Share of scanned documents with each required field filled in.</p>
      <ul className={styles.bars}>
        {rows.map((r) => (
          <li key={r.field} className={styles.barRow}>
            <span className={styles.barLabel}>{r.field}</span>
            <span className={styles.barTrack} title={`${r.field}: ${r.value === undefined ? 'no documents' : `${Math.round(r.value)}%`}`}>
              <span className={styles.barFill} style={{ width: `${r.value || 0}%`, background: SERIES.score }} />
            </span>
            <span className={styles.barValue}>{r.value === undefined ? '–' : `${Math.round(r.value)}%`}</span>
            {r.missingIn > 0 && (
              <span className={styles.barNote}>
                Column missing in {r.missingIn} {r.missingIn === 1 ? 'site' : 'sites'}
              </span>
            )}
          </li>
        ))}
      </ul>
    </section>
  );
};
