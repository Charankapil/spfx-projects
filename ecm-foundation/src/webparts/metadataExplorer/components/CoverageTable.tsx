import * as React from 'react';

import { ICoverage } from '../services/grouping';
import styles from './MetadataExplorer.module.scss';

export const CoverageTable = (props: { rows: ICoverage[]; total: number; onGroupBy: (fieldKey: string) => void }): JSX.Element => (
  <div className={styles.tableWrap}>
    <table className={styles.coverageTable}>
      <thead>
        <tr>
          <th>Metadata field</th>
          <th>Documents with a value</th>
          <th>Coverage</th>
          <th>Distinct values</th>
          <th>
            <span className={styles.srOnly}>Actions</span>
          </th>
        </tr>
      </thead>
      <tbody>
        {props.rows.map((row) => (
          <tr key={row.field.key}>
            <td className={styles.coverageField}>{row.field.title}</td>
            <td>
              {row.filled} of {props.total}
            </td>
            <td>
              <span className={styles.meter} role="img" aria-label={`${row.percent}% filled`}>
                <span
                  className={`${styles.meterFill} ${row.percent < 50 ? styles.meterLow : row.percent < 90 ? styles.meterMid : ''}`}
                  style={{ width: `${row.percent}%` }}
                />
              </span>
              <span className={styles.meterValue}>{row.percent}%</span>
            </td>
            <td>{row.distinct}</td>
            <td>
              <button type="button" className={styles.linkButton} onClick={() => props.onGroupBy(row.field.key)}>
                Group by this
              </button>
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  </div>
);
