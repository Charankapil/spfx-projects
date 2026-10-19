import * as React from 'react';

import { IDocumentRecord, IMetadataField } from '../models/IMetadata';
import styles from './MetadataExplorer.module.scss';

export interface IDocumentTableProps {
  documents: IDocumentRecord[];
  /** Fields shown as chips on each row (the grouping fields are left out: they are the group). */
  chipFields: IMetadataField[];
  locationTitle: string;
}

const PAGE = 20;

function formatDate(date?: Date): string {
  return date ? date.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' }) : '';
}

export const DocumentTable = (props: IDocumentTableProps): JSX.Element => {
  const [limit, setLimit] = React.useState(PAGE);
  const rows = props.documents.slice(0, limit);
  const remaining = props.documents.length - rows.length;

  return (
    <div className={styles.tableWrap}>
      <table className={styles.docTable}>
        <thead>
          <tr>
            <th>Name</th>
            <th>{props.locationTitle}</th>
            <th>Modified</th>
            <th>Metadata</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((doc) => (
            <tr key={doc.key}>
              <td>
                <span className={styles.docName}>
                  <span className={styles.fileBadge} aria-hidden="true">
                    {(doc.fileType || 'file').slice(0, 4)}
                  </span>
                  <a href={doc.openUrl} target="_blank" rel="noopener noreferrer" data-interception="off" title={doc.url}>
                    {doc.name}
                  </a>
                </span>
              </td>
              <td>
                {doc.locationUrl ? (
                  <a href={doc.locationUrl} target="_blank" rel="noopener noreferrer" data-interception="off">
                    {doc.location}
                  </a>
                ) : (
                  doc.location
                )}
              </td>
              <td className={styles.nowrap}>{formatDate(doc.modified)}</td>
              <td>
                <span className={styles.chips}>
                  {props.chipFields.map((f) =>
                    (doc.values[f.key] || []).map((v) => (
                      <span key={`${f.key}:${v}`} className={styles.chip} title={f.title}>
                        <span className={styles.chipLabel}>{f.title}</span>
                        {v}
                      </span>
                    ))
                  )}
                </span>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      {remaining > 0 && (
        <button type="button" className={styles.moreButton} onClick={() => setLimit(limit + PAGE * 5)}>
          Show {Math.min(remaining, PAGE * 5)} more of {remaining}
        </button>
      )}
    </div>
  );
};
