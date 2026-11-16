import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from '@fluentui/react';
import { IChange } from '../services/compare';
import styles from './AccessReport.module.scss';

export const Card: React.FC<{ title?: string; right?: React.ReactNode; children: React.ReactNode }> = ({ title, right, children }) => (
  <section className={styles.card}>
    {(title || right) && (
      <div className={styles.cardHead}>
        {title ? <h3>{title}</h3> : <span />}
        {right}
      </div>
    )}
    {children}
  </section>
);

export const Facts: React.FC<{ items: Array<{ label: string; value: React.ReactNode; warn?: boolean }>; big?: boolean }> = ({ items, big }) => (
  <dl className={styles.facts}>
    {items.map((i) => (
      <div key={i.label} className={`${big ? styles.stat : ''} ${i.warn ? styles.warn : ''}`}>
        <dt>{i.label}</dt>
        <dd>{i.value}</dd>
      </div>
    ))}
  </dl>
);

export const Empty: React.FC<{ text: string }> = ({ text }) => <div className={styles.empty}>{text}</div>;

export function errorText(e: unknown): string {
  return (e as Error).message || String(e);
}

export function dateTime(d: Date | undefined): string {
  return d ? d.toLocaleString(undefined, { year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' }) : '';
}

/** How many rows a grid shows; the files always hold everything. */
export const GRID_LIMIT = 2000;

const siteCell = (title: string, url: string): JSX.Element => (
  <span style={{ minWidth: 0, display: 'block' }}>
    <a href={url} target="_blank" rel="noopener noreferrer" className={styles.link}>
      {title || url}
    </a>
    <div className={styles.cellSub}>{url}</div>
  </span>
);

const whereCell = (r: string[], at: number): JSX.Element =>
  r[at] === 'Site' || r[at] === 'Subsite' ? (
    <span className={styles.pill}>{r[at] === 'Site' ? 'Whole site' : r[at + 5] === 'Yes' ? 'Subsite · own permissions' : 'Subsite'}</span>
  ) : (
    <span>
      {r[at + 2] ? (
        <a href={r[at + 2]} target="_blank" rel="noopener noreferrer" className={styles.link}>
          {r[at + 1] || '(no name)'}
        </a>
      ) : (
        r[at + 1] || '(no name)'
      )}
    </span>
  );

/** Report rows (REPORT_HEADERS order). */
export const RowsGrid: React.FC<{ rows: string[][] }> = ({ rows }) => {
  const columns: IColumn[] = [
    { key: 'who', name: 'Who', minWidth: 150, maxWidth: 240, isResizable: true, onRender: (r: string[]) => r[0] },
    { key: 'site', name: 'Site', minWidth: 190, maxWidth: 320, isResizable: true, onRender: (r: string[]) => siteCell(r[2], r[1]) },
    { key: 'where', name: 'Where', minWidth: 140, maxWidth: 240, isResizable: true, onRender: (r: string[]) => whereCell(r, 3) },
    { key: 'level', name: 'Access', minWidth: 110, maxWidth: 170, isResizable: true, onRender: (r: string[]) => r[6] },
    { key: 'how', name: 'How', minWidth: 170, isResizable: true, onRender: (r: string[]) => r[7] }
  ];
  return (
    <div className={styles.grid} data-is-scrollable="true">
      <DetailsList items={rows.slice(0, GRID_LIMIT)} columns={columns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} compact />
    </div>
  );
};

const CHANGE_CLASS: { [k: string]: string } = { Added: styles.pillAdded, Removed: styles.pillRemoved, Changed: styles.pillChanged };

export const ChangesGrid: React.FC<{ changes: IChange[] }> = ({ changes }) => {
  const columns: IColumn[] = [
    { key: 'kind', name: 'Change', minWidth: 80, maxWidth: 90, onRender: (c: IChange) => <span className={`${styles.pill} ${CHANGE_CLASS[c.kind]}`}>{c.kind}</span> },
    { key: 'who', name: 'Who', minWidth: 150, maxWidth: 240, isResizable: true, onRender: (c: IChange) => c.row[0] },
    { key: 'site', name: 'Site', minWidth: 190, maxWidth: 320, isResizable: true, onRender: (c: IChange) => siteCell(c.row[2], c.row[1]) },
    { key: 'where', name: 'Where', minWidth: 140, maxWidth: 240, isResizable: true, onRender: (c: IChange) => whereCell(c.row, 3) },
    { key: 'level', name: 'Access', minWidth: 150, maxWidth: 240, isResizable: true, onRender: (c: IChange) => (c.kind === 'Changed' ? `${c.was} → ${c.row[6]}` : c.row[6]) },
    { key: 'how', name: 'How', minWidth: 170, isResizable: true, onRender: (c: IChange) => c.row[7] }
  ];
  return (
    <div className={styles.grid} data-is-scrollable="true">
      <DetailsList items={changes.slice(0, GRID_LIMIT)} columns={columns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} compact />
    </div>
  );
};

/** Case-insensitive "contains" over a row's cells. */
export function matches(cells: string[], query: string): boolean {
  const q = query.trim().toLowerCase();
  return !q || cells.join(' ').toLowerCase().indexOf(q) >= 0;
}
