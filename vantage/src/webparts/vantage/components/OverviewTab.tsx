import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, type IColumn, MessageBarType, PrimaryButton, SelectionMode, Shimmer } from '@fluentui/react';

import styles from './VantageApp.module.scss';
import { Badge, Card, EmptyState, fmt, Kpi, useApp } from './common';
import { type IDashboard, WORK_STATUSES } from '../models/Models';

export const OverviewTab: React.FC<{ onGoToBoards: () => void }> = ({ onGoToBoards }) => {
  const { work, workspace, refreshToken, notify } = useApp();
  const [data, setData] = React.useState<IDashboard | undefined>(undefined);

  React.useEffect(() => {
    if (!workspace) return;
    let live = true;
    setData(undefined);
    work.dashboard(workspace)
      .then(d => { if (live) setData(d); })
      .catch(e => { if (live) notify(MessageBarType.error, `Could not load the dashboard: ${(e as Error).message}`); });
    return () => { live = false; };
  }, [workspace, work, refreshToken, notify]);

  if (!workspace) return null;
  if (!data) return <div className={styles.grid}>{[0, 1, 2, 3].map(i => <div key={i} className={styles.card}><Shimmer /><Shimmer width="60%" /></div>)}</div>;
  if (data.sampled === 0) {
    return (
      <Card>
        <EmptyState title="No tasks yet" text="Create a board and add the first tasks for this workspace."
          action={<PrimaryButton text="Go to Boards & tasks" onClick={onGoToBoards} />} />
      </Card>
    );
  }

  const total = Object.keys(data.byStatus).reduce((a, k) => a + data.byStatus[k], 0);
  const columns: IColumn[] = [
    { key: 'name', name: 'Assignee', fieldName: 'name', minWidth: 140, maxWidth: 260, isRowHeader: true },
    { key: 'open', name: 'Open', minWidth: 60, maxWidth: 80, onRender: (r: IDashboard['workload'][0]) => fmt(r.open) },
    { key: 'overdue', name: 'Overdue', minWidth: 70, maxWidth: 90,
      onRender: (r: IDashboard['workload'][0]) => r.overdue ? <span className={styles.bad}>{fmt(r.overdue)}</span> : '0' },
    { key: 'bar', name: 'Load', minWidth: 120,
      onRender: (r: IDashboard['workload'][0]) => {
        const max = data.workload.length ? data.workload[0].open : 1;
        return <div className={styles.meter} aria-hidden="true"><span style={{ width: `${Math.max(4, (r.open / max) * 100)}%` }} /></div>;
      } }
  ];

  return (
    <div className={styles.stack}>
      <div className={styles.grid}>
        <Kpi label="Open tasks" value={fmt(data.open)} note={`${fmt(data.dueThisWeek)} due in the next 7 days`} />
        <Kpi label="Overdue" value={fmt(data.overdue)} tone={data.overdue ? styles.bad : undefined} note="Past due date, not completed" />
        <Kpi label="Assigned to me" value={fmt(data.mine)} note="Open tasks" />
        <Kpi label="Completed / cancelled" value={fmt(data.completed)} note="Eligible for archiving" />
      </div>
      <div className={styles.twoColEven}>
        <Card title="Tasks by status">
          <ul className={styles.statusBars}>
            {WORK_STATUSES.map(s => {
              const n = data.byStatus[s] || 0;
              return (
                <li key={s}>
                  <span className={styles.statusLabel}><Badge value={s} /></span>
                  <span className={styles.meter} aria-hidden="true"><span style={{ width: `${total ? (n / total) * 100 : 0}%` }} /></span>
                  <span className={styles.statusCount}>{fmt(n)}</span>
                </li>
              );
            })}
          </ul>
        </Card>
        <Card title="Open work by assignee">
          {data.workload.length
            ? <DetailsList items={data.workload.slice(0, 10)} columns={columns} selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.justified} compact getKey={(r: IDashboard['workload'][0]) => r.key || 'unassigned'} />
            : <p className={styles.muted}>No open tasks.</p>}
        </Card>
      </div>
      {data.sampled >= data.sampleLimit && (
        <p className={styles.muted}>Figures are based on the {fmt(data.sampleLimit)} most recent tasks in this workspace.</p>
      )}
    </div>
  );
};
