import * as React from 'react';
import { MessageBar, MessageBarType, Pivot, PivotItem } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { useApp } from '../AppContext';
import { IMyWorkRow } from '../../models/types';
import { Loading, LabelChip } from '../common/Common';
import { todayIso, addDays, relativeDay } from '../../engine/dates';

export interface IMyWorkBuckets {
  overdue: IMyWorkRow[];
  today: IMyWorkRow[];
  week: IMyWorkRow[];
  later: IMyWorkRow[];
  noDate: IMyWorkRow[];
  done: IMyWorkRow[];
}

/** Sort My Work rows into date buckets, the way monday.com's My Work does. */
export function bucketMyWork(rows: IMyWorkRow[], now: Date = new Date()): IMyWorkBuckets {
  const today = todayIso(now);
  const weekEnd = addDays(today, 7);
  const b: IMyWorkBuckets = { overdue: [], today: [], week: [], later: [], noDate: [], done: [] };
  rows.forEach(r => {
    if (r.status && r.status.isDone) {
      b.done.push(r);
    } else if (!r.due) {
      b.noDate.push(r);
    } else if (r.due < today) {
      b.overdue.push(r);
    } else if (r.due === today) {
      b.today.push(r);
    } else if (r.due <= weekEnd) {
      b.week.push(r);
    } else {
      b.later.push(r);
    }
  });
  const byDue = (x: IMyWorkRow, y: IMyWorkRow): number => (x.due || '9999').localeCompare(y.due || '9999') || x.item.title.localeCompare(y.item.title);
  (Object.keys(b) as (keyof IMyWorkBuckets)[]).forEach(k => b[k].sort(byDue));
  return b;
}

/** Load My Work once per mount; shared by Home and My Work pages. */
export function useMyWork(): { rows: IMyWorkRow[] | null; error: string } {
  const app = useApp();
  const [rows, setRows] = React.useState<IMyWorkRow[] | null>(null);
  const [error, setError] = React.useState('');
  React.useEffect(() => {
    let cancelled = false;
    app.services.items.myWork(app.boards.filter(b => !b.archived), app.me)
      .then(r => { if (!cancelled) { setRows(r); } })
      .catch(e => { if (!cancelled) { setError((e as Error).message); setRows([]); } });
    return () => { cancelled = true; };
  }, [app.boards.length]);
  return { rows, error };
}

export function MyWorkTable(props: { rows: IMyWorkRow[]; empty: string; limit?: number }): JSX.Element {
  const app = useApp();
  const today = todayIso();
  const rows = props.limit ? props.rows.slice(0, props.limit) : props.rows;
  if (rows.length === 0) {
    return <p className={styles.muted}>{props.empty}</p>;
  }
  return (
    <div style={{ overflowX: 'auto' }}>
      <table className={styles.simpleTable}>
        <thead>
          <tr><th>Item</th><th>Board</th><th>Status</th><th>Due</th></tr>
        </thead>
        <tbody>
          {rows.map(r => (
            <tr key={r.board.id + '-' + r.item.id} className={styles.clickable} tabIndex={0}
              onClick={() => app.navigate({ page: 'board', boardId: r.board.id, itemId: r.item.id })}
              onKeyDown={e => { if (e.key === 'Enter') { app.navigate({ page: 'board', boardId: r.board.id, itemId: r.item.id }); } }}>
              <td>{r.item.parentId ? <span className={styles.muted}>↳ </span> : null}{r.item.title}</td>
              <td><span className={styles.row}><span className={styles.boardDot} style={{ background: r.board.color }} />{r.board.title}</span></td>
              <td><LabelChip label={r.status} /></td>
              <td className={r.due && r.due < today && !(r.status && r.status.isDone) ? styles.overdue : ''}>{relativeDay(r.due)}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export function MyWork(): JSX.Element {
  const { rows, error } = useMyWork();
  if (rows === null) {
    return <Loading label="Finding your work across boards…" />;
  }
  const b = bucketMyWork(rows);
  const sections: { key: keyof IMyWorkBuckets; title: string; empty: string }[] = [
    { key: 'overdue', title: 'Overdue', empty: 'Nothing overdue.' },
    { key: 'today', title: 'Today', empty: 'Nothing due today.' },
    { key: 'week', title: 'Next 7 days', empty: 'Nothing due in the next 7 days.' },
    { key: 'later', title: 'Later', empty: 'Nothing due later.' },
    { key: 'noDate', title: 'Without a date', empty: 'Everything assigned to you has a date.' }
  ];
  return (
    <div className={styles.page}>
      <h1 className={styles.pageTitle}>My Work</h1>
      <p className={styles.muted} style={{ margin: 0 }}>Items where you are in the Owner column, on every board you can open on this site.</p>
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      <Pivot>
        <PivotItem headerText="To do" itemCount={rows.length - b.done.length}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8, paddingTop: 12 }}>
            {sections.map(s => (
              <section key={s.key}>
                <h2 className={styles.sectionTitle} style={s.key === 'overdue' && b.overdue.length ? { color: 'var(--wb-danger)' } : undefined}>
                  {s.title} <span className={styles.muted}>({b[s.key].length})</span>
                </h2>
                <MyWorkTable rows={b[s.key]} empty={s.empty} />
              </section>
            ))}
          </div>
        </PivotItem>
        <PivotItem headerText="Done" itemCount={b.done.length}>
          <div style={{ paddingTop: 12 }}><MyWorkTable rows={b.done} empty="Nothing marked done yet." /></div>
        </PivotItem>
      </Pivot>
    </div>
  );
}
