import * as React from 'react';
import { PrimaryButton, MessageBar } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { useApp } from '../AppContext';
import { IBoard } from '../../models/types';
import { useMyWork, bucketMyWork, MyWorkTable } from './MyWork';
import { F } from '../../engine/fieldMap';

function greeting(now: Date): string {
  const h = now.getHours();
  return h < 12 ? 'Good morning' : h < 18 ? 'Good afternoon' : 'Good evening';
}

function BoardCard(props: { board: IBoard }): JSX.Element {
  const app = useApp();
  const { board } = props;
  const [counts, setCounts] = React.useState<{ [label: string]: number } | null>(null);
  React.useEffect(() => {
    let cancelled = false;
    app.services.items.statusCounts(board)
      .then(c => { if (!cancelled) { setCounts(c); } })
      .catch(() => { if (!cancelled) { setCounts({}); } });
    return () => { cancelled = true; };
  }, [board.id]);
  const statusCol = board.config.columns.filter(c => c.field === F.Status)[0];
  const total = counts ? Object.keys(counts).reduce((s, k) => s + counts[k], 0) : 0;
  return (
    <button type="button" className={styles.card} onClick={() => app.navigate({ page: 'board', boardId: board.id })}>
      <span className={styles.cardTitle}><span className={styles.boardDot} style={{ background: board.color }} />{board.title}</span>
      {board.description && <span className={styles.cardDesc}>{board.description}</span>}
      <div className={styles.mixBar} aria-label="Status mix">
        {counts && statusCol && (statusCol.labels || []).map(l => (counts[l.text] ? <span key={l.text} style={{ flex: counts[l.text], background: l.color }} title={`${l.text}: ${counts[l.text]}`} /> : null))}
        {counts && counts[''] ? <span style={{ flex: counts[''], background: 'var(--wb-border-light)' }} /> : null}
      </div>
      <span className={`${styles.small} ${styles.muted}`}>{counts ? `${total} ${total === 1 ? 'item' : 'items'}` : '…'}</span>
    </button>
  );
}

export function Home(props: { onNewBoard: () => void }): JSX.Element {
  const app = useApp();
  const { rows, error } = useMyWork();
  const b = rows ? bucketMyWork(rows) : null;
  const active = app.boards.filter(x => !x.archived);
  const recent = app.prefs.recent.map(id => active.filter(x => x.id === id)[0]).filter(x => !!x).slice(0, 6);
  const shown = recent.length > 0 ? recent : active.slice(0, 6);
  const firstName = app.me.title.split(' ')[0];

  return (
    <div className={styles.page}>
      <div className={styles.row}>
        <h1 className={styles.pageTitle}>{greeting(new Date())}, {firstName}</h1>
        <span className={styles.spacer} />
        <PrimaryButton text="New board" iconProps={{ iconName: 'Add' }} onClick={props.onNewBoard} />
      </div>

      {active.length === 0 ? (
        <div className={styles.center} style={{ border: '1px dashed var(--wb-border)', borderRadius: 8 }}>
          <h2 className={styles.sectionTitle}>Create your first board</h2>
          <p className={styles.muted} style={{ maxWidth: 420 }}>A board holds a project or a team&apos;s work: items in groups, with owners, status and dates. Start from a template and change it as you go.</p>
          <PrimaryButton text="New board" iconProps={{ iconName: 'Add' }} onClick={props.onNewBoard} />
        </div>
      ) : (
        <>
          <div className={styles.tiles}>
            <div className={styles.tile}><span className={styles.muted}>Overdue</span><span className={`${styles.tileValue} ${b && b.overdue.length ? styles.overdue : ''}`}>{b ? b.overdue.length : '–'}</span></div>
            <div className={styles.tile}><span className={styles.muted}>Due today</span><span className={styles.tileValue}>{b ? b.today.length : '–'}</span></div>
            <div className={styles.tile}><span className={styles.muted}>Due in the next 7 days</span><span className={styles.tileValue}>{b ? b.week.length : '–'}</span></div>
            <div className={styles.tile}><span className={styles.muted}>Assigned to me, open</span><span className={styles.tileValue}>{b && rows ? rows.length - b.done.length : '–'}</span></div>
          </div>

          <section>
            <div className={styles.row}>
              <h2 className={styles.sectionTitle}>Needs your attention</h2>
              <span className={styles.spacer} />
              <button type="button" className={styles.linkBtn} onClick={() => app.navigate({ page: 'mywork' })}>Open My Work</button>
            </div>
            {error && <MessageBar>{error}</MessageBar>}
            {b ? <MyWorkTable rows={b.overdue.concat(b.today, b.week)} limit={8} empty="Nothing overdue or due in the next 7 days." /> : <p className={styles.muted}>Loading…</p>}
          </section>

          <section>
            <h2 className={styles.sectionTitle} style={{ marginBottom: 8 }}>{recent.length > 0 ? 'Recent boards' : 'Boards'}</h2>
            <div className={styles.cards}>{shown.map(x => <BoardCard key={x.id} board={x} />)}</div>
          </section>
        </>
      )}
    </div>
  );
}
