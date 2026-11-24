import * as React from 'react';
import { PrimaryButton, MessageBar, Icon } from '@fluentui/react';
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

const TOUR_KEY = 'wb.tourDismissed';

function tourDismissed(): boolean {
  try {
    return window.localStorage.getItem(TOUR_KEY) === '1';
  } catch {
    return false;
  }
}

/** Three plain steps for someone new. Dismissed per browser; Help can bring it back. */
function GettingStarted(props: { hasBoards: boolean; onNewBoard: () => void; onSearch: () => void; onHelp: () => void; onDismiss: () => void }): JSX.Element {
  const steps = [
    {
      done: props.hasBoards,
      title: props.hasBoards ? 'Open a board' : 'Create your first board',
      text: props.hasBoards
        ? 'Pick a board in the left menu, or press Ctrl + K and type part of its name.'
        : 'A board is one project or one team\'s list of work. Start from a ready-made template.',
      action: props.hasBoards ? { label: 'Find a board', run: props.onSearch } : { label: 'Create a board', run: props.onNewBoard }
    },
    {
      done: false,
      title: 'Add work and assign people',
      text: 'Type in "+ Add item" at the bottom of a group. Click the Owner cell to assign someone, and the Status cell to set progress.',
      action: undefined
    },
    {
      done: false,
      title: 'Keep track without searching',
      text: 'My Work collects everything assigned to you. Mention colleagues with @ in an item\'s updates so they see it too.',
      action: { label: 'How it works', run: props.onHelp }
    }
  ];
  return (
    <section className={styles.tour} aria-label="Getting started">
      <div className={styles.row}>
        <h2 className={styles.sectionTitle} style={{ margin: 0 }}>Getting started</h2>
        <span className={styles.muted}>Three steps. You can come back to this from Help.</span>
        <span className={styles.spacer} />
        <button type="button" className={styles.linkBtn} onClick={props.onDismiss}>Hide</button>
      </div>
      <ol className={styles.tourSteps}>
        {steps.map((s, i) => (
          <li key={s.title} className={`${styles.tourStep} ${s.done ? styles.tourStepDone : ''}`}>
            <span className={styles.tourNum} aria-hidden="true">{s.done ? <Icon iconName="CheckMark" /> : i + 1}</span>
            <div className={styles.tourBody}>
              <strong>{s.title}</strong>
              <span className={styles.muted}>{s.text}</span>
              {s.action && <button type="button" className={styles.linkBtn} style={{ alignSelf: 'flex-start', padding: 0 }} onClick={s.action.run}>{s.action.label} →</button>}
            </div>
          </li>
        ))}
      </ol>
    </section>
  );
}

export function Home(props: { onNewBoard: () => void; onHelp: () => void; onSearch: () => void }): JSX.Element {
  const app = useApp();
  const { rows, error } = useMyWork();
  const b = rows ? bucketMyWork(rows) : null;
  const active = app.boards.filter(x => !x.archived);
  const recent = app.prefs.recent.map(id => active.filter(x => x.id === id)[0]).filter(x => !!x).slice(0, 6);
  const shown = recent.length > 0 ? recent : active.slice(0, 6);
  const firstName = app.me.title.split(' ')[0];
  const [showTour, setShowTour] = React.useState(() => !tourDismissed());
  const hideTour = (): void => {
    setShowTour(false);
    try {
      window.localStorage.setItem(TOUR_KEY, '1');
    } catch {
      // Hidden for this visit only.
    }
  };
  const quick: { label: string; icon: string; run: () => void }[] = [
    { label: 'New board', icon: 'Add', run: props.onNewBoard },
    { label: 'My Work', icon: 'CheckList', run: () => app.navigate({ page: 'mywork' }) },
    { label: 'My projects', icon: 'ProjectCollection', run: () => app.navigate({ page: 'projects' }) },
    { label: 'My team', icon: 'People', run: () => app.navigate({ page: 'team' }) },
    { label: 'Find a board', icon: 'Search', run: props.onSearch },
    { label: 'Help', icon: 'Help', run: props.onHelp }
  ];

  return (
    <div className={styles.page}>
      <div className={styles.row}>
        <h1 className={styles.pageTitle}>{greeting(new Date())}, {firstName}</h1>
      </div>

      {showTour && <GettingStarted hasBoards={active.length > 0} onNewBoard={props.onNewBoard} onSearch={props.onSearch} onHelp={props.onHelp} onDismiss={hideTour} />}

      <nav className={styles.quickActions} aria-label="Quick actions">
        {quick.map(q => (
          <button key={q.label} type="button" className={styles.quickAction} onClick={q.run}>
            <span className={styles.quickIcon}><Icon iconName={q.icon} /></span>
            <span>{q.label}</span>
          </button>
        ))}
      </nav>

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
