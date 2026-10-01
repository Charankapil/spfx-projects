import * as React from 'react';
import { MessageBar } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { useApp } from '../AppContext';
import { IBoard, BoardRole, ROLE_LABELS } from '../../models/types';
import { Avatars, usePeople } from '../common/Common';
import { todayIso } from '../../engine/dates';
import { F } from '../../engine/fieldMap';
import { mapLimit } from '../../services/SpClient';

type Summary = { total: number; done: number; overdue: number; counts: { [label: string]: number } };

/** The roles a person holds on a board, in display order. Board owner counts too. */
export function rolesOf(board: IBoard, personId: number): (BoardRole | 'owner')[] {
  const out: (BoardRole | 'owner')[] = [];
  (['head', 'lead', 'sponsor'] as BoardRole[]).forEach(r => {
    if (board.roles[r].indexOf(personId) >= 0) {
      out.push(r);
    }
  });
  if (board.ownerIds.indexOf(personId) >= 0) {
    out.push('owner');
  }
  return out;
}

export function roleLabel(role: BoardRole | 'owner'): string {
  return role === 'owner' ? 'Board owner' : ROLE_LABELS[role];
}

function RolePeople(props: { label: string; ids: number[] }): JSX.Element | null {
  const people = usePeople(props.ids);
  if (props.ids.length === 0) {
    return null;
  }
  return (
    <span className={styles.row} style={{ gap: 6 }}>
      <span className={`${styles.small} ${styles.muted}`}>{props.label}</span>
      <Avatars people={people} max={3} size={20} />
    </span>
  );
}

/** A board card with the viewer's roles and progress. Also used on My team. */
export function ProjectCard(props: { board: IBoard; roles: (BoardRole | 'owner')[]; summary?: Summary; roleOf?: string }): JSX.Element {
  const app = useApp();
  const { board, summary } = props;
  const statusCol = board.config.columns.filter(c => c.field === F.Status)[0];
  const pct = summary && summary.total > 0 ? Math.round((summary.done / summary.total) * 100) : 0;
  return (
    <button type="button" className={styles.card} onClick={() => app.navigate({ page: 'board', boardId: board.id })}>
      <span className={styles.cardTitle}>
        <span className={styles.boardDot} style={{ background: board.color }} />{board.title}
      </span>
      <span className={styles.row} style={{ gap: 4 }}>
        {props.roleOf && <span className={styles.small}>{props.roleOf}:</span>}
        {props.roles.map(r => <span key={r} className={styles.pill}>{roleLabel(r)}</span>)}
      </span>
      <div className={styles.mixBar} aria-label="Status mix">
        {summary && statusCol && (statusCol.labels || []).map(l => (summary.counts[l.text]
          ? <span key={l.text} style={{ flex: summary.counts[l.text], background: l.color }} title={`${l.text}: ${summary.counts[l.text]}`} /> : null))}
        {summary && summary.counts[''] ? <span style={{ flex: summary.counts[''], background: 'var(--wb-border-light)' }} /> : null}
      </div>
      <span className={`${styles.small} ${styles.muted}`}>
        {summary
          ? `${pct}% done · ${summary.total} item${summary.total === 1 ? '' : 's'}${summary.overdue ? ` · ` : ''}`
          : 'Loading…'}
        {summary && summary.overdue > 0 && <span className={styles.overdue}>{summary.overdue} overdue</span>}
      </span>
      <span className={styles.row} style={{ gap: 12 }}>
        <RolePeople label="Head" ids={board.roles.head} />
        <RolePeople label="Lead" ids={board.roles.lead} />
        <RolePeople label="Sponsor" ids={board.roles.sponsor} />
      </span>
    </button>
  );
}

/** Progress figures for many boards, loaded a few at a time. */
export function useSummaries(boards: IBoard[]): { [boardId: number]: Summary } {
  const app = useApp();
  const [summaries, setSummaries] = React.useState<{ [boardId: number]: Summary }>({});
  const key = boards.map(b => b.id).join(',');
  React.useEffect(() => {
    let cancelled = false;
    const today = todayIso();
    mapLimit(boards, 4, async b => {
      try {
        const s = await app.services.items.summary(b, today);
        if (!cancelled) {
          setSummaries(prev => ({ ...prev, [b.id]: s }));
        }
      } catch {
        // A board the viewer can no longer open simply shows no figures.
      }
    }).catch(() => undefined);
    return () => { cancelled = true; };
  }, [key]);
  return summaries;
}

export function MyProjects(): JSX.Element {
  const app = useApp();
  const mine = app.boards.filter(b => !b.archived && rolesOf(b, app.me.id).length > 0);
  const summaries = useSummaries(mine);
  const sections: { title: string; boards: IBoard[] }[] = [
    { title: 'As project head', boards: mine.filter(b => b.roles.head.indexOf(app.me.id) >= 0) },
    { title: 'As project lead', boards: mine.filter(b => b.roles.lead.indexOf(app.me.id) >= 0 && b.roles.head.indexOf(app.me.id) < 0) },
    { title: 'As project sponsor', boards: mine.filter(b => b.roles.sponsor.indexOf(app.me.id) >= 0 && b.roles.head.indexOf(app.me.id) < 0 && b.roles.lead.indexOf(app.me.id) < 0) },
    { title: 'As board owner', boards: mine.filter(b => rolesOf(b, app.me.id).join() === 'owner') }
  ];
  const totals = mine.reduce((t, b) => {
    const s = summaries[b.id];
    return s ? { open: t.open + s.total - s.done, overdue: t.overdue + s.overdue } : t;
  }, { open: 0, overdue: 0 });

  return (
    <div className={styles.page}>
      <h1 className={styles.pageTitle}>My projects</h1>
      <p className={styles.muted} style={{ margin: 0 }}>
        Boards where you are project head, lead, sponsor or board owner. Set project roles in each board&apos;s settings.
      </p>
      {mine.length === 0 ? (
        <MessageBar>You don&apos;t hold a project role on any board yet. A board owner can add you in the board&apos;s settings under Project roles.</MessageBar>
      ) : (
        <>
          <div className={styles.tiles}>
            <div className={styles.tile}><span className={styles.muted}>Projects</span><span className={styles.tileValue}>{mine.length}</span></div>
            <div className={styles.tile}><span className={styles.muted}>Open items</span><span className={styles.tileValue}>{totals.open}</span></div>
            <div className={styles.tile}><span className={styles.muted}>Overdue</span><span className={`${styles.tileValue} ${totals.overdue ? styles.overdue : ''}`}>{totals.overdue}</span></div>
          </div>
          {sections.filter(s => s.boards.length > 0).map(s => (
            <section key={s.title}>
              <h2 className={styles.sectionTitle} style={{ marginBottom: 8 }}>{s.title} <span className={styles.muted}>({s.boards.length})</span></h2>
              <div className={styles.cards}>
                {s.boards.map(b => <ProjectCard key={b.id} board={b} roles={rolesOf(b, app.me.id)} summary={summaries[b.id]} />)}
              </div>
            </section>
          ))}
        </>
      )}
    </div>
  );
}
