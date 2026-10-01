import * as React from 'react';
import { MessageBar, MessageBarType, Toggle, Pivot, PivotItem, Icon, Spinner, SpinnerSize } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { useApp } from '../AppContext';
import { IPerson, IPersonWorkRow } from '../../models/types';
import { Avatar, Avatars, LabelChip, PeoplePicker, usePeople } from '../common/Common';
import { todayIso, addDays, relativeDay } from '../../engine/dates';
import { teamTotals, sortTeamRows } from '../../engine/team';
import { ProjectCard, rolesOf, useSummaries } from './MyProjects';

/**
 * Work of the people who report to me: their items across boards, and the boards where they
 * hold project roles. Reports come from the SharePoint user profile (org data); people can
 * also be added by hand. Only boards the viewer can open are included.
 */
export function MyTeam(): JSX.Element {
  const app = useApp();
  const { services } = app;
  const [everyone, setEveryone] = React.useState(false);
  const [reports, setReports] = React.useState<IPerson[] | null>(null);
  const [reportNote, setReportNote] = React.useState('');
  const manual = usePeople(app.prefs.team);
  const [selected, setSelected] = React.useState<number | null>(null);
  const [rows, setRows] = React.useState<IPersonWorkRow[] | null>(null);
  const [error, setError] = React.useState('');

  React.useEffect(() => {
    let cancelled = false;
    setReports(null);
    setReportNote('');
    services.people.myReportLogins(everyone)
      .then(async logins => {
        const people = await services.people.siteUsersByLogin(logins);
        if (cancelled) {
          return;
        }
        setReports(people);
        const missing = logins.length - people.length;
        if (logins.length === 0) {
          setReportNote('Your profile lists no reports. Add people below to follow their work.');
        } else if (missing > 0) {
          setReportNote(`${missing} of your reports ${missing === 1 ? 'has' : 'have'} never been added to this site, so they have no work here.`);
        }
      })
      .catch(() => {
        if (!cancelled) {
          setReports([]);
          setReportNote('Your reports could not be read from your user profile. Add people below to follow their work.');
        }
      });
    return () => { cancelled = true; };
  }, [everyone]);

  const team: IPerson[] = (reports || []).concat(manual.filter(m => !(reports || []).some(r => r.id === m.id)));
  const teamKey = team.map(p => p.id).join(',');
  const active = app.boards.filter(b => !b.archived);

  React.useEffect(() => {
    if (reports === null) {
      return undefined;
    }
    let cancelled = false;
    setRows(null);
    services.items.itemsForPeople(active, team.map(p => p.id))
      .then(r => { if (!cancelled) { setRows(r); } })
      .catch(e => { if (!cancelled) { setError((e as Error).message); setRows([]); } });
    return () => { cancelled = true; };
  }, [teamKey, reports === null, active.length]);

  const today = todayIso();
  const totals = rows ? teamTotals(rows, team, today, addDays(today, 7)) : [];
  const shownPeople = selected === null ? team : team.filter(p => p.id === selected);
  const shownIds = shownPeople.map(p => p.id);
  const shownRows = sortTeamRows((rows || []).filter(r => r.people.some(p => shownIds.indexOf(p.id) >= 0)));
  const roleBoards = active.filter(b => shownIds.some(id => rolesOf(b, id).some(r => r !== 'owner')));
  const summaries = useSummaries(roleBoards);
  const open = shownRows.filter(r => !(r.status && r.status.isDone));
  const done = shownRows.filter(r => r.status && r.status.isDone);

  const table = (list: IPersonWorkRow[], empty: string): JSX.Element => (list.length === 0 ? <p className={styles.muted}>{empty}</p> : (
    <div style={{ overflowX: 'auto' }}>
      <table className={styles.simpleTable}>
        <thead><tr><th>Who</th><th>Item</th><th>Board</th><th>Status</th><th>Due</th></tr></thead>
        <tbody>
          {list.map(r => (
            <tr key={r.board.id + '-' + r.item.id} className={styles.clickable} tabIndex={0}
              onClick={() => app.navigate({ page: 'board', boardId: r.board.id, itemId: r.item.id })}
              onKeyDown={e => { if (e.key === 'Enter') { app.navigate({ page: 'board', boardId: r.board.id, itemId: r.item.id }); } }}>
              <td><span className={styles.row} style={{ gap: 6, flexWrap: 'nowrap' }}><Avatars people={r.people} max={2} size={22} />{r.people.length === 1 ? r.people[0].title : ''}</span></td>
              <td>{r.item.parentId ? <span className={styles.muted}>↳ </span> : null}{r.item.title}</td>
              <td><span className={styles.row}><span className={styles.boardDot} style={{ background: r.board.color }} />{r.board.title}</span></td>
              <td><LabelChip label={r.status} /></td>
              <td className={r.due && r.due < today && !(r.status && r.status.isDone) ? styles.overdue : ''}>{relativeDay(r.due)}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  ));

  return (
    <div className={styles.page}>
      <div className={styles.row}>
        <h1 className={styles.pageTitle}>My team</h1>
        <span className={styles.spacer} />
        <Toggle inlineLabel label="Include their teams" checked={everyone} onChange={(_, c) => setEveryone(!!c)} styles={{ root: { margin: 0 } }} />
      </div>
      <p className={styles.muted} style={{ margin: 0 }}>
        Work assigned to the people who report to you, in any People column, on boards you can open. Reports come from your organisation profile.
      </p>
      {reportNote && <MessageBar>{reportNote}</MessageBar>}
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}

      <div className={styles.row} style={{ alignItems: 'flex-end' }}>
        <div style={{ flex: '1 1 320px', maxWidth: 520 }}>
          <PeoplePicker label="Also follow (added by you)" selected={manual} onChange={people => app.setTeam(people.map(p => p.id))} />
        </div>
      </div>

      {reports === null ? <Spinner size={SpinnerSize.medium} label="Finding your team…" /> : team.length === 0 ? null : (
        <>
          <div className={styles.cards} role="list" aria-label="Team members">
            <button type="button" role="listitem" className={`${styles.card} ${selected === null ? styles.templateCardOn : ''}`} onClick={() => setSelected(null)}>
              <span className={styles.cardTitle}><Icon iconName="People" /> Everyone</span>
              <span className={`${styles.small} ${styles.muted}`}>{team.length} {team.length === 1 ? 'person' : 'people'}</span>
            </button>
            {team.map(p => {
              const t = totals.filter(x => x.person.id === p.id)[0];
              return (
                <button key={p.id} type="button" role="listitem" className={`${styles.card} ${selected === p.id ? styles.templateCardOn : ''}`}
                  onClick={() => setSelected(selected === p.id ? null : p.id)}>
                  <span className={styles.cardTitle}><Avatar person={p} size={24} />{p.title}</span>
                  <span className={`${styles.small} ${styles.muted}`}>
                    {t ? `${t.open} open · ${t.dueThisWeek} due this week · ${t.boards} board${t.boards === 1 ? '' : 's'}` : 'Loading…'}
                    {t && t.overdue > 0 && <> · <span className={styles.overdue}>{t.overdue} overdue</span></>}
                  </span>
                </button>
              );
            })}
          </div>

          {rows === null ? <Spinner size={SpinnerSize.medium} label="Collecting work across boards…" /> : (
            <Pivot>
              <PivotItem headerText="Open" itemCount={open.length}>
                <div style={{ paddingTop: 12 }}>{table(open, 'Nothing open.')}</div>
              </PivotItem>
              <PivotItem headerText="Done" itemCount={done.length}>
                <div style={{ paddingTop: 12 }}>{table(done, 'Nothing marked done yet.')}</div>
              </PivotItem>
              <PivotItem headerText="Projects they lead" itemCount={roleBoards.length}>
                <div style={{ paddingTop: 12 }}>
                  {roleBoards.length === 0 ? <p className={styles.muted}>No project roles on boards you can open.</p> : (
                    <div className={styles.cards}>
                      {roleBoards.map(b => {
                        const holders = shownPeople.filter(p => rolesOf(b, p.id).some(r => r !== 'owner'));
                        return (
                          <ProjectCard key={b.id} board={b} summary={summaries[b.id]}
                            roleOf={holders.length === 1 ? holders[0].title : `${holders.length} people`}
                            roles={holders.length === 1 ? rolesOf(b, holders[0].id).filter(r => r !== 'owner') : []} />
                        );
                      })}
                    </div>
                  )}
                </div>
              </PivotItem>
            </Pivot>
          )}
        </>
      )}
      <span className={`${styles.small} ${styles.muted}`}>Private boards you are not a member of are not included. People in a project role on a private board can see that board.</span>
    </div>
  );
}
