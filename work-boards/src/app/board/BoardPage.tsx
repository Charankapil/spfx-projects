import * as React from 'react';
import { SearchBox, DefaultButton, PrimaryButton, Icon, Callout, Checkbox, MessageBar, MessageBarType, ChoiceGroup, TooltipHost, DirectionalHint, IContextualMenuItem } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { useApp } from '../AppContext';
import { useBoard } from './useBoard';
import { BoardViewType, IBoardColumn } from '../../models/types';
import { IViewState, EMPTY_VIEW, activeFilterCount, peopleOnBoard } from '../../engine/viewQuery';
import { TableView } from '../views/TableView';
import { KanbanView } from '../views/KanbanView';
import { TimelineView } from '../views/TimelineView';
import { ItemPanel } from '../item/ItemPanel';
import { LabelEditor } from './LabelEditor';
import { BoardSettings } from './BoardSettings';
import { Avatar, Avatars, Loading, download, usePeople } from '../common/Common';
import { boardToCsv } from '../../engine/csv';

export interface IBoardPageProps {
  boardId: number;
  view?: BoardViewType;
  itemId?: number;
}

const VIEW_TABS: { key: BoardViewType; text: string; icon: string }[] = [
  { key: 'table', text: 'Main table', icon: 'Table' },
  { key: 'kanban', text: 'Kanban', icon: 'BacklogBoard' },
  { key: 'timeline', text: 'Timeline', icon: 'Timeline' }
];

function viewKey(webUrl: string, boardId: number): string {
  return `wb.view:${webUrl}:${boardId}`;
}

function loadView(webUrl: string, boardId: number): IViewState {
  try {
    const raw = window.localStorage.getItem(viewKey(webUrl, boardId));
    return raw ? { ...EMPTY_VIEW, ...JSON.parse(raw), search: '' } : EMPTY_VIEW;
  } catch {
    return EMPTY_VIEW;
  }
}

function saveView(webUrl: string, boardId: number, v: IViewState): void {
  try {
    window.localStorage.setItem(viewKey(webUrl, boardId), JSON.stringify({ ...v, search: '' }));
  } catch {
    // Storage may be blocked; views then last for the session only.
  }
}

export function BoardPage(props: IBoardPageProps): JSX.Element {
  const app = useApp();
  const [state, actions] = useBoard(props.boardId);
  const { board, items, rights } = state;
  const [view, setViewState] = React.useState<IViewState>(() => loadView(app.webUrl, props.boardId));
  const [labelCol, setLabelCol] = React.useState<IBoardColumn | null>(null);
  const [settingsOpen, setSettingsOpen] = React.useState(false);
  const [popup, setPopup] = React.useState<{ kind: 'person' | 'filter' | 'hide'; target: HTMLElement } | null>(null);
  const [title, setTitle] = React.useState('');
  const heads = usePeople(board ? board.roles.head : []);
  const leads = usePeople(board ? board.roles.lead : []);
  const sponsors = usePeople(board ? board.roles.sponsor : []);

  React.useEffect(() => setViewState(loadView(app.webUrl, props.boardId)), [props.boardId]);
  React.useEffect(() => { if (board) { setTitle(board.title); } }, [board && board.title]);
  React.useEffect(() => {
    app.services.prefs.touchRecent(props.boardId).catch(() => undefined);
  }, [props.boardId]);

  const setView = (v: IViewState): void => {
    setViewState(v);
    saveView(app.webUrl, props.boardId, v);
  };

  if (state.loading && !board) {
    return <Loading label="Loading board…" />;
  }
  if (!board) {
    return (
      <div className={styles.center}>
        <p>{state.error || 'This board could not be found.'}</p>
        <DefaultButton text="Back to home" onClick={() => app.navigate({ page: 'home' })} />
      </div>
    );
  }

  const activeView: BoardViewType = props.view || board.config.defaultView || 'table';
  const openItem = (itemId: number): void => app.navigate({ page: 'board', boardId: board.id, view: activeView, itemId });
  const closeItem = (): void => app.navigate({ page: 'board', boardId: board.id, view: activeView });
  const openedItem = props.itemId ? items.filter(i => i.id === props.itemId)[0] : undefined;
  const people = peopleOnBoard(board, items);
  const person = people.filter(p => p.id === view.personId)[0];
  const filterCount = activeFilterCount(view);
  const labelCols = board.config.columns.filter(c => c.type === 'status' || c.type === 'dropdown');
  const isFav = app.prefs.favourites.indexOf(board.id) >= 0;
  const sortMenu: IContextualMenuItem[] = [{ key: 'title', text: 'Item' }].concat(board.config.columns.map(c => ({ key: c.id, text: c.title }))).map(o => ({
    key: o.key,
    text: o.text,
    subMenuProps: {
      items: [
        { key: 'asc', text: 'Ascending', onClick: () => setView({ ...view, sort: { columnId: o.key, dir: 'asc' } }) },
        { key: 'desc', text: 'Descending', onClick: () => setView({ ...view, sort: { columnId: o.key, dir: 'desc' } }) }
      ]
    }
  }));
  if (view.sort) {
    sortMenu.push({ key: 'clear', text: 'Clear sort', onClick: () => setView({ ...view, sort: null }) });
  }

  const newItem = async (): Promise<void> => {
    const group = board.config.groups[0];
    const created = await actions.addItem(group.id, 'New item', { atTop: true });
    if (created) {
      openItem(created.id);
    }
  };

  const commitTitle = (): void => {
    const t = title.trim();
    if (!t || t === board.title) {
      setTitle(board.title);
      return;
    }
    app.services.boards.updateProps(board, { title: t }).then(actions.setBoard).catch(e => actions.setError((e as Error).message));
  };

  const toggleLabelFilter = (colId: string, text: string): void => {
    const cur = view.labelFilters[colId] || [];
    const next = cur.indexOf(text) >= 0 ? cur.filter(t => t !== text) : cur.concat([text]);
    setView({ ...view, labelFilters: { ...view.labelFilters, [colId]: next } });
  };

  return (
    <div className={styles.main}>
      <header className={styles.boardHeader}>
        <div className={styles.row} style={{ flexWrap: 'nowrap' }}>
          <span className={styles.boardDot} style={{ background: board.color, width: 14, height: 14 }} />
          <input className={styles.boardTitle} value={title} disabled={!rights.canManage} aria-label="Board name" size={Math.max(8, title.length + 1)}
            onChange={e => setTitle(e.target.value)} onBlur={commitTitle} onKeyDown={e => { if (e.key === 'Enter') { (e.target as HTMLInputElement).blur(); } }} />
          <TooltipHost content={isFav ? 'Remove from favourites' : 'Add to favourites'}>
            <button type="button" className={styles.iconBtn} onClick={() => app.toggleFavourite(board.id)} aria-pressed={isFav} aria-label="Favourite">
              <Icon iconName={isFav ? 'FavoriteStarFill' : 'FavoriteStar'} style={{ color: isFav ? '#ffcb00' : undefined }} />
            </button>
          </TooltipHost>
          {board.privacy === 'Private' && <span className={styles.pill}><Icon iconName="Lock" /> Private</span>}
          {board.archived && <span className={styles.pill}><Icon iconName="Archive" /> Archived</span>}
          <span className={styles.spacer} />
          <DefaultButton text="Export" iconProps={{ iconName: 'Download' }} onClick={() => download(`${board.title.replace(/[^\w\- ]+/g, '')}.csv`, boardToCsv(board, items), 'text/csv;charset=utf-8')} />
          <DefaultButton text="Settings" iconProps={{ iconName: 'Settings' }} onClick={() => setSettingsOpen(true)} />
        </div>
        {board.description && <p className={styles.muted} style={{ margin: 0 }}>{board.description}</p>}
        {(heads.length > 0 || leads.length > 0 || sponsors.length > 0) && (
          <div className={styles.row} style={{ gap: 16 }} aria-label="Project roles">
            {([['Head', heads], ['Lead', leads], ['Sponsor', sponsors]] as [string, typeof heads][]).filter(r => r[1].length > 0).map(([label, people]) => (
              <span key={label} className={styles.row} style={{ gap: 6 }}>
                <span className={`${styles.small} ${styles.muted}`}>{label}</span>
                <Avatars people={people} max={3} size={22} />
                <span className={styles.small}>{people.length === 1 ? people[0].title : `${people.length} people`}</span>
              </span>
            ))}
          </div>
        )}
        <nav className={styles.tabs} aria-label="Board views">
          {VIEW_TABS.map(t => (
            <button key={t.key} type="button" className={`${styles.tab} ${activeView === t.key ? styles.tabActive : ''}`} aria-current={activeView === t.key ? 'page' : undefined}
              onClick={() => app.navigate({ page: 'board', boardId: board.id, view: t.key })}>
              <Icon iconName={t.icon} /> {t.text}
            </button>
          ))}
        </nav>
      </header>

      <div className={styles.toolbar} role="toolbar" aria-label="Board tools">
        {rights.canAdd && <PrimaryButton text="New item" iconProps={{ iconName: 'Add' }} onClick={() => { newItem().catch(() => undefined); }} />}
        <SearchBox placeholder="Search" value={view.search} onChange={(_, v) => setView({ ...view, search: v || '' })} styles={{ root: { width: 200 } }} />
        <DefaultButton onClick={e => setPopup({ kind: 'person', target: e.currentTarget as HTMLElement })} iconProps={person ? undefined : { iconName: 'Contact' }}
          checked={!!person} text={person ? person.title : 'Person'} />
        <DefaultButton onClick={e => setPopup({ kind: 'filter', target: e.currentTarget as HTMLElement })} iconProps={{ iconName: 'Filter' }}
          checked={filterCount > 0} text={filterCount > 0 ? `Filter · ${filterCount}` : 'Filter'} />
        {activeView === 'table' && (
          <DefaultButton iconProps={{ iconName: 'Sort' }} checked={!!view.sort}
            text={view.sort ? 'Sorted by ' + (view.sort.columnId === 'title' ? 'Item' : (board.config.columns.filter(c => c.id === (view.sort as { columnId: string }).columnId)[0] || { title: '?' }).title) : 'Sort'}
            menuProps={{ items: sortMenu }} />
        )}
        {activeView === 'table' && (
          <DefaultButton onClick={e => setPopup({ kind: 'hide', target: e.currentTarget as HTMLElement })} iconProps={{ iconName: 'Hide3' }}
            checked={view.hiddenColumnIds.length > 0} text={view.hiddenColumnIds.length > 0 ? `Hidden · ${view.hiddenColumnIds.length}` : 'Hide'} />
        )}
        {(filterCount > 0 || person || view.search || view.sort) && (
          <button type="button" className={styles.linkBtn} onClick={() => setView({ ...view, search: '', personId: null, labelFilters: {}, dateFilter: '', sort: null })}>Clear all</button>
        )}
      </div>

      {state.error && <MessageBar messageBarType={MessageBarType.error} onDismiss={actions.clearMessages} styles={{ root: { margin: '0 24px' } }}>{state.error}</MessageBar>}
      {state.notice && <MessageBar messageBarType={MessageBarType.info} onDismiss={actions.clearMessages} styles={{ root: { margin: '0 24px' } }}>{state.notice}</MessageBar>}
      {!rights.canEdit && <MessageBar styles={{ root: { margin: '0 24px' } }}>You can view this board but not change it.</MessageBar>}

      {activeView === 'table' && (
        <TableView board={board} items={items} rights={rights} view={view} setView={setView} actions={actions}
          updateCounts={state.updateCounts} onOpenItem={openItem} onEditLabels={setLabelCol} />
      )}
      {activeView === 'kanban' && (
        <KanbanView board={board} items={items} rights={rights} view={view} actions={actions} updateCounts={state.updateCounts} onOpenItem={openItem} />
      )}
      {activeView === 'timeline' && (
        <TimelineView board={board} items={items} rights={rights} view={view} actions={actions} onOpenItem={openItem} />
      )}

      {openedItem && (
        <ItemPanel board={board} item={openedItem} items={items} rights={rights} actions={actions}
          onClose={closeItem} onOpenItem={openItem} onEditLabels={setLabelCol} />
      )}
      {labelCol && (
        <LabelEditor column={board.config.columns.filter(c => c.id === labelCol.id)[0] || labelCol}
          onSave={(labels, renames) => actions.saveLabels(labelCol.id, labels, renames)} onClose={() => setLabelCol(null)} />
      )}
      {settingsOpen && (
        <BoardSettings board={board} rights={rights} onClose={() => setSettingsOpen(false)}
          onChanged={b => { actions.setBoard(b); setTitle(b.title); if (b.privacy !== board.privacy) { actions.reload(true).catch(() => undefined); } }}
          onDeleted={() => { setSettingsOpen(false); app.reloadBoards().catch(() => undefined); app.navigate({ page: 'home' }); }} />
      )}

      {popup && (
        <Callout target={popup.target} onDismiss={() => setPopup(null)} directionalHint={DirectionalHint.bottomLeftEdge} setInitialFocus>
          <div className={styles.popup} style={{ minWidth: 240, maxHeight: 420, overflowY: 'auto' }}>
            {popup.kind === 'person' && (
              <>
                <strong>Show items for</strong>
                {people.length === 0 && <span className={styles.muted}>Nobody is assigned on this board yet.</span>}
                {people.map(p => (
                  <button key={p.id} type="button" className={`${styles.navItem} ${view.personId === p.id ? styles.navItemActive : ''}`}
                    onClick={() => { setView({ ...view, personId: view.personId === p.id ? null : p.id }); setPopup(null); }}>
                    <Avatar person={p} size={22} /> <span className={styles.navLabel}>{p.title}</span>
                  </button>
                ))}
                {view.personId !== null && <button type="button" className={styles.linkBtn} onClick={() => { setView({ ...view, personId: null }); setPopup(null); }}>Show everyone</button>}
              </>
            )}
            {popup.kind === 'filter' && (
              <>
                {labelCols.map(col => (
                  <div key={col.id}>
                    <strong>{col.title}</strong>
                    {(col.labels || []).map(l => (
                      <Checkbox key={l.id || l.text} label={l.text} checked={(view.labelFilters[col.id] || []).indexOf(l.text) >= 0}
                        onChange={() => toggleLabelFilter(col.id, l.text)} styles={{ root: { marginTop: 4 } }} />
                    ))}
                    <Checkbox label="(Empty)" checked={(view.labelFilters[col.id] || []).indexOf('') >= 0} onChange={() => toggleLabelFilter(col.id, '')} styles={{ root: { marginTop: 4 } }} />
                  </div>
                ))}
                <ChoiceGroup label="Date" selectedKey={view.dateFilter || 'any'}
                  options={[
                    { key: 'any', text: 'Any date' },
                    { key: 'overdue', text: 'Overdue' },
                    { key: 'thisweek', text: 'Due in the next 7 days' },
                    { key: 'nodate', text: 'No date' }
                  ]}
                  onChange={(_, o) => setView({ ...view, dateFilter: o && o.key !== 'any' ? (o.key as IViewState['dateFilter']) : '' })} />
                {filterCount > 0 && <button type="button" className={styles.linkBtn} onClick={() => setView({ ...view, labelFilters: {}, dateFilter: '' })}>Clear filters</button>}
              </>
            )}
            {popup.kind === 'hide' && (
              <>
                <strong>Columns</strong>
                {board.config.columns.map(c => (
                  <Checkbox key={c.id} label={c.title} checked={view.hiddenColumnIds.indexOf(c.id) < 0}
                    onChange={(_, checked) => setView({ ...view, hiddenColumnIds: checked ? view.hiddenColumnIds.filter(id => id !== c.id) : view.hiddenColumnIds.concat([c.id]) })} />
                ))}
              </>
            )}
          </div>
        </Callout>
      )}
    </div>
  );
}
