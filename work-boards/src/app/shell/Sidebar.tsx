import * as React from 'react';
import { Icon, SearchBox } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { useApp } from '../AppContext';
import { Route } from '../router';
import { IBoard } from '../../models/types';

export function Sidebar(props: { route: Route; onNewBoard: () => void }): JSX.Element {
  const app = useApp();
  const [filter, setFilter] = React.useState('');
  const [showArchived, setShowArchived] = React.useState(false);
  const [closedFolders, setClosedFolders] = React.useState<string[]>([]);

  const active = (b: IBoard): boolean => props.route.page === 'board' && props.route.boardId === b.id;
  const q = filter.trim().toLowerCase();
  const visible = app.boards.filter(b => (showArchived ? b.archived : !b.archived) && (!q || b.title.toLowerCase().indexOf(q) >= 0));
  const favourites = app.boards.filter(b => !b.archived && app.prefs.favourites.indexOf(b.id) >= 0);
  const folders = visible.map(b => b.folder).filter((f, i, a) => f && a.indexOf(f) === i).sort();
  const loose = visible.filter(b => !b.folder);

  const boardLink = (b: IBoard, indent: boolean): JSX.Element => (
    <button key={b.id} type="button" className={`${styles.navItem} ${active(b) ? styles.navItemActive : ''}`} style={indent ? { paddingLeft: 26 } : undefined}
      onClick={() => app.navigate({ page: 'board', boardId: b.id })} aria-current={active(b) ? 'page' : undefined}>
      <span className={styles.boardDot} style={{ background: b.color }} />
      <span className={styles.navLabel}>{b.title}</span>
      {b.privacy === 'Private' && <Icon iconName="Lock" style={{ fontSize: 11, color: 'var(--wb-subtle)' }} aria-label="Private" />}
    </button>
  );

  return (
    <nav className={styles.sidebar} aria-label="Work Boards">
      <div className={styles.brand}>
        <span className={styles.brandMark}><Icon iconName="BacklogBoard" /></span>
        <span className={styles.navLabel}>{app.siteTitle}</span>
      </div>
      <button type="button" className={`${styles.navItem} ${props.route.page === 'home' ? styles.navItemActive : ''}`} onClick={() => app.navigate({ page: 'home' })}>
        <Icon iconName="Home" /> <span className={styles.navLabel}>Home</span>
      </button>
      <button type="button" className={`${styles.navItem} ${props.route.page === 'mywork' ? styles.navItemActive : ''}`} onClick={() => app.navigate({ page: 'mywork' })}>
        <Icon iconName="CheckList" /> <span className={styles.navLabel}>My Work</span>
      </button>

      {favourites.length > 0 && (
        <>
          <div className={styles.navSection}>Favourites</div>
          {favourites.map(b => boardLink(b, false))}
        </>
      )}

      <div className={styles.navSection}>
        <span>{showArchived ? 'Archived boards' : 'Boards'}</span>
        <button type="button" className={styles.iconBtn} style={{ width: 24, height: 24 }} onClick={props.onNewBoard} aria-label="New board" title="New board">
          <Icon iconName="Add" />
        </button>
      </div>
      {app.boards.length > 6 && (
        <SearchBox className={styles.sidebarSearch} placeholder="Find a board" value={filter} onChange={(_, v) => setFilter(v || '')} underlined />
      )}
      {folders.map(f => {
        const open = closedFolders.indexOf(f) < 0;
        return (
          <React.Fragment key={f}>
            <button type="button" className={styles.navItem} aria-expanded={open}
              onClick={() => setClosedFolders(c => (open ? c.concat([f]) : c.filter(x => x !== f)))}>
              <Icon iconName={open ? 'ChevronDown' : 'ChevronRight'} style={{ fontSize: 10 }} />
              <Icon iconName="FabricFolder" />
              <span className={styles.navLabel}>{f}</span>
            </button>
            {open && visible.filter(b => b.folder === f).map(b => boardLink(b, true))}
          </React.Fragment>
        );
      })}
      {loose.map(b => boardLink(b, false))}
      {visible.length === 0 && <span className={`${styles.small} ${styles.muted}`} style={{ padding: '4px 8px' }}>{q ? 'No boards match.' : showArchived ? 'No archived boards.' : 'No boards yet.'}</span>}
      <button type="button" className={styles.navItem} style={{ color: 'var(--wb-primary)' }} onClick={props.onNewBoard}>
        <Icon iconName="Add" /> <span className={styles.navLabel}>New board</span>
      </button>
      {app.boards.some(b => b.archived) && (
        <button type="button" className={`${styles.navItem} ${styles.small}`} style={{ marginTop: 'auto', color: 'var(--wb-subtle)' }} onClick={() => setShowArchived(s => !s)}>
          <Icon iconName="Archive" /> <span className={styles.navLabel}>{showArchived ? 'Show active boards' : 'Show archived boards'}</span>
        </button>
      )}
    </nav>
  );
}
