import * as React from 'react';
import { Icon, SearchBox } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { useApp } from '../AppContext';
import { Route } from '../router';
import { IBoard } from '../../models/types';
import { Logo } from '../common/Logo';
import { PAGES } from '../../engine/navigation';

export function Sidebar(props: {
  route: Route; version?: string; onNewBoard: () => void;
  /** Narrow, icons-only menu (desktop). */
  collapsed: boolean; onToggleCollapsed: () => void;
  /** Phones: the menu slides in over the page. */
  open: boolean; onClose: () => void;
}): JSX.Element {
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

  const go = (r: Route): void => {
    app.navigate(r);
    props.onClose();
  };

  const boardLink = (b: IBoard, indent: boolean): JSX.Element => (
    <button key={b.id} type="button" className={`${styles.navItem} ${active(b) ? styles.navItemActive : ''}`} style={indent && !props.collapsed ? { paddingLeft: 26 } : undefined}
      title={props.collapsed ? b.title : undefined}
      onClick={() => go({ page: 'board', boardId: b.id })} aria-current={active(b) ? 'page' : undefined}>
      <span className={styles.boardDot} style={{ background: b.color }} />
      <span className={styles.navLabel}>{b.title}</span>
      {b.privacy === 'Private' && <Icon iconName="Lock" style={{ fontSize: 11, color: 'var(--wb-subtle)' }} aria-label="Private" />}
    </button>
  );

  return (
    <>
    {props.open && <div className={styles.backdrop} onClick={props.onClose} aria-hidden="true" />}
    <nav className={`${styles.sidebar} ${props.collapsed ? styles.sidebarCollapsed : ''} ${props.open ? styles.sidebarOpen : ''}`} aria-label="Work Boards">
      <div className={styles.brand}>
        <Logo size={24} />
        <span className={styles.navLabel}>{app.siteTitle}</span>
        <button type="button" className={`${styles.iconBtn} ${styles.collapseBtn}`} onClick={props.onToggleCollapsed}
          aria-label={props.collapsed ? 'Expand menu' : 'Collapse menu'} title={props.collapsed ? 'Expand menu' : 'Collapse menu'}>
          <Icon iconName={props.collapsed ? 'DoubleChevronRight' : 'DoubleChevronLeft'} />
        </button>
        <button type="button" className={`${styles.iconBtn} ${styles.closeMenuBtn}`} onClick={props.onClose} aria-label="Close menu">
          <Icon iconName="Cancel" />
        </button>
      </div>
      <div className={styles.navSection}>For me</div>
      {PAGES.map(p => {
        const on = props.route.page === p.route.page;
        return (
          <button key={p.label} type="button" className={`${styles.navItem} ${styles.navItemTall} ${on ? styles.navItemActive : ''}`}
            onClick={() => go(p.route)} aria-current={on ? 'page' : undefined} title={props.collapsed ? `${p.label}: ${p.hint}` : p.hint}>
            <Icon iconName={p.icon} />
            <span className={styles.navText}>
              <span className={styles.navLabel}>{p.label}</span>
              <span className={styles.navHint}>{p.hint}</span>
            </span>
          </button>
        );
      })}

      {favourites.length > 0 && (
        <>
          <div className={styles.navSection}>Favourites</div>
          {favourites.map(b => boardLink(b, false))}
        </>
      )}

      <div className={styles.navSection}>
        <span>{showArchived ? 'Archived boards' : `Boards (${visible.length})`}</span>
        <button type="button" className={styles.iconBtn} style={{ width: 24, height: 24 }} onClick={props.onNewBoard} aria-label="New board" title="New board">
          <Icon iconName="Add" />
        </button>
      </div>
      {app.boards.length > 6 && !props.collapsed && (
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
      {props.version && <span className={`${styles.small} ${styles.muted} ${styles.navVersion}`}>Work Boards v{props.version}</span>}
    </nav>
    </>
  );
}
