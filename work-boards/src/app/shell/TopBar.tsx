import * as React from 'react';
import { Icon, TooltipHost } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { useApp } from '../AppContext';
import { Route } from '../router';
import { crumbsFor } from '../../engine/navigation';

/** The bar above every page: menu (phones), where you are, search, new board and help. */
export function TopBar(props: {
  route: Route;
  onMenu: () => void;
  onSearch: () => void;
  onNewBoard: () => void;
  onHelp: () => void;
}): JSX.Element {
  const app = useApp();
  const crumbs = crumbsFor(props.route, app.boards);
  return (
    <div className={styles.topBar}>
      <button type="button" className={`${styles.iconBtn} ${styles.menuBtn}`} onClick={props.onMenu} aria-label="Open menu">
        <Icon iconName="GlobalNavButton" />
      </button>
      <nav aria-label="You are here" className={styles.crumbs}>
        <ol>
          {crumbs.map((c, i) => {
            const last = i === crumbs.length - 1;
            return (
              <li key={i}>
                {i > 0 && <Icon iconName="ChevronRight" className={styles.crumbSep} />}
                {c.route && !last
                  ? <button type="button" className={styles.crumbLink} onClick={() => app.navigate(c.route as Route)}>{c.label}</button>
                  : <span className={last ? styles.crumbCurrent : styles.muted} aria-current={last ? 'page' : undefined}>{c.label}</span>}
              </li>
            );
          })}
        </ol>
      </nav>
      <span className={styles.spacer} />
      <button type="button" className={styles.searchTrigger} onClick={props.onSearch} aria-label="Search or jump to a board or page (Ctrl+K)">
        <Icon iconName="Search" />
        <span className={styles.searchTriggerText}>Search or jump to…</span>
        <kbd className={styles.kbd}>Ctrl K</kbd>
      </button>
      <TooltipHost content="Create a new board">
        <button type="button" className={styles.newBtn} onClick={props.onNewBoard}>
          <Icon iconName="Add" /> <span className={styles.newBtnText}>New board</span>
        </button>
      </TooltipHost>
      <TooltipHost content="Help and keyboard shortcuts (?)">
        <button type="button" className={styles.iconBtn} onClick={props.onHelp} aria-label="Help">
          <Icon iconName="Help" />
        </button>
      </TooltipHost>
    </div>
  );
}
