import * as React from 'react';
import { Modal, Icon } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { useApp } from '../AppContext';
import { Route } from '../router';
import { PAGES, rankMatch } from '../../engine/navigation';

interface IEntry {
  key: string;
  label: string;
  hint: string;
  icon?: string;
  color?: string;
  route: Route;
}

/** "Search or jump to…": find any board or page by typing part of its name. Ctrl+K opens it. */
export function QuickSwitcher(props: { onClose: () => void; onNewBoard: () => void }): JSX.Element {
  const app = useApp();
  const [query, setQuery] = React.useState('');
  const [active, setActive] = React.useState(0);

  const pageEntries: IEntry[] = PAGES.map(p => ({ key: 'p:' + p.label, label: p.label, hint: p.hint, icon: p.icon, route: p.route }));
  const entries: IEntry[] = pageEntries
    .concat(app.boards.filter(b => !b.archived).map((b): IEntry => ({
      key: 'b:' + b.id,
      label: b.title,
      hint: (b.folder ? b.folder + ' · ' : '') + 'Board' + (b.privacy === 'Private' ? ' · Private' : ''),
      color: b.color,
      route: { page: 'board', boardId: b.id }
    })));
  const results = entries
    .map(e => ({ e, score: rankMatch(e.label, query) }))
    .filter(x => x.score > 0)
    .sort((a, b) => b.score - a.score || a.e.label.localeCompare(b.e.label))
    .slice(0, 12)
    .map(x => x.e);
  const showNew = query.trim().length > 0;
  const count = results.length + (showNew ? 1 : 0);

  const choose = (i: number): void => {
    if (i < results.length) {
      app.navigate(results[i].route);
      props.onClose();
    } else if (showNew) {
      props.onClose();
      props.onNewBoard();
    }
  };

  return (
    <Modal isOpen onDismiss={props.onClose} isBlocking={false} styles={{ main: { width: 560, maxWidth: '92vw', borderRadius: 10, top: '12vh', position: 'absolute' } }}
      titleAriaId="wb-switcher-title">
      <div className={styles.switcher}>
        <h2 id="wb-switcher-title" className={styles.visuallyHidden}>Search or jump to</h2>
        <div className={styles.switcherInput}>
          <Icon iconName="Search" />
          <input
            autoFocus
            value={query}
            placeholder="Type a board or page name…"
            aria-label="Search boards and pages"
            aria-controls="wb-switcher-results"
            aria-activedescendant={count ? `wb-sw-${active}` : undefined}
            onChange={e => { setQuery(e.target.value); setActive(0); }}
            onKeyDown={e => {
              if (e.key === 'ArrowDown') { e.preventDefault(); setActive(a => (count ? (a + 1) % count : 0)); }
              if (e.key === 'ArrowUp') { e.preventDefault(); setActive(a => (count ? (a - 1 + count) % count : 0)); }
              if (e.key === 'Enter') { e.preventDefault(); choose(active); }
              if (e.key === 'Escape') { props.onClose(); }
            }}
          />
          <kbd className={styles.kbd}>Esc</kbd>
        </div>
        <div id="wb-switcher-results" role="listbox" className={styles.switcherList}>
          {results.length === 0 && !showNew && <p className={styles.muted} style={{ padding: 12, margin: 0 }}>No boards yet.</p>}
          {results.map((r, i) => (
            <button key={r.key} id={`wb-sw-${i}`} type="button" role="option" aria-selected={i === active}
              className={`${styles.switcherItem} ${i === active ? styles.switcherItemActive : ''}`}
              onMouseEnter={() => setActive(i)} onClick={() => choose(i)}>
              {r.icon ? <Icon iconName={r.icon} /> : <span className={styles.boardDot} style={{ background: r.color, width: 12, height: 12 }} />}
              <span className={styles.switcherLabel}>{r.label}</span>
              <span className={`${styles.small} ${styles.muted} ${styles.ellipsis}`}>{r.hint}</span>
            </button>
          ))}
          {showNew && (
            <button id={`wb-sw-${results.length}`} type="button" role="option" aria-selected={active === results.length}
              className={`${styles.switcherItem} ${active === results.length ? styles.switcherItemActive : ''}`}
              onMouseEnter={() => setActive(results.length)} onClick={() => choose(results.length)}>
              <Icon iconName="Add" />
              <span className={styles.switcherLabel}>Create a new board…</span>
            </button>
          )}
        </div>
        <div className={`${styles.switcherFoot} ${styles.small} ${styles.muted}`}>
          <span><kbd className={styles.kbd}>↑</kbd> <kbd className={styles.kbd}>↓</kbd> to move</span>
          <span><kbd className={styles.kbd}>Enter</kbd> to open</span>
          <span><kbd className={styles.kbd}>Ctrl</kbd> + <kbd className={styles.kbd}>K</kbd> opens this from anywhere</span>
        </div>
      </div>
    </Modal>
  );
}
