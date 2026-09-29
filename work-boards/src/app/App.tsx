import * as React from 'react';
import { MessageBar, MessageBarType, DefaultButton } from '@fluentui/react';
import styles from './WorkBoards.module.scss';
import { AppContext, IAppContext, IServices } from './AppContext';
import { parseRoute, routeToHash, Route } from './router';
import { SpClient } from '../services/SpClient';
import { BoardService } from '../services/BoardService';
import { ItemService } from '../services/ItemService';
import { UpdateService } from '../services/UpdateService';
import { PeopleService } from '../services/PeopleService';
import { PrefsService } from '../services/PrefsService';
import { getSetupStatus, ISetupStatus } from '../services/Provisioner';
import { hasPermission, PermissionKind, IBasePermissions } from '../services/permissions';
import { IBoard, IPerson, IUserPrefs } from '../models/types';
import { Setup } from './shell/Setup';
import { Sidebar } from './shell/Sidebar';
import { Home } from './shell/Home';
import { MyWork } from './shell/MyWork';
import { NewBoardDialog } from './shell/NewBoardDialog';
import { BoardPage } from './board/BoardPage';
import { Loading } from './common/Common';

export interface IAppProps {
  sp: SpClient;
  siteTitle: string;
  /** Fill the page (single-part app page) rather than sit in a section. */
  fullPage: boolean;
  /** Open this board instead of home when the URL has no route (property pane setting). */
  startBoardId?: number;
  /** --wb-* colour variables from the SharePoint theme of the web part's section. */
  themeVars?: { [name: string]: string };
  /** Package version, shown so people can tell which build is running. */
  version?: string;
}

/** Colour variables used by the styles. Copied to <body> so dialogs, panels and callouts (rendered outside the app) get them too. */
const THEME_VARS = ['--wb-primary', '--wb-primary-text', '--wb-bg', '--wb-surface', '--wb-text', '--wb-subtle', '--wb-border',
  '--wb-border-light', '--wb-hover', '--wb-selected', '--wb-danger'];

type Phase = { kind: 'loading' } | { kind: 'setup'; status: ISetupStatus } | { kind: 'ready' } | { kind: 'error'; message: string };

export function App(props: IAppProps): JSX.Element {
  const services = React.useMemo<IServices>(() => {
    const people = new PeopleService(props.sp);
    return {
      sp: props.sp,
      boards: new BoardService(props.sp),
      items: new ItemService(props.sp, people),
      updates: new UpdateService(props.sp),
      people,
      prefs: new PrefsService(props.sp)
    };
  }, [props.sp]);

  const [phase, setPhase] = React.useState<Phase>({ kind: 'loading' });
  const [me, setMe] = React.useState<IPerson | null>(null);
  const [boards, setBoards] = React.useState<IBoard[]>([]);
  const [prefs, setPrefs] = React.useState<IUserPrefs>({ favourites: [], recent: [] });
  const [isSiteOwner, setIsSiteOwner] = React.useState(false);
  const [route, setRoute] = React.useState<Route>(() => {
    const r = parseRoute(window.location.hash);
    return r.page === 'home' && props.startBoardId ? { page: 'board', boardId: props.startBoardId } : r;
  });
  const [newBoardOpen, setNewBoardOpen] = React.useState(false);
  const rootRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    if (!rootRef.current) {
      return;
    }
    const computed = window.getComputedStyle(rootRef.current);
    THEME_VARS.forEach(name => {
      const value = computed.getPropertyValue(name).trim();
      if (value) {
        document.body.style.setProperty(name, value);
      }
    });
  });

  const start = React.useCallback(async (): Promise<void> => {
    setPhase({ kind: 'loading' });
    try {
      const status = await getSetupStatus(props.sp);
      if (status.installedVersion < status.requiredVersion) {
        setPhase({ kind: 'setup', status });
        return;
      }
      const [user, list, p, webPerms] = await Promise.all([
        services.people.currentUser(),
        services.boards.listBoards(),
        services.prefs.load(),
        props.sp.get<IBasePermissions>('web/EffectiveBasePermissions')
      ]);
      setMe(user);
      setBoards(list);
      setPrefs(p);
      setIsSiteOwner(hasPermission(webPerms, PermissionKind.ManagePermissions));
      setPhase({ kind: 'ready' });
    } catch (e) {
      setPhase({ kind: 'error', message: (e as Error).message });
    }
  }, [props.sp]);

  React.useEffect(() => { start().catch(() => undefined); }, [start]);

  React.useEffect(() => {
    const onHash = (): void => setRoute(parseRoute(window.location.hash));
    window.addEventListener('hashchange', onHash);
    return () => window.removeEventListener('hashchange', onHash);
  }, []);

  const rootClass = `${styles.app} ${props.fullPage ? styles.fullPage : ''}`;

  if (phase.kind === 'loading') {
    return <div className={rootClass} ref={rootRef} style={props.themeVars as React.CSSProperties}><div className={styles.main}><Loading label="Opening Work Boards…" /></div></div>;
  }
  if (phase.kind === 'error') {
    return (
      <div className={rootClass} ref={rootRef} style={props.themeVars as React.CSSProperties}>
        <div className={styles.main}>
          <div className={styles.setup}>
            <MessageBar messageBarType={MessageBarType.error}>Work Boards could not load: {phase.message}</MessageBar>
            <div><DefaultButton text="Try again" onClick={() => { start().catch(() => undefined); }} /></div>
          </div>
        </div>
      </div>
    );
  }
  if (phase.kind === 'setup') {
    return (
      <div className={rootClass} ref={rootRef} style={props.themeVars as React.CSSProperties}>
        <div className={styles.main}>
          <Setup sp={props.sp} siteTitle={props.siteTitle} version={props.version} status={phase.status} onDone={() => { start().catch(() => undefined); }} />
        </div>
      </div>
    );
  }

  const ctx: IAppContext = {
    services,
    me: me as IPerson,
    siteTitle: props.siteTitle,
    webUrl: props.sp.webUrl,
    isSiteOwner,
    boards,
    prefs,
    reloadBoards: async () => {
      const list = await services.boards.listBoards();
      setBoards(list);
      return list;
    },
    replaceBoard: b => setBoards(list => (list.some(x => x.id === b.id) ? list.map(x => (x.id === b.id ? b : x)) : list.concat([b]))),
    toggleFavourite: id => {
      services.prefs.toggleFavourite(id).then(setPrefs).catch(() => undefined);
      setPrefs(p => ({ ...p, favourites: p.favourites.indexOf(id) >= 0 ? p.favourites.filter(x => x !== id) : p.favourites.concat([id]) }));
    },
    navigate: r => {
      const hash = routeToHash(r);
      if (window.location.hash === hash) {
        setRoute(r);
      } else {
        window.location.hash = hash;
      }
    }
  };

  return (
    <AppContext.Provider value={ctx}>
      <div className={rootClass} ref={rootRef} style={props.themeVars as React.CSSProperties}>
        <Sidebar route={route} version={props.version} onNewBoard={() => setNewBoardOpen(true)} />
        {route.page === 'home' && <div className={styles.main}><Home onNewBoard={() => setNewBoardOpen(true)} /></div>}
        {route.page === 'mywork' && <div className={styles.main}><MyWork /></div>}
        {route.page === 'board' && <BoardPage key={route.boardId} boardId={route.boardId} view={route.view} itemId={route.itemId} />}
        {newBoardOpen && (
          <NewBoardDialog onClose={() => setNewBoardOpen(false)} onCreated={b => {
            setNewBoardOpen(false);
            ctx.replaceBoard(b);
            ctx.navigate({ page: 'board', boardId: b.id });
          }} />
        )}
      </div>
    </AppContext.Provider>
  );
}
