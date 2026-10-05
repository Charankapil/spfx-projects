import * as React from 'react';
import {
  Dialog,
  DialogFooter,
  DefaultButton,
  Icon,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
  PrimaryButton,
  TextField
} from '@fluentui/react';
import { SPClient, trimSlash } from '../core/SPClient';
import { IActionLogEntry, ITarget, ViewKey } from '../models';
import { AdminApi } from '../services/AdminApi';
import { GrowthStore } from '../services/GrowthStore';
import { SearchApi } from '../services/SearchApi';
import { TenantStore } from '../services/TenantStore';
import styles from './AdminCenter.module.scss';
import { Logo } from './shared/ui';
import { AdminContext, IAdminContext, IConfirmOptions } from './shared/context';
import { ActivityView } from './views/ActivityView';
import { ContentView } from './views/ContentView';
import { GrowthView } from './views/GrowthView';
import { HealthView } from './views/HealthView';
import { OverviewView } from './views/OverviewView';
import { PeopleView } from './views/PeopleView';
import { RecycleView } from './views/RecycleView';
import { SettingsView } from './views/SettingsView';
import { SitesView } from './views/SitesView';
import { StorageView } from './views/StorageView';

export interface IAdminCenterProps {
  client: SPClient;
  homeWebUrl: string;
  homeSiteUrl: string;
  homeTitle: string;
  heading: string;
  currentUser: string;
}

const NAV: Array<{ key: ViewKey; text: string; icon: string }> = [
  { key: 'overview', text: 'Dashboard', icon: 'ViewDashboard' },
  { key: 'sites', text: 'Sites', icon: 'Globe' },
  { key: 'people', text: 'People & permissions', icon: 'Permissions' },
  { key: 'content', text: 'Lists & libraries', icon: 'Library' },
  { key: 'storage', text: 'Storage insights', icon: 'Database' },
  { key: 'growth', text: 'Storage growth', icon: 'AreaChart' },
  { key: 'recycle', text: 'Recycle bin', icon: 'RecycleBin' },
  { key: 'activity', text: 'Activity', icon: 'History' },
  { key: 'health', text: 'Health check', icon: 'HealthSolid' },
  { key: 'settings', text: 'Settings & log', icon: 'Settings' }
];

interface IToast {
  id: number;
  message: string;
  kind: 'success' | 'error' | 'info';
}

const RECENT_KEY = 'spAdminCenter.recentTargets';

function readRecent(): ITarget[] {
  try {
    const raw = window.localStorage.getItem(RECENT_KEY);
    return raw ? (JSON.parse(raw) as ITarget[]).slice(0, 6) : [];
  } catch {
    return [];
  }
}

function writeRecent(t: ITarget): void {
  try {
    const list = [t].concat(readRecent().filter((r) => r.webUrl !== t.webUrl)).slice(0, 6);
    window.localStorage.setItem(RECENT_KEY, JSON.stringify(list));
  } catch {
    /* storage can be unavailable (private mode); recents are a convenience */
  }
}

export const AdminCenter: React.FC<IAdminCenterProps> = ({ client, homeWebUrl, homeSiteUrl, homeTitle, heading, currentUser }) => {
  const api = React.useMemo(() => new AdminApi(client), [client]);
  const search = React.useMemo(() => new SearchApi(client), [client]);
  const growth = React.useMemo(() => new GrowthStore(client, api, trimSlash(homeWebUrl)), [client, api, homeWebUrl]);
  const tenant = React.useMemo(() => new TenantStore(client, trimSlash(homeWebUrl)), [client, homeWebUrl]);
  const home: ITarget = React.useMemo(() => ({ webUrl: trimSlash(homeWebUrl), siteUrl: trimSlash(homeSiteUrl), title: homeTitle }), [homeWebUrl, homeSiteUrl, homeTitle]);

  const [view, setView] = React.useState<ViewKey>('overview');
  const [target, setTargetState] = React.useState<ITarget>(home);
  const [reloadToken, setReloadToken] = React.useState(0);
  const [actionLog, setActionLog] = React.useState<IActionLogEntry[]>([]);
  const [toasts, setToasts] = React.useState<IToast[]>([]);
  const [badge, setBadge] = React.useState(0);
  const [growthBadge, setGrowthBadge] = React.useState(0);
  const [stats, setStats] = React.useState({ ...client.stats });
  const [now, setNow] = React.useState(Date.now());
  const [switchOpen, setSwitchOpen] = React.useState(false);
  const [confirmState, setConfirmState] = React.useState<{ o: IConfirmOptions; resolve: (v: boolean) => void } | undefined>();
  const toastId = React.useRef(0);

  // Live request meter (no polling: it only re-renders when the client reports a change).
  React.useEffect(() => client.subscribe(() => setStats({ ...client.stats })), [client]);
  const paused = stats.pausedUntil > now;
  React.useEffect(() => {
    if (!paused) {
      return undefined;
    }
    const t = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(t);
  }, [paused]);
  React.useEffect(() => {
    setNow(Date.now());
  }, [stats.pausedUntil]);

  const notify = React.useCallback((message: string, kind: 'success' | 'error' | 'info' = 'info') => {
    const id = ++toastId.current;
    setToasts((t) => t.concat({ id, message, kind }).slice(-4));
    setTimeout(() => setToasts((t) => t.filter((x) => x.id !== id)), kind === 'error' ? 9000 : 4500);
  }, []);

  const setTarget = React.useCallback(
    (t: ITarget) => {
      client.cancelPending();
      client.clearCache();
      setBadge(0);
      setTargetState(t);
      writeRecent(t);
      setView((v) => (v === 'sites' ? 'overview' : v));
    },
    [client]
  );

  const ctx: IAdminContext = {
    client,
    api,
    search,
    growth,
    tenant,
    homeWebUrl: home.webUrl,
    currentUser: (currentUser || '').toLowerCase(),
    target,
    setTarget,
    openView: setView,
    confirm: (o) => new Promise<boolean>((resolve) => setConfirmState({ o, resolve })),
    log: (action, tgt, ok, message) => setActionLog((l) => l.concat({ at: new Date(), action, target: tgt, ok, message })),
    actionLog,
    reloadToken,
    reloadAll: () => {
      client.clearCache();
      setReloadToken((n) => n + 1);
    },
    notify,
    setHealthBadge: setBadge,
    setGrowthBadge
  };

  const resolveConfirm = (v: boolean): void => {
    if (confirmState) {
      confirmState.resolve(v);
    }
    setConfirmState(undefined);
  };

  const renderView = (): React.ReactNode => {
    switch (view) {
      case 'sites':
        return <SitesView />;
      case 'people':
        return <PeopleView />;
      case 'content':
        return <ContentView />;
      case 'storage':
        return <StorageView />;
      case 'growth':
        return <GrowthView />;
      case 'recycle':
        return <RecycleView />;
      case 'activity':
        return <ActivityView />;
      case 'health':
        return <HealthView />;
      case 'settings':
        return <SettingsView stats={stats} />;
      default:
        return <OverviewView />;
    }
  };

  const secondsLeft = Math.max(0, Math.ceil((stats.pausedUntil - now) / 1000));

  return (
    <AdminContext.Provider value={ctx}>
      <div className={styles.root}>
        <header className={styles.header}>
          <div className={styles.brand}>
            <Logo size={44} />
            <div>
            <h1 className={styles.headerTitle}>{heading}</h1>
            <p className={styles.headerSub}>Administer SharePoint sites as yourself: no app registration, nothing runs in the background.</p>
            </div>
          </div>
          <div className={styles.headerRight}>
            {paused ? (
              <span className={`${styles.meter} ${styles.meterPaused}`} role="status">
                SharePoint asked us to slow down. Resuming in {secondsLeft}s
              </span>
            ) : (
              <span className={styles.meter} title="Requests sent to SharePoint in this session">
                {stats.queued > 0 ? 'Working… ' : ''}
                {stats.requests} requests · {stats.cacheHits} cached
              </span>
            )}
            <span className={styles.targetChip} title={target.webUrl}>
              <Icon iconName="Link" aria-hidden="true" />
              <span>
                <strong>{target.title}</strong>
              </span>
            </span>
            <DefaultButton iconProps={{ iconName: 'SwitcherStartEnd' }} onClick={() => setSwitchOpen(true)}>
              Switch site
            </DefaultButton>
          </div>
        </header>

        <div className={styles.body}>
          <nav className={styles.nav} aria-label="Administration sections">
            {NAV.map((n) => (
              <button key={n.key} type="button" className={`${styles.navItem} ${view === n.key ? styles.navItemActive : ''}`} onClick={() => setView(n.key)} aria-current={view === n.key ? 'page' : undefined}>
                <Icon iconName={n.icon} aria-hidden="true" />
                {n.text}
                {n.key === 'growth' && growthBadge > 0 && <span className={styles.navBadge} aria-label={`${growthBadge} sites growing fast`}>{growthBadge}</span>}
                {n.key === 'health' && badge > 0 && <span className={styles.navBadge} aria-label={`${badge} findings`}>{badge}</span>}
              </button>
            ))}
          </nav>
          <main className={styles.content}>
            {/* key forces a clean state when the administered site changes */}
            <React.Fragment key={target.webUrl}>{renderView()}</React.Fragment>
          </main>
        </div>

        <div style={{ position: 'fixed', right: 20, bottom: 20, zIndex: 1000, display: 'flex', flexDirection: 'column', gap: 8, maxWidth: 420 }} aria-live="polite">
          {toasts.map((t) => (
            <MessageBar key={t.id} messageBarType={t.kind === 'success' ? MessageBarType.success : t.kind === 'error' ? MessageBarType.error : MessageBarType.info} onDismiss={() => setToasts((l) => l.filter((x) => x.id !== t.id))} isMultiline>
              {t.message}
            </MessageBar>
          ))}
        </div>

        <Dialog
          hidden={!confirmState}
          onDismiss={() => resolveConfirm(false)}
          dialogContentProps={{ title: confirmState ? confirmState.o.title : '', subText: confirmState ? confirmState.o.message : '' }}
          modalProps={{ isBlocking: true }}
        >
          <DialogFooter>
            <PrimaryButton
              onClick={() => resolveConfirm(true)}
              styles={confirmState && confirmState.o.danger ? { root: { background: '#c4314b', borderColor: '#c4314b' }, rootHovered: { background: '#a4262c', borderColor: '#a4262c' } } : undefined}
            >
              {confirmState ? confirmState.o.confirmText : 'OK'}
            </PrimaryButton>
            <DefaultButton onClick={() => resolveConfirm(false)}>Cancel</DefaultButton>
          </DialogFooter>
        </Dialog>

        <SwitchSitePanel
          open={switchOpen}
          onDismiss={() => setSwitchOpen(false)}
          home={home}
          current={target}
          client={client}
          api={api}
          onPick={(t) => {
            setSwitchOpen(false);
            setTarget(t);
          }}
          openSites={() => {
            setSwitchOpen(false);
            setView('sites');
          }}
        />
      </div>
    </AdminContext.Provider>
  );
};

const SwitchSitePanel: React.FC<{
  open: boolean;
  onDismiss: () => void;
  home: ITarget;
  current: ITarget;
  client: SPClient;
  api: AdminApi;
  onPick: (t: ITarget) => void;
  openSites: () => void;
}> = ({ open, onDismiss, home, current, client, api, onPick, openSites }) => {
  const [url, setUrl] = React.useState('');
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>();
  const recent = open ? readRecent() : [];

  const go = async (): Promise<void> => {
    setError(undefined);
    let u = trimSlash(url.trim());
    if (u.charAt(0) === '/') {
      u = (/^(https:\/\/[^/]+)/i.exec(home.webUrl) || [''])[0] + u;
    }
    try {
      client.assertSameOrigin(u);
    } catch {
      setError('Enter a SharePoint address on this tenant, e.g. ' + home.siteUrl);
      return;
    }
    setBusy(true);
    try {
      const [web, site] = await Promise.all([api.getWeb(u), api.getSite(u)]);
      onPick({ webUrl: web.url, siteUrl: site.url, title: web.title });
      setUrl('');
    } catch (e) {
      setError('Could not open that site: ' + (e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  return (
    <Panel isOpen={open} onDismiss={onDismiss} type={PanelType.smallFixedFar} headerText="Switch site" closeButtonAriaLabel="Close" isLightDismiss>
      <p className={styles.muted}>Every view works on the selected site, so you can administer many sites from here. You need the right permissions on it.</p>
      <TextField label="Site address" placeholder={home.siteUrl} value={url} onChange={(_, v) => setUrl(v || '')} onKeyDown={(e) => e.key === 'Enter' && !busy && url.trim() && go()} errorMessage={error} />
      <div className={styles.actions} style={{ marginTop: 10 }}>
        <PrimaryButton onClick={go} disabled={busy || !url.trim()}>
          {busy ? 'Opening…' : 'Open'}
        </PrimaryButton>
        <DefaultButton onClick={openSites}>Browse all sites</DefaultButton>
      </div>
      <h3 style={{ fontSize: 14, marginTop: 24 }}>Quick picks</h3>
      <ul className={styles.memberList}>
        {[home].concat(recent.filter((r) => r.webUrl !== home.webUrl)).map((t) => (
          <li key={t.webUrl} className={styles.memberRow}>
            <span style={{ minWidth: 0 }}>
              <div>{t.title}</div>
              <div className={styles.muted} style={{ fontSize: 12, overflow: 'hidden', textOverflow: 'ellipsis' }}>
                {t.webUrl}
              </div>
            </span>
            {t.webUrl === current.webUrl ? <span className={styles.pill}>Current</span> : (
              <button type="button" className={styles.link} onClick={() => onPick(t)}>
                Switch
              </button>
            )}
          </li>
        ))}
      </ul>
    </Panel>
  );
};
