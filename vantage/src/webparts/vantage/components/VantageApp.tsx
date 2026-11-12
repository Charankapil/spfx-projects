import * as React from 'react';
import {
  DefaultButton, Dropdown, IconButton, type ITheme, MessageBar, MessageBarType, Panel, PanelType, Pivot, PivotItem,
  PrimaryButton, ProgressIndicator, Spinner, SpinnerSize, ThemeProvider
} from '@fluentui/react';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import styles from './VantageApp.module.scss';
import { AppContext, canAdmin, Card, EmptyState, errorText, type IAppContext, relativeTime, themeVars } from './common';
import { AccessTab } from './AccessTab';
import { ArchiveTab } from './ArchiveTab';
import { BoardsTab } from './BoardsTab';
import { OverviewTab } from './OverviewTab';
import { WorkspacesTab } from './WorkspacesTab';
import { ArchiveService } from '../services/ArchiveService';
import { PermissionService } from '../services/PermissionService';
import { type ISetupState, ProvisioningService } from '../services/ProvisioningService';
import { SharePointService } from '../services/SharePointService';
import { WorkService } from '../services/WorkService';
import type { ICurrentUser, INotification, IWorkspaceSummary, WorkspaceRole } from '../models/Models';
import { VANTAGE_LOGO } from './logo';

export interface IVantageAppProps {
  sp: SharePointService;
  webServerRelativeUrl: string;
  user: ICurrentUser;
  /** Can manage permissions on the site (site owner). */
  siteOwner: boolean;
  canManageLists: boolean;
  siteCanEdit: boolean;
  theme: IReadonlyTheme | undefined;
  title: string;
  subtitle: string;
}

const LAST_WS_KEY = 'vantage:lastWorkspace';
const readLast = (): string => { try { return window.localStorage.getItem(LAST_WS_KEY) || ''; } catch { return ''; } };
const writeLast = (id: string): void => { try { window.localStorage.setItem(LAST_WS_KEY, id); } catch { /* storage unavailable */ } };

export const VantageApp: React.FC<IVantageAppProps> = props => {
  const { sp, user, siteOwner, canManageLists, siteCanEdit, theme } = props;

  const provisioner = React.useMemo(() => new ProvisioningService(sp), [sp]);
  const work = React.useMemo(() => new WorkService(sp, user), [sp, user]);
  const archiver = React.useMemo(() => new ArchiveService(sp), [sp]);
  const permissions = React.useMemo(() => new PermissionService(sp), [sp]);

  const [loading, setLoading] = React.useState(true);
  const [setup, setSetup] = React.useState<ISetupState | undefined>(undefined);
  const [busyLabel, setBusyLabel] = React.useState('');
  const [notice, setNotice] = React.useState<{ type: MessageBarType; text: string } | undefined>(undefined);
  const [workspaces, setWorkspaces] = React.useState<IWorkspaceSummary[]>([]);
  const [roles, setRoles] = React.useState<Record<string, WorkspaceRole>>({});
  const [workspaceId, setWorkspaceId] = React.useState('');
  const [tab, setTab] = React.useState('overview');
  const [refreshToken, setRefreshToken] = React.useState(0);
  const [archiveRunning, setArchiveRunning] = React.useState(false);
  const [notifications, setNotifications] = React.useState<INotification[]>([]);
  const [notificationsOpen, setNotificationsOpen] = React.useState(false);

  const initialized = !!setup?.initialized;
  const accessible = workspaces.filter(w => roles[w.id] && roles[w.id] !== 'none');
  const workspace = accessible.find(w => w.id === workspaceId);
  const role: WorkspaceRole = workspace ? roles[workspace.id] : 'none';

  const notify = React.useCallback((type: MessageBarType, text: string) => setNotice({ type, text }), []);
  const perform = React.useCallback(async (label: string, fn: () => Promise<void>): Promise<boolean> => {
    setBusyLabel(label);
    setNotice(undefined);
    try {
      await fn();
      return true;
    } catch (e) {
      setNotice({ type: MessageBarType.error, text: errorText(e) });
      return false;
    } finally {
      setBusyLabel('');
    }
  }, []);
  const bumpRefresh = React.useCallback(() => setRefreshToken(t => t + 1), []);

  const loadWorkspaces = React.useCallback(async (preferId?: string): Promise<void> => {
    const list = await provisioner.listWorkspaces();
    const r = await permissions.resolveRoles(list, siteOwner, siteCanEdit);
    setWorkspaces(list);
    setRoles(r);
    const usable = list.filter(w => r[w.id] !== 'none');
    setWorkspaceId(current => {
      const want = preferId || current || readLast();
      return usable.some(w => w.id === want) ? want : (usable[0] ? usable[0].id : '');
    });
  }, [provisioner, permissions, siteOwner, siteCanEdit]);

  // --- Initial load --------------------------------------------------------
  React.useEffect(() => {
    let live = true;
    (async () => {
      try {
        const state = await provisioner.getSetupState();
        if (!live) return;
        setSetup(state);
        if (state.initialized) await loadWorkspaces();
      } catch (e) {
        if (live) setNotice({ type: MessageBarType.error, text: `Could not load Vantage: ${errorText(e)}` });
      } finally {
        if (live) setLoading(false);
      }
    })().catch(() => undefined);
    return () => { live = false; };
  }, [provisioner, loadWorkspaces]);

  React.useEffect(() => { if (workspaceId) writeLast(workspaceId); }, [workspaceId]);

  // --- Notifications for the current workspace ------------------------------
  React.useEffect(() => {
    if (!workspace) { setNotifications([]); return; }
    let live = true;
    work.listNotifications(workspace).then(n => { if (live) setNotifications(n); }).catch(() => { if (live) setNotifications([]); });
    return () => { live = false; };
  }, [workspace, work, refreshToken]);
  const unread = notifications.filter(n => !n.NotificationRead);

  // Keep the selected tab valid for the user's role.
  React.useEffect(() => {
    if ((tab === 'archive' || tab === 'access') && !canAdmin(role)) setTab('overview');
  }, [role, tab]);

  const initialize = (): Promise<boolean> => perform('Setting up Vantage…', async () => {
    await provisioner.initialize(step => setBusyLabel(step));
    const state = await provisioner.getSetupState();
    setSetup(state);
    await loadWorkspaces();
    setNotice({ type: MessageBarType.success, text: setup && setup.upgradeRequired ? 'Vantage was upgraded to the latest version.' : 'Vantage is set up on this site. Create your first workspace to get started.' });
  });

  const ctx: IAppContext = {
    sp, provisioner, work, archiver, permissions, user, webServerRelativeUrl: props.webServerRelativeUrl, siteOwner,
    workspace, role, notify, perform, busy: !!busyLabel, refreshToken, bumpRefresh, themeStyle: themeVars(theme)
  };

  // --- Render ------------------------------------------------------------------
  const header = (
    <header className={styles.header}>
      <div className={styles.brand}>
        <img className={styles.logo} src={VANTAGE_LOGO} alt="" aria-hidden="true" />
        <div>
          <h2 className={styles.title}>{props.title}</h2>
          <p className={styles.subtitle}>{props.subtitle}</p>
        </div>
      </div>
      {initialized && accessible.length > 0 && (
        <div className={styles.headerActions}>
          <Dropdown className={styles.workspacePicker} ariaLabel="Current workspace" selectedKey={workspaceId || undefined}
            disabled={!!busyLabel || archiveRunning}
            options={accessible.map(w => ({ key: w.id, text: w.title }))} onChange={(_, o) => setWorkspaceId(String(o?.key))} />
          <div className={styles.bellWrap}>
            <IconButton iconProps={{ iconName: 'Ringer' }} title="Notifications" ariaLabel={`Notifications, ${unread.length} unread`}
              onClick={() => setNotificationsOpen(true)} />
            {unread.length > 0 && <span className={styles.bellCount} aria-hidden="true">{unread.length > 9 ? '9+' : unread.length}</span>}
          </div>
        </div>
      )}
    </header>
  );

  let content: React.ReactNode;
  if (loading) {
    content = <div className={styles.emptyState}><Spinner size={SpinnerSize.large} label="Loading Vantage…" /></div>;
  } else if (!initialized) {
    const upgrade = !!setup && setup.upgradeRequired;
    const partial = !!setup && !upgrade && setup.missingLists.length > 0 && setup.missingLists.length < 7;
    content = (
      <Card className={styles.setupCard}>
        <EmptyState
          title={upgrade ? 'Upgrade Vantage on this site' : 'Set up Vantage on this site'}
          text={upgrade
            ? 'This site was set up by an earlier version of Vantage. The upgrade adds boards, collaboration and restore support to the existing lists. Your data is kept.'
            : 'A one-time setup creates the Vantage control lists (settings, workspaces, boards, archive policies, jobs, audit log and locks) in this site.'}
          action={<>
            {partial && <MessageBar messageBarType={MessageBarType.warning}>A previous setup did not finish. Missing: {setup?.missingLists.join(', ')}. Run setup again to repair it.</MessageBar>}
            {!canManageLists && <MessageBar messageBarType={MessageBarType.info}>You need permission to manage lists on this site (Site Owner) to run setup. Ask a site owner to open this page.</MessageBar>}
            <div className={styles.actions} style={{ justifyContent: 'center' }}>
              <PrimaryButton text={upgrade ? 'Upgrade now' : partial ? 'Repair setup' : 'Set up Vantage'} iconProps={{ iconName: 'Settings' }}
                onClick={() => { initialize().catch(() => undefined); }} disabled={!!busyLabel || !canManageLists} />
            </div>
          </>} />
      </Card>
    );
  } else if (!accessible.length) {
    content = (
      <div className={styles.section}>
        {canManageLists
          ? <WorkspacesTab workspaces={workspaces} roles={roles} canCreate onSelect={w => setWorkspaceId(w.id)}
              onCreated={w => { loadWorkspaces(w.id).then(() => setTab('boards')).catch(() => undefined); }} />
          : <Card><EmptyState title="No workspaces available" text="You are not a member of any Vantage workspace on this site yet. Ask a workspace administrator to add you." /></Card>}
      </div>
    );
  } else {
    content = (
      <Pivot aria-label="Vantage sections" selectedKey={tab} onLinkClick={item => item && setTab(item.props.itemKey as string)} overflowBehavior="menu">
        <PivotItem headerText="Overview" itemKey="overview" itemIcon="ViewDashboard">
          <div className={styles.section}><OverviewTab onGoToBoards={() => setTab('boards')} /></div>
        </PivotItem>
        <PivotItem headerText="Boards & tasks" itemKey="boards" itemIcon="TaskManager">
          <div className={styles.section}><BoardsTab /></div>
        </PivotItem>
        {canAdmin(role) && (
          <PivotItem headerText="Archive & retention" itemKey="archive" itemIcon="Archive">
            <div className={styles.section}><ArchiveTab onRunningChange={setArchiveRunning} /></div>
          </PivotItem>
        )}
        <PivotItem headerText="Workspaces" itemKey="workspaces" itemIcon="Org">
          <div className={styles.section}>
            <WorkspacesTab workspaces={workspaces} roles={roles} canCreate={canManageLists} onSelect={w => setWorkspaceId(w.id)}
              onCreated={w => { loadWorkspaces(w.id).then(() => setTab('boards')).catch(() => undefined); }} />
          </div>
        </PivotItem>
        {canAdmin(role) && (
          <PivotItem headerText="Access" itemKey="access" itemIcon="People">
            <div className={styles.section}><AccessTab onSecured={() => { loadWorkspaces(workspaceId).catch(() => undefined); }} /></div>
          </PivotItem>
        )}
      </Pivot>
    );
  }

  return (
    <ThemeProvider theme={theme as ITheme | undefined} applyTo="none">
      <AppContext.Provider value={ctx}>
        <section className={styles.root} style={ctx.themeStyle}>
          {header}
          <div className={styles.body}>
            {busyLabel && <div className={styles.messages} aria-live="polite"><ProgressIndicator label={busyLabel} /></div>}
            {notice && (
              <div className={styles.messages}>
                <MessageBar messageBarType={notice.type} isMultiline onDismiss={() => setNotice(undefined)} dismissButtonAriaLabel="Close">{notice.text}</MessageBar>
              </div>
            )}
            {initialized && workspace && !workspace.secured && canAdmin(role) && tab !== 'access' && (
              <div className={styles.messages}>
                <MessageBar messageBarType={MessageBarType.warning} actions={<DefaultButton text="Review access" onClick={() => setTab('access')} />}>
                  “{workspace.title}” uses the site’s permissions, so everyone on this site can see its tasks.
                </MessageBar>
              </div>
            )}
            {content}
          </div>
        </section>

        <Panel isOpen={notificationsOpen} onDismiss={() => setNotificationsOpen(false)} type={PanelType.smallFixedFar}
          headerText="Notifications" closeButtonAriaLabel="Close" isLightDismiss>
          <div style={ctx.themeStyle}>
          {unread.length > 0 && (
            <DefaultButton text="Mark all as read" className={styles.markAll} onClick={() => {
              if (!workspace) return;
              perform('Updating…', async () => {
                for (const n of unread) await work.markNotificationRead(workspace, n.Id);
                setNotifications(notifications.map(n => ({ ...n, NotificationRead: true })));
              }).catch(() => undefined);
            }} />
          )}
          <ul className={styles.notificationList}>
            {notifications.map(n => (
              <li key={n.Id} className={n.NotificationRead ? undefined : styles.unread}>
                <div>{n.Title}</div>
                <div className={styles.muted}>{n.ActorName ? `${n.ActorName} · ` : ''}{relativeTime(n.Created)}</div>
              </li>
            ))}
            {!notifications.length && <li className={styles.muted}>You’re all caught up.</li>}
          </ul>
          </div>
        </Panel>
      </AppContext.Provider>
    </ThemeProvider>
  );
};
