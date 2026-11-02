import * as React from 'react';
import {
  CommandButton, Dropdown, IDropdownOption, INavLinkGroup, ITheme, MessageBar, MessageBarType, Nav, Text, ThemeProvider
} from '@fluentui/react';
import { AppContext, IAppContext, IDetail, PageId } from './AppContext';
import { IDataProvider } from '../providers/IDataProvider';
import { IEnvironment } from '../models';
import { PAGES } from './sections';
import { DetailPanel } from './components/PageFrame';
import { LensLogo } from './components/LensLogo';
import { DataSourcesPanel } from './DataSourcesPanel';
import {
  AgentsPage, AppsPage, CapacityPage, ConnectorsPage, EnvironmentsPage, FlowsPage, GovernancePage, MakersPage, OverviewPage
} from './pages/PowerPlatformPages';
import { AdoptionPage, HealthTables, IdentityPage, SitesPage, TeamsPage } from './pages/M365Pages';
import { ExportCentrePage } from './pages/ExportCentre';
import { fmtDateTime } from '../logic/format';

export interface IAppProps {
  provider: IDataProvider;
  theme: ITheme;
  isDark: boolean;
  title: string;
  initialPage?: PageId;
}

const PAGE_COMPONENTS: Record<PageId, React.FC> = {
  overview: () => (
    <>
      <OverviewPage />
      <div style={{ marginTop: 12 }}>
        <HealthTables />
      </div>
    </>
  ),
  environments: EnvironmentsPage,
  apps: AppsPage,
  flows: FlowsPage,
  agents: AgentsPage,
  makers: MakersPage,
  connectors: ConnectorsPage,
  capacity: CapacityPage,
  governance: GovernancePage,
  sites: SitesPage,
  teams: TeamsPage,
  adoption: AdoptionPage,
  identity: IdentityPage,
  export: ExportCentrePage
};

class PageErrorBoundary extends React.Component<{ children?: React.ReactNode; resetKey: string }, { error?: Error }> {
  public state: { error?: Error } = {};
  public static getDerivedStateFromError(error: Error): { error: Error } {
    return { error };
  }
  public componentDidUpdate(prev: { resetKey: string }): void {
    if (prev.resetKey !== this.props.resetKey && this.state.error) this.setState({ error: undefined });
  }
  public render(): React.ReactNode {
    if (this.state.error) {
      return (
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          {`This page hit an unexpected error: ${this.state.error.message}. Try Refresh, or open another page.`}
        </MessageBar>
      );
    }
    return this.props.children;
  }
}

export const App: React.FC<IAppProps> = (props) => {
  const { provider, theme, isDark } = props;
  const [page, setPage] = React.useState<PageId>(props.initialPage || 'overview');
  const [envFilter, setEnvFilter] = React.useState<string[]>([]);
  const [detail, setDetail] = React.useState<IDetail | undefined>(undefined);
  const [toast, setToast] = React.useState<{ message: string; kind: 'info' | 'success' | 'error' } | undefined>(undefined);
  const [version, setVersion] = React.useState(0);
  const [refreshKey, setRefreshKey] = React.useState(0);
  const [sourcesOpen, setSourcesOpen] = React.useState(false);
  const [envs, setEnvs] = React.useState<IEnvironment[]>([]);
  const [loadedAt, setLoadedAt] = React.useState<string | undefined>(undefined);
  const [width, setWidth] = React.useState(1200);
  const rootRef = React.useRef<HTMLDivElement>(null);
  const toastTimer = React.useRef<number | undefined>(undefined);

  React.useEffect(() => provider.subscribe(() => setVersion((v) => v + 1)), [provider]);

  React.useEffect(() => {
    let alive = true;
    provider.getInventory().then(
      (inv) => {
        if (!alive) return;
        setEnvs(inv.environments.slice().sort((a, b) => a.displayName.localeCompare(b.displayName)));
        setLoadedAt(inv.loadedAt);
      },
      () => undefined
    );
    return () => {
      alive = false;
    };
  }, [provider, refreshKey]);

  React.useEffect(() => {
    const el = rootRef.current;
    if (!el) return undefined;
    const update = (): void => setWidth(el.getBoundingClientRect().width);
    update();
    const RO = (window as unknown as { ResizeObserver?: typeof ResizeObserver }).ResizeObserver;
    if (RO) {
      const ro = new RO(update);
      ro.observe(el);
      return () => ro.disconnect();
    }
    window.addEventListener('resize', update);
    return () => window.removeEventListener('resize', update);
  }, []);

  React.useEffect(() => () => window.clearTimeout(toastTimer.current), []);

  const compact = width < 900;
  const notify = React.useCallback((message: string, kind: 'info' | 'success' | 'error' = 'info') => {
    setToast({ message, kind });
    window.clearTimeout(toastTimer.current);
    toastTimer.current = window.setTimeout(() => setToast(undefined), kind === 'error' ? 12000 : 6000);
  }, []);

  const ctx: IAppContext = {
    provider, theme, isDark, version, refreshKey, envFilter, setEnvFilter,
    navigate: (p) => {
      setPage(p);
      if (rootRef.current && rootRef.current.scrollIntoView && rootRef.current.getBoundingClientRect().top < 0) rootRef.current.scrollIntoView();
    },
    openDetail: setDetail,
    notify,
    compact
  };

  const refresh = (): void => {
    provider.refresh();
    setRefreshKey((k) => k + 1);
    notify('Reloading data from Microsoft 365 and Power Platform…');
  };

  const failing = provider.getSources().filter((s) => s.state === 'forbidden' || s.state === 'blocked' || s.state === 'error').length;
  const groups = ['Overview', 'Power Platform', 'Microsoft 365', 'Reports'];
  const navGroups: INavLinkGroup[] = groups.map((g) => ({
    name: g === 'Overview' ? undefined : g,
    links: PAGES.filter((p) => p.group === g).map((p) => ({ key: p.id, name: p.title, url: '', icon: p.icon, onClick: () => ctx.navigate(p.id) }))
  }));
  const envOptions: IDropdownOption[] = envs.map((e) => ({ key: e.id, text: `${e.displayName} (${e.type})` }));
  const PageComponent = PAGE_COMPONENTS[page];

  return (
    <AppContext.Provider value={ctx}>
      <ThemeProvider theme={theme} style={{ background: theme.semanticColors.bodyBackground, color: theme.semanticColors.bodyText }}>
        <div ref={rootRef} style={{ display: 'flex', flexDirection: 'column', gap: 12, minWidth: 0 }}>
          <header
            style={{
              display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 8, padding: '8px 12px', borderRadius: 6,
              background: theme.palette.neutralLighterAlt, border: `1px solid ${theme.palette.neutralLight}`
            }}
          >
            <LensLogo size={36} />
            <div style={{ display: 'flex', flexDirection: 'column', minWidth: 0, flex: '1 1 220px' }}>
              <Text variant="large" style={{ fontWeight: 600 }}>{props.title}</Text>
              <Text variant="small" style={{ color: theme.semanticColors.bodySubtext }}>
                {`${provider.tenantLabel}${loadedAt ? ` · data as of ${fmtDateTime(loadedAt)}` : ''}`}
              </Text>
            </div>
            <Dropdown
              placeholder="All environments"
              ariaLabel="Filter by environment"
              multiSelect
              options={envOptions}
              selectedKeys={envFilter}
              onChange={(_, o) => {
                if (!o) return;
                const key = String(o.key);
                setEnvFilter(o.selected ? envFilter.concat([key]) : envFilter.filter((k) => k !== key));
              }}
              styles={{ root: { width: compact ? '100%' : 240 } }}
              disabled={!envs.length}
            />
            {envFilter.length > 0 && <CommandButton text="Clear filter" iconProps={{ iconName: 'ClearFilter' }} onClick={() => setEnvFilter([])} />}
            <CommandButton text="Refresh" iconProps={{ iconName: 'Refresh' }} onClick={refresh} />
            <CommandButton
              text={failing ? `Data sources (${failing} issue${failing === 1 ? '' : 's'})` : 'Data sources'}
              iconProps={{ iconName: failing ? 'Warning' : 'PlugConnected' }}
              onClick={() => setSourcesOpen(true)}
              styles={failing ? { root: { color: theme.semanticColors.errorText }, icon: { color: theme.semanticColors.errorText } } : undefined}
            />
          </header>

          {provider.mode === 'sample' && (
            <MessageBar messageBarType={MessageBarType.info}>
              You are looking at sample data for a fictional tenant. Switch “Data source” to live data in the web part settings to read your tenant.
            </MessageBar>
          )}
          {toast && (
            <MessageBar
              messageBarType={toast.kind === 'error' ? MessageBarType.error : toast.kind === 'success' ? MessageBarType.success : MessageBarType.info}
              onDismiss={() => setToast(undefined)}
              dismissButtonAriaLabel="Close"
            >
              {toast.message}
            </MessageBar>
          )}

          <div style={{ display: 'flex', flexDirection: compact ? 'column' : 'row', gap: 16, alignItems: 'flex-start', minWidth: 0 }}>
            {compact ? (
              <Dropdown
                ariaLabel="Page"
                selectedKey={page}
                options={PAGES.map((p) => ({ key: p.id, text: `${p.group === 'Overview' ? '' : p.group + ' · '}${p.title}` }))}
                onChange={(_, o) => o && ctx.navigate(o.key as PageId)}
                styles={{ root: { width: '100%' } }}
              />
            ) : (
              <nav style={{ width: 210, flex: 'none', position: 'sticky', top: 8 }} aria-label="Dashboard pages">
                <Nav groups={navGroups} selectedKey={page} styles={{ root: { width: 210 }, groupContent: { marginBottom: 12 } }} />
              </nav>
            )}
            <main style={{ flex: 1, minWidth: 0, width: compact ? '100%' : undefined }}>
              <PageErrorBoundary resetKey={page}>
                <PageComponent />
              </PageErrorBoundary>
            </main>
          </div>
        </div>
        <DetailPanel detail={detail} onDismiss={() => setDetail(undefined)} />
        <DataSourcesPanel isOpen={sourcesOpen} onDismiss={() => setSourcesOpen(false)} />
      </ThemeProvider>
    </AppContext.Provider>
  );
};
