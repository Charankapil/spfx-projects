import * as React from 'react';
import { ITheme } from '@fluentui/react';
import { IDataProvider } from '../providers/IDataProvider';

export type PageId =
  | 'overview' | 'environments' | 'apps' | 'flows' | 'agents' | 'makers' | 'connectors' | 'capacity' | 'governance'
  | 'pbi' | 'pbiContent' | 'pbiActivity'
  | 'sites' | 'teams' | 'adoption' | 'identity' | 'export';

export interface IDetail {
  title: string;
  subtitle?: string;
  fields: { label: string; value: React.ReactNode }[];
  links?: { label: string; href: string }[];
  /** Extra content loaded on demand (sharing, run history...). */
  extra?: React.ReactNode;
}

export interface IAppContext {
  provider: IDataProvider;
  theme: ITheme;
  isDark: boolean;
  /** Incremented whenever the provider reports new data or status. */
  version: number;
  /** Incremented when the user refreshes; data hooks reload on change. */
  refreshKey: number;
  envFilter: string[];
  setEnvFilter: (ids: string[]) => void;
  navigate: (page: PageId) => void;
  openDetail: (detail: IDetail) => void;
  notify: (message: string, kind?: 'info' | 'success' | 'error') => void;
  compact: boolean;
}

export const AppContext = React.createContext<IAppContext | undefined>(undefined);

export function useApp(): IAppContext {
  const ctx = React.useContext(AppContext);
  if (!ctx) throw new Error('AppContext is missing');
  return ctx;
}

export interface IDataState<T> {
  data: T | undefined;
  error: unknown;
  loading: boolean;
  reload: () => void;
}

/**
 * Loads data from the provider. Results are cached by the provider, so
 * revisiting a page is instant; `reload` re-runs after a refresh.
 */
export function useData<T>(load: (p: IDataProvider) => Promise<T>, deps: React.DependencyList = []): IDataState<T> {
  const { provider, refreshKey } = useApp();
  const [state, setState] = React.useState<{ data: T | undefined; error: unknown; loading: boolean }>({ data: undefined, error: undefined, loading: true });
  const [nonce, setNonce] = React.useState(0);
  React.useEffect(() => {
    let alive = true;
    setState((s) => ({ data: s.data, error: undefined, loading: true }));
    load(provider).then(
      (data) => alive && setState({ data, error: undefined, loading: false }),
      (error) => alive && setState({ data: undefined, error, loading: false })
    );
    return () => {
      alive = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [provider, nonce, refreshKey, ...deps]);
  return { ...state, reload: () => setNonce((n) => n + 1) };
}

/** Filters any environment-scoped list by the global environment filter. */
export function byEnv<T extends { environmentId: string }>(items: T[], envFilter: string[]): T[] {
  if (!envFilter.length) return items;
  return items.filter((i) => envFilter.indexOf(i.environmentId) >= 0);
}
