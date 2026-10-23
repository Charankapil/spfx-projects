import * as React from 'react';
import { CancelledError } from '../../core/SPClient';
import { useAdmin } from './context';

export interface ILoader<T> {
  data: T | undefined;
  loading: boolean;
  error: string | undefined;
  /** Re-run, bypassing the read cache. */
  reload: () => void;
}

/**
 * Runs `fn` when the view mounts, when the target site changes, or on Refresh.
 * Nothing polls; navigating back to a view within the cache window reuses the
 * cached responses instead of calling SharePoint again.
 */
export function useLoader<T>(fn: () => Promise<T>, deps: unknown[], enabled = true): ILoader<T> {
  const ctx = useAdmin();
  const [state, setState] = React.useState<{ data?: T; loading: boolean; error?: string }>({ loading: enabled });
  const [tick, setTick] = React.useState(0);

  React.useEffect(() => {
    if (!enabled) {
      setState({ loading: false });
      return undefined;
    }
    let alive = true;
    setState((s) => ({ data: s.data, loading: true }));
    fn().then(
      (data) => {
        if (alive) {
          setState({ data, loading: false });
        }
      },
      (e: Error) => {
        if (alive && !(e instanceof CancelledError)) {
          setState({ loading: false, error: e && e.message ? e.message : String(e) });
        }
      }
    );
    return () => {
      alive = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [...deps, tick, ctx.reloadToken, enabled]);

  const reload = React.useCallback(() => {
    ctx.client.clearCache();
    setTick((t) => t + 1);
  }, [ctx.client]);

  return { data: state.data, loading: state.loading, error: state.error, reload };
}
