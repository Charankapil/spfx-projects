import * as React from 'react';
import { DefaultButton, MessageBar, MessageBarType } from '@fluentui/react';
import { GraphAccessError } from '../../core/GraphClient';
import { hasScope } from '../../services/GraphApi';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from './context';
import { useLoader } from './hooks';
import { Card, Loading } from './ui';

/**
 * Which Graph scopes are approved, read from the access token (no Graph request).
 * `undefined` while checking.
 */
export function useGraphScopes(): { scopes: string[] | undefined; error?: string; reload: () => void } {
  const ctx = useAdmin();
  const loader = useLoader(() => ctx.graph.grantedScopes(), []);
  return { scopes: loader.data, error: loader.error, reload: loader.reload };
}

/** Shown in place of a Microsoft 365 feature until its permission is approved. */
export const NeedsApproval: React.FC<{ scopes: string[]; feature: string }> = ({ scopes, feature }) => {
  const ctx = useAdmin();
  return (
    <Card>
      <MessageBar messageBarType={MessageBarType.warning} isMultiline>
        <strong>{feature} needs Microsoft Graph API access that is not approved yet.</strong> Required permission: {scopes.join(', ')} (delegated, Microsoft Graph).
        It was requested when the solution was added to the app catalog; a <strong>Global Administrator</strong> approves it once in SharePoint admin center &gt; Advanced &gt; API access.
        Everything else in this tool keeps working without it.
      </MessageBar>
      <div className={styles.actions} style={{ marginTop: 10 }}>
        <DefaultButton iconProps={{ iconName: 'PlugConnected' }} onClick={() => ctx.openView('apiaccess')}>
          How to approve
        </DefaultButton>
      </div>
    </Card>
  );
};

/** Renders children only when every scope is approved; otherwise explains what is missing. */
export const GraphGate: React.FC<{ scopes: string[]; feature: string; children: React.ReactNode }> = ({ scopes, feature, children }) => {
  const access = useGraphScopes();
  if (access.scopes === undefined && !access.error) {
    return <Loading text="Checking Microsoft 365 API access…" />;
  }
  // If the token could not be read we still let the feature try: its own errors explain what is wrong.
  if (access.scopes && !scopes.every((s) => hasScope(access.scopes as string[], s))) {
    return <NeedsApproval scopes={scopes.filter((s) => !hasScope(access.scopes as string[], s))} feature={feature} />;
  }
  return <>{children}</>;
};

export function graphErrorText(e: unknown): string {
  return e instanceof GraphAccessError || e instanceof Error ? e.message : String(e);
}
