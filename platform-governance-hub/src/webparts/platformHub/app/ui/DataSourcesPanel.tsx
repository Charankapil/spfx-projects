import * as React from 'react';
import { DefaultButton, Panel, PanelType, PrimaryButton, Text } from '@fluentui/react';
import { useApp } from './AppContext';
import { Pill, Tone } from './components/Layout';
import { SourceState } from '../models';
import { fmtDateTime } from '../logic/format';

const LABELS: Record<SourceState, { text: string; tone: Tone }> = {
  ok: { text: 'Working', tone: 'good' },
  pending: { text: 'Not checked yet', tone: 'neutral' },
  forbidden: { text: 'Access denied', tone: 'critical' },
  blocked: { text: 'Blocked (network/CORS)', tone: 'critical' },
  error: { text: 'Error', tone: 'warning' },
  skipped: { text: 'Skipped', tone: 'neutral' }
};

/** Shows, per API, whether it works for the signed-in admin and which permission it needs. */
export const DataSourcesPanel: React.FC<{ isOpen: boolean; onDismiss: () => void }> = ({ isOpen, onDismiss }) => {
  const { provider, theme } = useApp();
  const [checking, setChecking] = React.useState(false);
  const sources = provider.getSources();
  const checkAll = async (): Promise<void> => {
    setChecking(true);
    const swallow = (p: Promise<unknown>): Promise<unknown> => p.then(() => undefined, () => undefined);
    await Promise.all([
      swallow(provider.getInventory().then(() => provider.getMakers())),
      swallow(provider.getDlpPolicies()),
      swallow(provider.getSkus()),
      swallow(provider.getServiceHealth()),
      swallow(provider.getReportsConcealed()),
      swallow(provider.getSites()),
      swallow(provider.getTeams()),
      swallow(provider.getAdoption()),
      swallow(provider.getIdentity())
    ]);
    setChecking(false);
  };
  return (
    <Panel isOpen={isOpen} onDismiss={onDismiss} type={PanelType.medium} headerText="Data sources" closeButtonAriaLabel="Close" isLightDismiss>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 14, paddingBottom: 24 }}>
        <Text variant="small" style={{ color: theme.semanticColors.bodySubtext }}>
          Every API this dashboard calls, with the permission it needs. “Access denied” means a permission request is not approved in
          SharePoint admin center &gt; Advanced &gt; API access, or your account lacks the admin role shown. “Blocked” means the browser could not reach the API.
        </Text>
        <div style={{ display: 'flex', gap: 8 }}>
          <PrimaryButton text={checking ? 'Checking…' : 'Check all sources'} disabled={checking} onClick={() => { checkAll().catch(() => undefined); }} />
          <DefaultButton text="Refresh data" onClick={() => provider.refresh()} />
        </div>
        {sources.map((s) => {
          const label = LABELS[s.state];
          return (
            <div key={s.id} style={{ borderTop: `1px solid ${theme.palette.neutralLighter}`, paddingTop: 10, display: 'flex', flexDirection: 'column', gap: 2 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', gap: 8, alignItems: 'center' }}>
                <Text style={{ fontWeight: 600 }}>{s.label}</Text>
                <Pill tone={label.tone} text={label.text} />
              </div>
              <Text variant="small" style={{ color: theme.semanticColors.bodySubtext }}>{s.endpoint}</Text>
              <Text variant="small">{s.permission}</Text>
              {s.message && <Text variant="small" style={{ color: label.tone === 'critical' ? theme.semanticColors.errorText : theme.semanticColors.bodySubtext, wordBreak: 'break-word' }}>{s.message}</Text>}
              {s.checkedAt && <Text variant="small" style={{ color: theme.semanticColors.bodySubtext }}>{`Checked ${fmtDateTime(s.checkedAt)}`}</Text>}
            </div>
          );
        })}
      </div>
    </Panel>
  );
};
