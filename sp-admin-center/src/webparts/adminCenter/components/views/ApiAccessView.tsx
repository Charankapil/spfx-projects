import * as React from 'react';
import { DefaultButton, MessageBar, MessageBarType } from '@fluentui/react';
import { GRAPH_PERMISSIONS, hasScope } from '../../services/GraphApi';
import styles from '../AdminCenter.module.scss';
import { useGraphScopes } from '../shared/graph';
import { Card, Loading, Pill, ViewHeader } from '../shared/ui';

/**
 * What the optional Microsoft 365 features need, what is approved right now (read from the
 * access token, no Graph request) and how to approve it.
 */
export const ApiAccessView: React.FC = () => {
  const access = useGraphScopes();
  const scopes = access.scopes || [];
  const approved = GRAPH_PERMISSIONS.filter((p) => hasScope(scopes, p.scope)).length;

  return (
    <div className={styles.view}>
      <ViewHeader title="API access" hint="The Microsoft 365 features use Microsoft Graph. Everything else works without it.">
        <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={access.reload}>
          Check again
        </DefaultButton>
      </ViewHeader>

      {access.scopes === undefined && !access.error ? (
        <Loading text="Checking which permissions are approved…" />
      ) : (
        <>
          {access.error ? (
            <MessageBar messageBarType={MessageBarType.warning} isMultiline>
              Could not read the approved permissions: {access.error}
            </MessageBar>
          ) : (
            <MessageBar messageBarType={approved === GRAPH_PERMISSIONS.length ? MessageBarType.success : approved ? MessageBarType.warning : MessageBarType.info} isMultiline>
              {approved === GRAPH_PERMISSIONS.length
                ? 'All Microsoft 365 features are available.'
                : `${approved} of ${GRAPH_PERMISSIONS.length} permissions approved. Features that need a missing permission explain it and stay off; nothing else is affected.`}
            </MessageBar>
          )}
          <Card title="Permissions this solution requests (Microsoft Graph, delegated)">
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
              <thead>
                <tr style={{ textAlign: 'left' }}>
                  <th style={{ padding: '6px 8px' }}>Permission</th>
                  <th style={{ padding: '6px 8px' }}>Status</th>
                  <th style={{ padding: '6px 8px' }}>Unlocks</th>
                  <th style={{ padding: '6px 8px' }}>Your role must also allow it</th>
                </tr>
              </thead>
              <tbody>
                {GRAPH_PERMISSIONS.map((p) => (
                  <tr key={p.scope} style={{ borderTop: '1px solid #e5e7eb', verticalAlign: 'top' }}>
                    <td style={{ padding: '8px' }}>
                      <strong>{p.scope}</strong>
                      <div className={styles.muted}>{p.why}</div>
                    </td>
                    <td style={{ padding: '8px' }}>{access.error ? <Pill>Unknown</Pill> : hasScope(scopes, p.scope) ? <Pill kind="good">Approved</Pill> : <Pill kind="warning">Not approved</Pill>}</td>
                    <td style={{ padding: '8px' }}>
                      <ul style={{ margin: 0, paddingLeft: 16 }}>
                        {p.features.map((f) => (
                          <li key={f}>{f}</li>
                        ))}
                      </ul>
                    </td>
                    <td style={{ padding: '8px' }}>{p.role}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </Card>
        </>
      )}

      <div className={styles.grid}>
        <Card title="How to approve (one time, per tenant)">
          <ol style={{ margin: 0, paddingLeft: 18, fontSize: 13, lineHeight: 1.7 }}>
            <li>Upload the .sppkg to the tenant app catalog and click Deploy. The four requests are created automatically.</li>
            <li>
              A <strong>Global Administrator</strong> opens SharePoint admin center &gt; Advanced &gt; <strong>API access</strong>, selects each pending request from &ldquo;sharepoint-admin-center-client-side-solution&rdquo; and clicks <strong>Approve</strong>. Approving only some is fine.
            </li>
            <li>Come back here and click Check again (sign-out/in may be needed for a fresh token).</li>
          </ol>
          <p className={styles.muted}>
            Or with PowerShell (SharePoint Online Management Shell, as a Global Administrator): <code>Get-SPOTenantServicePrincipalPermissionRequests</code> then <code>Approve-SPOTenantServicePrincipalPermissionRequest -RequestId &lt;id&gt;</code>.
          </p>
          <p className={styles.muted} style={{ marginBottom: 0 }}>
            Still no app registration or client ID: the permissions are granted to Microsoft&rsquo;s built-in &ldquo;SharePoint Online Client Extensibility&rdquo; app, and every call runs as the signed-in admin, limited by that person&rsquo;s own role.
          </p>
        </Card>
        <Card title="For the live tenant dashboard">
          <ul style={{ margin: 0, paddingLeft: 18, fontSize: 13, lineHeight: 1.7 }}>
            <li>
              Site addresses are hidden in usage reports by default. A Global Administrator or Reports administrator turns that off in Microsoft 365 admin center &gt; Settings &gt; Org settings &gt; Reports: clear &ldquo;Display concealed user, group, and site names in all reports&rdquo;.
            </li>
            <li>Usage reports are refreshed by Microsoft once a day and are usually 1&ndash;2 days behind.</li>
            <li>The report has no archive status, so archived sites count as active; the CSV import from your flow still covers that.</li>
          </ul>
        </Card>
      </div>
    </div>
  );
};
