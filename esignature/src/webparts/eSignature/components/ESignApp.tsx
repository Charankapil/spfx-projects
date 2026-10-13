import * as React from 'react';
import { DefaultButton, Icon, MessageBar, MessageBarType, PrimaryButton, Spinner } from '@fluentui/react';

import { ITemplate } from '../models/types';
import { ESignService } from '../services/ESignService';
import { Dashboard } from './Dashboard';
import { RequestView } from './RequestView';
import { SendWizard } from './SendWizard';
import { Setup } from './Setup';
import { Templates } from './Templates';
import { Verify } from './Verify';
import { errorText } from './shared';
import styles from './ESign.module.scss';

export interface IESignAppProps {
  title: string;
  service: ESignService;
  canManage: boolean;
}

type View =
  | { kind: 'dashboard' }
  | { kind: 'send'; sourceUrl?: string; template?: ITemplate }
  | { kind: 'request'; id: string }
  | { kind: 'templates' }
  | { kind: 'verify' }
  | { kind: 'setup' };

/** Deep links: ?esignRequest=<id> opens a request, ?esignSource=<server-relative url> starts sending that file. */
function initialView(): View {
  const params = new URLSearchParams(window.location.search);
  const request = params.get('esignRequest');
  if (request) {
    return { kind: 'request', id: request };
  }
  const source = params.get('esignSource');
  if (source) {
    return { kind: 'send', sourceUrl: source };
  }
  return { kind: 'dashboard' };
}

function setQuery(view: View): void {
  try {
    const url = new URL(window.location.href);
    url.searchParams.delete('esignRequest');
    url.searchParams.delete('esignSource');
    if (view.kind === 'request') {
      url.searchParams.set('esignRequest', view.id);
    }
    window.history.replaceState(window.history.state, '', url.toString());
  } catch {
    // Some hosts (Teams) do not allow rewriting the URL; deep links just are not updated there.
  }
}

export function ESignApp(props: IESignAppProps): JSX.Element {
  const { service } = props;
  const [provisioned, setProvisioned] = React.useState<boolean>();
  const [error, setError] = React.useState<string>();
  const [view, setViewState] = React.useState<View>(initialView);
  const [notice, setNotice] = React.useState<{ text: string; warnings: string[] }>();
  const [epoch, setEpoch] = React.useState(0);

  const setView = (next: View): void => {
    setViewState(next);
    setQuery(next);
    setEpoch(e => e + 1);
  };

  React.useEffect(() => {
    service.provisioning
      .isProvisioned()
      .then(setProvisioned)
      .catch(e => setError(errorText(e)));
  }, [service]);

  const nav = (kind: View['kind'], label: string, icon: string): JSX.Element =>
    view.kind === kind ? (
      <PrimaryButton key={kind} text={label} iconProps={{ iconName: icon }} onClick={() => setView({ kind } as View)} />
    ) : (
      <DefaultButton key={kind} text={label} iconProps={{ iconName: icon }} onClick={() => setView({ kind } as View)} />
    );

  let body: JSX.Element;
  if (error) {
    body = <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>;
  } else if (provisioned === undefined) {
    body = <Spinner label="Loading..." styles={{ root: { padding: 32 } }} />;
  } else if (!provisioned || view.kind === 'setup') {
    body = (
      <Setup
        service={service}
        canManage={props.canManage}
        provisioned={!!provisioned}
        onDone={() => {
          setProvisioned(true);
          setView({ kind: 'dashboard' });
        }}
      />
    );
  } else {
    switch (view.kind) {
      case 'send':
        body = (
          <SendWizard
            key={epoch}
            service={service}
            sourceUrl={view.sourceUrl}
            template={view.template}
            onCancel={() => setView({ kind: 'dashboard' })}
            onSent={(id, warnings) => {
              setNotice({ text: 'Your request has been sent.', warnings });
              setView({ kind: 'request', id });
            }}
          />
        );
        break;
      case 'request':
        body = <RequestView key={view.id} service={service} requestId={view.id} onBack={() => setView({ kind: 'dashboard' })} />;
        break;
      case 'templates':
        body = <Templates service={service} onUse={template => setView({ kind: 'send', template })} />;
        break;
      case 'verify':
        body = <Verify service={service} onOpen={id => setView({ kind: 'request', id })} />;
        break;
      default:
        body = <Dashboard key={epoch} service={service} onOpen={id => setView({ kind: 'request', id })} />;
    }
  }

  return (
    <div className={styles.app}>
      <div className={styles.header}>
        <h2 className={styles.title}>
          <Icon iconName="InsertSignatureLine" className={styles.titleIcon} />
          {props.title}
        </h2>
        {provisioned && (
          <div className={styles.nav}>
            {nav('dashboard', 'Dashboard', 'ViewDashboard')}
            {nav('send', 'New request', 'Add')}
            {nav('templates', 'Templates', 'FileTemplate')}
            {nav('verify', 'Verify', 'Shield')}
            {props.canManage && nav('setup', 'Setup', 'Settings')}
          </div>
        )}
      </div>
      {notice && (
        <MessageBar messageBarType={MessageBarType.success} onDismiss={() => setNotice(undefined)}>
          {notice.text}
          {notice.warnings.map((w, i) => (
            <div key={i}>{w}</div>
          ))}
        </MessageBar>
      )}
      {body}
    </div>
  );
}
