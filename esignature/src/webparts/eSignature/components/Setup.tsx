import * as React from 'react';
import { MessageBar, MessageBarType, PrimaryButton } from '@fluentui/react';

import { ESignService } from '../services/ESignService';
import { errorText } from './shared';
import styles from './ESign.module.scss';

export interface ISetupProps {
  service: ESignService;
  canManage: boolean;
  provisioned: boolean;
  onDone: () => void;
}

export function Setup(props: ISetupProps): JSX.Element {
  const [log, setLog] = React.useState<string[]>([]);
  const [running, setRunning] = React.useState(false);
  const [error, setError] = React.useState<string>();
  const [done, setDone] = React.useState(false);

  const run = async (): Promise<void> => {
    setRunning(true);
    setError(undefined);
    setLog([]);
    try {
      await props.service.provisioning.provision(message => setLog(l => l.concat(message)));
      setDone(true);
    } catch (e) {
      setError(errorText(e));
    } finally {
      setRunning(false);
    }
  };

  if (!props.canManage) {
    return (
      <div className={styles.card}>
        <h3 className={styles.cardTitle}>eSignature is not set up on this site yet</h3>
        <p>A site owner needs to open this page once and select Set up eSignature.</p>
      </div>
    );
  }

  return (
    <div className={styles.card}>
      <h3 className={styles.cardTitle}>{props.provisioned ? 'Repair eSignature storage' : 'Set up eSignature on this site'}</h3>
      <p>This creates, on this site:</p>
      <ul>
        <li>
          <strong>eSignature Requests</strong> and <strong>eSignature Activity</strong> lists, where people can only edit the items they
          created - so nobody can change someone else&apos;s request or record a signature on their behalf;
        </li>
        <li>
          <strong>eSignature Templates</strong>, for reusable recipient roles and field layouts;
        </li>
        <li>
          <strong>eSignature Documents</strong>, a library holding each document sent for signature and its signed copy.
        </li>
      </ul>
      <p className={styles.muted}>
        Everyone who sends or signs needs at least Contribute permission on this site (or on these four lists). Running setup again is safe:
        anything that already exists is kept.
      </p>
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      {log.length > 0 && <div className={styles.log}>{log.join('\n')}</div>}
      <div className={styles.footer}>
        {done ? (
          <PrimaryButton text="Start using eSignature" onClick={props.onDone} />
        ) : (
          <PrimaryButton text={running ? 'Setting up...' : 'Set up eSignature'} disabled={running} onClick={run} />
        )}
      </div>
    </div>
  );
}
