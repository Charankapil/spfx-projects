import * as React from 'react';
import { PrimaryButton, MessageBar, MessageBarType, Icon, Spinner, SpinnerSize } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { SpClient } from '../../services/SpClient';
import { MIGRATIONS, runMigrations, ISetupStatus } from '../../services/Provisioner';
import { Logo } from '../common/Logo';

type StepState = 'pending' | 'running' | 'done' | 'error';

/**
 * First-run setup and later upgrades. Runs the numbered migrations as the signed-in site owner.
 * Every step checks before it creates, so running it again after a failure is safe.
 */
export function Setup(props: { sp: SpClient; siteTitle: string; status: ISetupStatus; onDone: () => void }): JSX.Element {
  const { status } = props;
  const upgrading = status.installedVersion > 0;
  const [states, setStates] = React.useState<StepState[]>(MIGRATIONS.map(m => (m.version <= status.installedVersion ? 'done' : 'pending')));
  const [running, setRunning] = React.useState(false);
  const [error, setError] = React.useState('');

  if (!status.canSetup) {
    return (
      <div className={styles.setup}>
        <h1 className={styles.pageTitle} style={{ display: 'flex', alignItems: 'center', gap: 10 }}><Logo size={32} /> Work Boards</h1>
        <MessageBar messageBarType={MessageBarType.warning}>
          {upgrading
            ? 'Work Boards on this site needs a quick update before it can be used. Ask a site owner to open this page to run it.'
            : `Work Boards isn't set up on ${props.siteTitle} yet. Ask a site owner to open this page and set it up.`}
        </MessageBar>
      </div>
    );
  }

  const run = async (): Promise<void> => {
    setRunning(true);
    setError('');
    try {
      await runMigrations(props.sp, status.installedVersion, (i, s, message) => {
        setStates(prev => prev.map((p, j) => (j === i ? s : p)));
        if (message) {
          setError(message);
        }
      });
    } catch (e) {
      setError((e as Error).message + ' You can run setup again; finished steps are skipped.');
      setRunning(false);
      return;
    }
    props.onDone();
  };

  return (
    <div className={styles.setup}>
      <h1 className={styles.pageTitle} style={{ display: 'flex', alignItems: 'center', gap: 10 }}><Logo size={32} /> {upgrading ? 'Update Work Boards' : 'Set up Work Boards'}</h1>
      <p style={{ margin: 0 }}>
        {upgrading
          ? `This version of Work Boards needs a few changes on ${props.siteTitle}. Your boards and items are not touched.`
          : `This creates hidden lists on ${props.siteTitle} to hold boards and settings. It takes under a minute. Everyone with access to this site can then use Work Boards, and each board you create gets its own list.`}
      </p>
      <p className={styles.muted} style={{ margin: 0 }}>
        Nothing leaves this site: no app registration, no Microsoft Graph permissions and no external services.
      </p>
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
      <div className={styles.steps}>
        {MIGRATIONS.map((m, i) => (
          <div key={m.version} className={styles.step}>
            <span className={`${styles.stepIcon} ${states[i] === 'done' ? styles.stepDone : states[i] === 'error' ? styles.stepError : ''}`}>
              {states[i] === 'done' ? <Icon iconName="CheckMark" /> : states[i] === 'error' ? '!' : states[i] === 'running' ? <Spinner size={SpinnerSize.xSmall} /> : m.version}
            </span>
            <span>{m.title}</span>
          </div>
        ))}
      </div>
      <div>
        <PrimaryButton text={running ? 'Working…' : upgrading ? 'Update now' : 'Set up'} disabled={running} onClick={() => { run().catch(() => undefined); }} />
      </div>
    </div>
  );
}
