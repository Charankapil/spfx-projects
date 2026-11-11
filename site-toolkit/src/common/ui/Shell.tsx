import * as React from 'react';
import { useEffect, useState } from 'react';
import { Icon, MessageBar, MessageBarType } from '@fluentui/react';

import { getQueueSnapshot, IQueueSnapshot, subscribeToQueue } from '../RequestQueue';
import styles from './Toolkit.module.scss';

export interface IShellProps {
  icon: string;
  title: string;
  subtitle?: string;
  actions?: React.ReactNode;
  error?: string;
  onDismissError?: () => void;
  notice?: string;
  onDismissNotice?: () => void;
}

/**
 * Shows "SharePoint asked us to slow down" while the page-wide queue is
 * paused, so a throttle reads as a short wait instead of a broken web part.
 */
const ThrottleNotice: React.FC = () => {
  const [snapshot, setSnapshot] = useState<IQueueSnapshot>(getQueueSnapshot());
  const [now, setNow] = useState<number>(Date.now());

  useEffect(() => subscribeToQueue(setSnapshot), []);

  useEffect(() => {
    if (!snapshot.pausedUntil) {
      return undefined;
    }
    const timer = window.setInterval(() => setNow(Date.now()), 1000);
    return () => window.clearInterval(timer);
  }, [snapshot.pausedUntil]);

  const remaining = Math.ceil((snapshot.pausedUntil - now) / 1000);
  if (!snapshot.pausedUntil || remaining <= 0) {
    return null;
  }
  return (
    <MessageBar messageBarType={MessageBarType.warning} isMultiline={false}>
      SharePoint asked this page to slow down. Carrying on automatically in {remaining} second{remaining === 1 ? '' : 's'}.
    </MessageBar>
  );
};

export const Shell: React.FC<IShellProps> = (props) => (
  <section className={styles.shell}>
    <div className={styles.header}>
      <div className={styles.headerText}>
        <span className={styles.headerIcon} aria-hidden="true">
          <Icon iconName={props.icon} />
        </span>
        <div>
          <h2 className={styles.title}>{props.title}</h2>
          {props.subtitle && <p className={styles.subtitle}>{props.subtitle}</p>}
        </div>
      </div>
      {props.actions && <div className={styles.actions}>{props.actions}</div>}
    </div>
    <div className={styles.body}>
      <ThrottleNotice />
      {props.error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={props.onDismissError} dismissButtonAriaLabel="Close">
          {props.error}
        </MessageBar>
      )}
      {props.notice && (
        <MessageBar messageBarType={MessageBarType.success} onDismiss={props.onDismissNotice} dismissButtonAriaLabel="Close">
          {props.notice}
        </MessageBar>
      )}
      {props.children}
    </div>
  </section>
);

export const Empty: React.FC<{ icon: string; text: string }> = (props) => (
  <div className={styles.empty}>
    <Icon iconName={props.icon} className={styles.emptyIcon} aria-hidden="true" />
    <div>{props.text}</div>
    {props.children && <div style={{ marginTop: 12 }}>{props.children}</div>}
  </div>
);
