import * as React from 'react';
import { useMemo } from 'react';
import { Icon } from '@fluentui/react';
import * as strings from 'StoragePulseWebPartStrings';

import { IScanResult } from '../../models/IScanResult';
import { safeHref } from '../../services/safeData';
import { format } from '../text';
import styles from './Dashboard.module.scss';

interface IIssue {
  key: string;
  kind: string;
  title: string;
  where: string;
  url?: string;
  reason: string;
  warning: boolean;
}

/**
 * Every site and library the scan could not read completely, with SharePoint's
 * reason, so an owner can tell throttling, timeouts and permissions apart
 * (and see what "Retry failed" will read again).
 */
export const ScanIssues: React.FC<{ result: IScanResult }> = ({ result }) => {
  const origin = useMemo(() => new URL(result.rootUrl).origin, [result]);
  const issues = useMemo(() => {
    const list: IIssue[] = [];
    for (const web of result.webs) {
      if (web.error) {
        list.push({
          key: `w|${web.url}`,
          kind: strings.IssueKindSite,
          title: web.title,
          where: web.url,
          url: web.url,
          reason: web.error,
          warning: false
        });
      }
    }
    for (const lib of result.libraries) {
      if (lib.error || lib.unreadItems) {
        list.push({
          key: `l|${lib.webUrl}|${lib.id}`,
          kind: strings.IssueKindLibrary,
          title: lib.title,
          where: lib.webTitle,
          url: lib.url,
          reason: lib.error
            ? lib.partial
              ? format(strings.PartlyRead, { error: lib.error })
              : lib.error
            : format(strings.IssueHidden, {
                hidden: (lib.unreadItems || 0).toLocaleString(),
                total: lib.itemCount.toLocaleString()
              }),
          warning: !lib.error
        });
      }
    }
    return list;
  }, [result]);

  if (issues.length === 0) {
    return null;
  }
  return (
    <section className={styles.card} aria-label={strings.IssuesTitle}>
      <div className={styles.cardHeader}>
        <div>
          <h3 className={styles.cardTitle}>{strings.IssuesTitle}</h3>
          <div className={styles.cardSubtitle}>{strings.IssuesSubtitle}</div>
        </div>
      </div>
      <ul className={styles.issues}>
        {issues.map((issue) => {
          const href = safeHref(issue.url, origin);
          return (
            <li key={issue.key} className={styles.issue}>
              <Icon iconName={issue.warning ? 'Info' : 'Warning'} className={issue.warning ? styles.issueInfo : styles.issueError} />
              <div className={styles.issueBody}>
                <div>
                  <span className={styles.issueKind}>{issue.kind}</span>{' '}
                  {href ? (
                    <a href={href} target="_blank" rel="noopener noreferrer" className={styles.nameLink}>
                      {issue.title}
                    </a>
                  ) : (
                    <span className={styles.nameText}>{issue.title}</span>
                  )}
                  <span className={styles.subtle}> · {issue.where}</span>
                </div>
                <div className={styles.issueReason}>{issue.reason}</div>
              </div>
            </li>
          );
        })}
      </ul>
    </section>
  );
};
