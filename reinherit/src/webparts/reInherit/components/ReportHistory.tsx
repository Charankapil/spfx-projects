import * as React from 'react';
import { useEffect, useState } from 'react';
import {
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  Icon,
  MessageBar,
  MessageBarType,
  SelectionMode,
  Spinner,
  SpinnerSize
} from '@fluentui/react';

import { IReportSummary, IRunReport } from '../models/IRunReport';
import { ReportStore } from '../services/ReportStore';
import { formatDate } from './format';
import { ReportView } from './ReportView';
import styles from './ReInherit.module.scss';

interface IReportHistoryProps {
  store: ReportStore;
  /** Bumped by the parent after it saves a report, so the list reloads. */
  refreshToken: number;
}

export const ReportHistory: React.FC<IReportHistoryProps> = ({ store, refreshToken }) => {
  const [summaries, setSummaries] = useState<IReportSummary[] | undefined>(undefined);
  const [error, setError] = useState<string | undefined>(undefined);
  const [open, setOpen] = useState<IRunReport | undefined>(undefined);
  const [opening, setOpening] = useState(false);

  useEffect(() => {
    let active = true;
    setSummaries(undefined);
    store
      .listReports()
      .then((list) => active && setSummaries(list))
      .catch((err) => active && setError(err instanceof Error ? err.message : 'Reports could not be loaded.'));
    return () => {
      active = false;
    };
  }, [store, refreshToken]);

  const openReport = (summary: IReportSummary): void => {
    setOpening(true);
    setError(undefined);
    store
      .loadReport(summary)
      .then((report) => {
        setOpening(false);
        if (report) {
          setOpen(report);
        } else {
          setError('That report file is missing or could not be read.');
        }
      })
      .catch((err) => {
        setOpening(false);
        setError(err instanceof Error ? err.message : 'The report could not be loaded.');
      });
  };

  if (open) {
    return (
      <div>
        <DefaultButton
          text="All reports"
          iconProps={{ iconName: 'Back' }}
          onClick={() => setOpen(undefined)}
          className={styles.backButton}
        />
        <ReportView report={open} />
      </div>
    );
  }

  const columns: IColumn[] = [
    {
      key: 'when',
      name: 'Run',
      minWidth: 150,
      maxWidth: 190,
      onRender: (s: IReportSummary) => (
        <button type="button" className={styles.linkButton} onClick={() => openReport(s)}>
          {formatDate(new Date(s.startedAt))}
        </button>
      )
    },
    {
      key: 'mode',
      name: 'Type',
      minWidth: 110,
      maxWidth: 140,
      onRender: (s: IReportSummary) => (
        <span className={`${styles.statusPill} ${s.mode === 'restore' ? styles.status_restored : styles.status_found}`}>
          {s.mode === 'restore' ? (s.cancelled ? 'Restore (cancelled)' : 'Restore') : 'Scan only'}
        </span>
      )
    },
    {
      key: 'scope',
      name: 'Scope',
      minWidth: 180,
      isResizable: true,
      onRender: (s: IReportSummary) => (
        <span className={styles.pathCell} title={s.scopes.map((x) => x.url).join('\n')}>
          {s.scopes.map((x) => x.title).join(', ')}
        </span>
      )
    },
    {
      key: 'found',
      name: 'Unique',
      minWidth: 70,
      maxWidth: 90,
      onRender: (s: IReportSummary) => <span className={styles.numCell}>{s.stats.uniqueFound.toLocaleString()}</span>
    },
    {
      key: 'restored',
      name: 'Restored',
      minWidth: 70,
      maxWidth: 90,
      onRender: (s: IReportSummary) => <span className={styles.numCell}>{s.mode === 'restore' ? s.stats.restored.toLocaleString() : ''}</span>
    },
    {
      key: 'failed',
      name: 'Failed',
      minWidth: 60,
      maxWidth: 80,
      onRender: (s: IReportSummary) => (
        <span className={s.stats.failed ? `${styles.numCell} ${styles.errorText}` : styles.numCell}>
          {s.mode === 'restore' ? s.stats.failed.toLocaleString() : ''}
        </span>
      )
    },
    { key: 'by', name: 'Run by', fieldName: 'runBy', minWidth: 120, isResizable: true }
  ];

  return (
    <div className={styles.tableCard}>
      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(undefined)}>
          {error}
        </MessageBar>
      )}
      {opening && <Spinner size={SpinnerSize.small} label="Opening report…" />}
      {!summaries && !error && <Spinner size={SpinnerSize.medium} label="Loading reports…" />}
      {summaries && summaries.length === 0 && (
        <div className={styles.emptyTable}>
          <Icon iconName="ReportDocument" className={styles.emptyTableIcon} />
          No reports yet. Every scan and restore run is saved here.
        </div>
      )}
      {summaries && summaries.length > 0 && (
        <DetailsList
          items={summaries}
          columns={columns}
          compact
          selectionMode={SelectionMode.none}
          layoutMode={DetailsListLayoutMode.justified}
          onItemInvoked={(s: IReportSummary) => openReport(s)}
        />
      )}
    </div>
  );
};
