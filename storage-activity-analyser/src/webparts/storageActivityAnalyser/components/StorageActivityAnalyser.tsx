import * as React from 'react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  DefaultButton,
  Dropdown,
  Icon,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  ProgressIndicator,
  Spinner,
  SpinnerSize,
  Stack,
  Text
} from '@fluentui/react';
import { SPPermission } from '@microsoft/sp-page-context';

import { IStorageActivityAnalyserProps } from './IStorageActivityAnalyserProps';
import styles from './StorageActivityAnalyser.module.scss';
import { Dashboard } from './dashboard/Dashboard';
import { formatCompact, formatDate, formatDuration } from './dashboard/format';
import { IScanProgress, IScanResult } from '../models/IScanResult';
import { THRESHOLD_OPTIONS, thresholdLabel } from '../services/activity';
import { exportLibrariesCsv } from '../services/ExportService';
import { formatBytes } from '../services/formatBytes';
import { ResultsStore } from '../services/ResultsStore';
import { ScanCancelledError, StorageScanService } from '../services/StorageScanService';

const INITIAL_PROGRESS: IScanProgress = {
  phase: 'idle',
  currentItem: '',
  websFound: 0,
  librariesFound: 0,
  librariesDone: 0,
  itemsExpected: 0,
  filesRead: 0,
  bytesRead: 0
};

interface INotice {
  type: MessageBarType;
  text: string;
}

export const StorageActivityAnalyser: React.FC<IStorageActivityAnalyserProps> = (props) => {
  const serviceRef = useRef<StorageScanService>();
  if (!serviceRef.current) {
    serviceRef.current = new StorageScanService(props.context);
  }
  const storeRef = useRef<ResultsStore>();
  if (!storeRef.current) {
    storeRef.current = new ResultsStore(props.context);
  }

  // Anyone who can open the page can run a scan (it only reads what they can
  // already see). Only people who manage the site save the result for others.
  const canSave = useMemo(() => {
    const ctx = props.context.pageContext;
    const legacy = ctx.legacyPageContext as { isSiteAdmin?: boolean } | undefined;
    return !!(legacy && legacy.isSiteAdmin) || ctx.web.permissions.hasPermission(SPPermission.manageWeb);
  }, [props.context]);

  const [result, setResult] = useState<IScanResult | undefined>(undefined);
  const [threshold, setThreshold] = useState<number>(props.thresholdMonths);
  const [isLoadingSaved, setIsLoadingSaved] = useState(true);
  const [isScanning, setIsScanning] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [progress, setProgress] = useState<IScanProgress>(INITIAL_PROGRESS);
  const [error, setError] = useState<string | undefined>(undefined);
  const [notice, setNotice] = useState<INotice | undefined>(undefined);

  useEffect(() => {
    setThreshold(props.thresholdMonths);
  }, [props.thresholdMonths]);

  useEffect(() => {
    let active = true;
    setIsLoadingSaved(true);
    setResult(undefined);
    (storeRef.current as ResultsStore)
      .load(props.scope)
      .then((saved) => {
        if (active && saved) {
          setResult(saved);
        }
      })
      .catch((err) => {
        if (active) {
          setNotice({
            type: MessageBarType.warning,
            text: `The last saved scan could not be loaded (${err instanceof Error ? err.message : 'unknown error'}).`
          });
        }
      })
      .then(() => {
        if (active) {
          setIsLoadingSaved(false);
        }
      })
      .catch(() => undefined);
    return () => {
      active = false;
    };
  }, [props.scope]);

  const startScan = useCallback(() => {
    const service = serviceRef.current as StorageScanService;
    const store = storeRef.current as ResultsStore;
    setError(undefined);
    setNotice(undefined);
    setIsScanning(true);
    setProgress(INITIAL_PROGRESS);

    const run = async (): Promise<void> => {
      let scanned: IScanResult;
      try {
        scanned = await service.scan(props.scope, (p) => setProgress(p));
      } catch (err) {
        setError(
          err instanceof ScanCancelledError
            ? 'Scan cancelled. Any previous results are still shown.'
            : err instanceof Error
            ? err.message
            : 'The scan failed unexpectedly.'
        );
        setIsScanning(false);
        return;
      }
      setResult(scanned);
      setIsScanning(false);

      if (!canSave) {
        setNotice({
          type: MessageBarType.info,
          text: 'These results are shown to you only. When a site owner runs a scan, it is saved for everyone who opens this page.'
        });
        return;
      }
      setIsSaving(true);
      try {
        await store.save(scanned);
        setNotice({
          type: MessageBarType.success,
          text: 'Scan saved. Everyone who opens this page will see these results until the next scan.'
        });
      } catch (err) {
        setNotice({
          type: MessageBarType.warning,
          text:
            'The scan finished, but the results could not be saved for other people ' +
            `(${err instanceof Error ? err.message : 'unknown error'}). They are shown here until you leave the page.`
        });
      } finally {
        setIsSaving(false);
      }
    };
    run().catch(() => undefined);
  }, [props.scope, canSave]);

  const cancelScan = useCallback(() => {
    (serviceRef.current as StorageScanService).cancel();
  }, []);

  const thresholdOptions: IDropdownOption[] = useMemo(
    () => THRESHOLD_OPTIONS.map((m) => ({ key: m, text: thresholdLabel(m) })),
    []
  );

  const progressLabel = useMemo(() => {
    switch (progress.phase) {
      case 'starting':
        return 'Starting scan…';
      case 'discovering':
        return `Finding sites and libraries… ${progress.websFound} sites, ${progress.librariesFound} libraries`;
      case 'reading-files':
        return (
          `Reading file sizes and dates… library ${Math.min(progress.librariesDone + 1, progress.librariesFound)} of ` +
          `${progress.librariesFound} · ${formatCompact(progress.filesRead)} files · ${formatBytes(progress.bytesRead)}`
        );
      default:
        return '';
    }
  }, [progress]);

  const percent =
    progress.phase === 'reading-files' && progress.itemsExpected > 0
      ? Math.min(1, progress.filesRead / progress.itemsExpected)
      : undefined;

  const scopeText = props.scope === 'currentWeb' ? 'this site and its subsites' : 'this site collection';
  const lastScanLine = result
    ? `Last scanned ${formatDate(new Date(result.scanCompletedAt))}` +
      (result.scannedBy ? ` by ${result.scannedBy}` : '') +
      ` · took ${formatDuration(new Date(result.scanCompletedAt).getTime() - new Date(result.scanStartedAt).getTime())}`
    : undefined;

  return (
    <div className={styles.analyser}>
      <div className={styles.hero}>
        <div className={styles.heroText}>
          <div className={styles.heroEyebrow}>
            <Icon iconName="History" /> Storage activity · {scopeText}
          </div>
          <h2 className={styles.heroTitle}>{props.title || 'Storage Activity Analyser'}</h2>
          <div className={styles.heroMeta}>
            {lastScanLine || (isLoadingSaved ? 'Loading the last scan…' : 'No scan yet')}
            {isSaving && ' · saving…'}
          </div>
        </div>
        <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 8 }} className={styles.heroActions}>
          <Dropdown
            label="Inactive after"
            ariaLabel="Treat files as inactive when they have not been modified for"
            options={thresholdOptions}
            selectedKey={threshold}
            onChange={(_, option) => option && setThreshold(Number(option.key))}
            className={styles.thresholdPicker}
          />
          {!isScanning && (
            <PrimaryButton
              text={result ? 'Run new scan' : 'Start scan'}
              iconProps={{ iconName: result ? 'Refresh' : 'Search' }}
              onClick={startScan}
              disabled={isLoadingSaved || isSaving}
            />
          )}
          {isScanning && <DefaultButton text="Cancel" iconProps={{ iconName: 'Cancel' }} onClick={cancelScan} />}
          <DefaultButton
            text="Export CSV"
            iconProps={{ iconName: 'ExcelDocument' }}
            disabled={!result || isScanning}
            onClick={() => result && exportLibrariesCsv(result, threshold)}
          />
        </Stack>
      </div>

      {error && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(undefined)}>
          {error}
        </MessageBar>
      )}
      {notice && (
        <MessageBar messageBarType={notice.type} onDismiss={() => setNotice(undefined)}>
          {notice.text}
        </MessageBar>
      )}

      {isScanning && (
        <div className={styles.progressCard}>
          <ProgressIndicator label={progressLabel} percentComplete={percent} />
          <Text variant="small" block className={styles.currentItem}>
            {progress.currentItem}
          </Text>
          <Text variant="small" block className={styles.progressHint}>
            The scan reads each file&apos;s size and last-modified date, 5,000 files per request, using your own
            access. Large site collections can take several minutes; you can keep working in another tab.
          </Text>
        </div>
      )}

      {isLoadingSaved && !isScanning && (
        <div className={styles.loading}>
          <Spinner size={SpinnerSize.medium} label="Loading the last scan…" />
        </div>
      )}

      {!isScanning && result && <Dashboard result={result} thresholdMonths={threshold} />}

      {!isLoadingSaved && !isScanning && !result && (
        <Stack horizontalAlign="center" className={styles.emptyState}>
          <span className={styles.emptyBadge}>
            <Icon iconName="History" />
          </span>
          <Text variant="large" block className={styles.emptyTitle}>
            How much of your storage is still in use?
          </Text>
          <Text variant="medium" className={styles.emptyText}>
            Click “Start scan” to see how much of {scopeText}&apos;s file storage is in files that are still being worked
            on, and how much is in files nobody has changed for {thresholdLabel(threshold)} or more.
          </Text>
        </Stack>
      )}
    </div>
  );
};
