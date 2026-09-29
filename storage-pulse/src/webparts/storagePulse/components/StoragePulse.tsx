import * as React from 'react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  DefaultButton,
  Dropdown,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Stack
} from '@fluentui/react';
import { SPPermission } from '@microsoft/sp-page-context';
import * as strings from 'StoragePulseWebPartStrings';

import { IStoragePulseProps } from './IStoragePulseProps';
import styles from './StoragePulse.module.scss';
import { Dashboard } from './dashboard/Dashboard';
import { formatDate, formatDuration } from './dashboard/format';
import { ErrorBoundary } from './ErrorBoundary';
import { PulseMark } from './PulseMark';
import { ScanPanel } from './ScanPanel';
import { format, thresholdLabel } from './text';
import { IScanProgress, IScanResult } from '../models/IScanResult';
import { THRESHOLD_OPTIONS } from '../services/activity';
import { exportLibrariesCsv } from '../services/ExportService';
import { ResultsStore } from '../services/ResultsStore';
import { ScanCancelledError, StorageScanService } from '../services/StorageScanService';

const INITIAL_PROGRESS: IScanProgress = {
  phase: 'idle',
  currentItem: '',
  websFound: 0,
  librariesFound: 0,
  librariesDone: 0,
  itemsExpected: 0,
  itemsRead: 0,
  filesRead: 0,
  bytesRead: 0
};

const DAY_MS = 24 * 60 * 60 * 1000;

interface INotice {
  type: MessageBarType;
  text: string;
}

/** Sites and libraries that could not be read completely (what "Retry failed" reads again). */
function countFailures(result: IScanResult): number {
  return result.libraries.filter((l) => l.error).length + result.webs.filter((w) => w.error).length;
}

function errorText(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

export const StoragePulse: React.FC<IStoragePulseProps> = (props) => {
  const serviceRef = useRef<StorageScanService>();
  if (!serviceRef.current) {
    serviceRef.current = new StorageScanService(props.context);
  }
  const storeRef = useRef<ResultsStore>();
  if (!storeRef.current) {
    storeRef.current = new ResultsStore(props.context);
  }

  // People who manage the site save scans for everyone. Whether anyone else
  // may scan (for themselves only) is the owner's choice in the settings.
  const isOwner = useMemo(() => {
    const ctx = props.context.pageContext;
    const legacy = ctx.legacyPageContext as { isSiteAdmin?: boolean } | undefined;
    return !!(legacy && legacy.isSiteAdmin) || ctx.web.permissions.hasPermission(SPPermission.manageWeb);
  }, [props.context]);
  const canScan = isOwner || props.scanPermission === 'everyone';

  const [result, setResult] = useState<IScanResult | undefined>(undefined);
  const [savedBy, setSavedBy] = useState<string | undefined>(undefined);
  const [threshold, setThreshold] = useState<number>(props.thresholdMonths);
  const [isLoadingSaved, setIsLoadingSaved] = useState(true);
  const [isScanning, setIsScanning] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [progress, setProgress] = useState<IScanProgress>(INITIAL_PROGRESS);
  const [error, setError] = useState<string | undefined>(undefined);
  const [notice, setNotice] = useState<INotice | undefined>(undefined);
  const mounted = useRef(true);

  useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
      // Leaving the page (or removing the web part) stops a running scan.
      (serviceRef.current as StorageScanService).cancel();
    };
  }, []);

  useEffect(() => {
    setThreshold(props.thresholdMonths);
  }, [props.thresholdMonths]);

  useEffect(() => {
    let active = true;
    setIsLoadingSaved(true);
    setResult(undefined);
    setSavedBy(undefined);
    (storeRef.current as ResultsStore)
      .load(props.scope)
      .then((saved) => {
        if (active && saved) {
          setResult(saved.result);
          setSavedBy(saved.savedBy);
        }
      })
      .catch((err) => {
        if (active) {
          setNotice({ type: MessageBarType.warning, text: format(strings.LoadFailed, { error: errorText(err) }) });
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

  const resultRef = useRef<IScanResult | undefined>(undefined);
  resultRef.current = result;

  const { scope, includeHidden, excludeSystemLibraries, excludedLibraries, scanMode } = props;
  /** Runs a full scan, or (retry) re-reads only what failed in the current result, then saves for owners. */
  const runScan = useCallback(
    (retry: boolean) => {
      const service = serviceRef.current as StorageScanService;
      const store = storeRef.current as ResultsStore;
      const previous = resultRef.current;
      setError(undefined);
      setNotice(undefined);
      setIsScanning(true);
      setProgress(INITIAL_PROGRESS);

      const options = {
        scope,
        includeHidden,
        excludeSystemLibraries,
        excludedLibraries,
        // Quick scans measure libraries unchanged for the chosen "Inactive after" period as a whole.
        quickAfterMonths: scanMode === 'quick' ? threshold : undefined
      };
      const onProgress = (p: IScanProgress): void => {
        if (mounted.current) {
          setProgress(p);
        }
      };

      const run = async (): Promise<void> => {
        let scanned: IScanResult;
        try {
          scanned =
            retry && previous
              ? await service.retryFailed(previous, options, onProgress)
              : await service.scan(options, onProgress);
        } catch (err) {
          if (mounted.current) {
            setError(err instanceof ScanCancelledError ? strings.ScanCancelled : errorText(err) || strings.ScanFailed);
            setIsScanning(false);
          }
          return;
        }
        if (!mounted.current) {
          return;
        }
        setResult(scanned);
        setSavedBy(undefined);
        const retryNotice = retry
          ? countFailures(scanned) === 0
            ? strings.RetryAllRead
            : format(strings.RetrySomeLeft, { count: countFailures(scanned) })
          : undefined;

        if (!isOwner) {
          setIsScanning(false);
          setNotice({ type: MessageBarType.info, text: retryNotice || strings.ScanNotSavedForOthers });
          return;
        }
        setIsSaving(true);
        try {
          await store.save(scanned);
          if (mounted.current) {
            setNotice({ type: MessageBarType.success, text: retryNotice || strings.ScanSaved });
          }
        } catch (err) {
          if (mounted.current) {
            setNotice({ type: MessageBarType.warning, text: format(strings.SaveFailed, { error: errorText(err) }) });
          }
        } finally {
          if (mounted.current) {
            setIsSaving(false);
            setIsScanning(false);
          }
        }
      };
      run().catch(() => undefined);
    },
    [scope, includeHidden, excludeSystemLibraries, excludedLibraries, scanMode, threshold, isOwner]
  );
  const startScan = useCallback(() => runScan(false), [runScan]);
  const retryFailed = useCallback(() => runScan(true), [runScan]);

  const cancelScan = useCallback(() => {
    (serviceRef.current as StorageScanService).cancel();
  }, []);

  // After a quick scan, periods longer than the one it used are approximate for libraries measured as a whole.
  const approxAbove =
    result && result.quickAfterMonths !== undefined && result.libraries.some((l) => l.measuredAsWhole)
      ? result.quickAfterMonths
      : undefined;
  const thresholdOptions: IDropdownOption[] = useMemo(
    () =>
      THRESHOLD_OPTIONS.map((m) => ({
        key: m,
        text: thresholdLabel(m) + (approxAbove !== undefined && m > approxAbove ? strings.ApproxSuffix : '')
      })),
    [approxAbove]
  );
  const failures = result ? countFailures(result) : 0;

  const scopeShort = props.scope === 'currentWeb' ? strings.ScopeCurrentWebShort : strings.ScopeSiteCollectionShort;
  const scopePhrase = props.scope === 'currentWeb' ? strings.ScopeCurrentWebPhrase : strings.ScopeSiteCollectionPhrase;

  let lastScanLine: string | undefined;
  let ageDays = 0;
  if (result) {
    const completed = new Date(result.scanCompletedAt);
    const who = savedBy || result.scannedBy;
    const values = {
      date: formatDate(completed),
      name: who,
      duration: formatDuration(completed.getTime() - new Date(result.scanStartedAt).getTime())
    };
    lastScanLine =
      format(who ? strings.LastScanned : strings.LastScannedNoName, values) +
      ` \u00b7 ${result.quickAfterMonths !== undefined ? strings.QuickScanTag : strings.DetailedScanTag}`;
    ageDays = Math.floor((Date.now() - completed.getTime()) / DAY_MS);
  }
  const isStale = !!result && !isScanning && props.staleAfterDays > 0 && ageDays >= props.staleAfterDays;

  const rootClass = `${styles.storagePulse} ${props.theme.isDark ? styles.dark : ''}`;
  const rootStyle = props.theme.accent ? ({ '--sp-accent': props.theme.accent } as React.CSSProperties) : undefined;

  return (
    <div className={rootClass} style={rootStyle}>
      <header className={styles.hero}>
        <span className={styles.aurora} aria-hidden="true" />
        <div className={styles.heroBrand}>
          <PulseMark size={52} busy={isScanning} />
          <div className={styles.heroText}>
            <div className={styles.heroEyebrow}>
              {strings.DefaultTitle} · {scopeShort}
            </div>
            <h2 className={styles.heroTitle}>{props.title || strings.DefaultTitle}</h2>
            <div className={styles.heroMeta}>
              {lastScanLine || (isLoadingSaved ? strings.LoadingLastScan : strings.NoScanYet)}
              {isSaving && ` · ${strings.Saving}`}
            </div>
          </div>
        </div>
        <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 8 }} className={styles.heroActions}>
          <Dropdown
            label={strings.InactiveAfterLabel}
            ariaLabel={strings.InactiveAfterAria}
            options={thresholdOptions}
            selectedKey={threshold}
            onChange={(_, option) => option && setThreshold(Number(option.key))}
            className={styles.thresholdPicker}
          />
          {canScan && !isScanning && (
            <PrimaryButton
              text={result ? strings.RunNewScan : strings.StartScan}
              iconProps={{ iconName: result ? 'Refresh' : 'Search' }}
              onClick={startScan}
              disabled={isLoadingSaved}
              className={styles.scanButton}
            />
          )}
          {canScan && !isScanning && failures > 0 && (
            <DefaultButton
              text={format(strings.RetryFailed, { count: failures })}
              iconProps={{ iconName: 'Sync' }}
              onClick={retryFailed}
              disabled={isLoadingSaved}
            />
          )}
          {isScanning && !isSaving && (
            <DefaultButton text={strings.Cancel} iconProps={{ iconName: 'Cancel' }} onClick={cancelScan} />
          )}
          <DefaultButton
            text={strings.ExportCsv}
            iconProps={{ iconName: 'ExcelDocument' }}
            disabled={!result || isScanning}
            onClick={() => result && exportLibrariesCsv(result, threshold)}
          />
        </Stack>
      </header>

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
      {isStale && (
        <MessageBar messageBarType={MessageBarType.warning}>
          {format(canScan ? strings.StaleResultsCanScan : strings.StaleResults, { days: ageDays })}
        </MessageBar>
      )}

      {isScanning && <ScanPanel progress={progress} isSaving={isSaving} />}

      {isLoadingSaved && !isScanning && (
        <div className={styles.skeleton} aria-hidden="true">
          <span />
          <span />
          <span />
        </div>
      )}

      {!isScanning && result && (
        <ErrorBoundary resetKey={result}>
          <Dashboard result={result} thresholdMonths={threshold} />
        </ErrorBoundary>
      )}

      {!isLoadingSaved && !isScanning && !result && (
        <div className={styles.emptyState}>
          <PulseMark size={72} />
          <h3 className={styles.emptyTitle}>{strings.EmptyTitle}</h3>
          <p className={styles.emptyText}>
            {canScan
              ? format(strings.EmptyText, { scope: scopePhrase, period: thresholdLabel(threshold) })
              : strings.EmptyNoPermission}
          </p>
          {canScan && (
            <PrimaryButton
              text={strings.StartScan}
              iconProps={{ iconName: 'Search' }}
              onClick={startScan}
              className={styles.scanButton}
            />
          )}
        </div>
      )}

      {result && !isScanning && (
        <span className={styles.srOnly} aria-live="polite">
          {strings.LiveDone}
        </span>
      )}
    </div>
  );
};
