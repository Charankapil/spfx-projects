import * as React from 'react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  DefaultButton,
  getTheme,
  IButtonStyles,
  Icon,
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

import { IFileTypeAnalyserProps } from './IFileTypeAnalyserProps';
import styles from './FileTypeAnalyser.module.scss';
import { WebNodeTree } from './WebNodeTree';
import { Dashboard } from './dashboard/Dashboard';
import { formatDate, formatDuration } from './dashboard/format';
import { IScanProgress } from '../models/IScanProgress';
import { ISiteCollectionOverview } from '../models/ISiteCollectionOverview';
import { IWebNode } from '../models/IWebNode';
import { exportOverviewToCsv } from '../services/ExportService';
import { ResultsStore, summarize } from '../services/ResultsStore';
import { ScanCancelledError, SharePointService } from '../services/SharePointService';

const INITIAL_PROGRESS: IScanProgress = {
  phase: 'idle',
  currentItem: '',
  websDiscovered: 0,
  librariesDiscovered: 0,
  librariesScanned: 0
};

/**
 * Buttons sit on the theme-coloured hero band, so they are styled from the
 * Fluent theme - which SPFx loads with the site's own theme colours.
 */
function heroButtonStyles(): { primary: IButtonStyles; secondary: IButtonStyles } {
  const { palette } = getTheme();
  const white = '#ffffff';
  return {
    primary: {
      root: { background: white, color: palette.themeDarker, border: 'none', borderRadius: 6, fontWeight: 600 },
      rootHovered: { background: palette.themeLighterAlt, color: palette.themeDarker },
      rootPressed: { background: palette.themeLighter, color: palette.themeDarker },
      rootDisabled: { background: 'rgba(255,255,255,0.55)', color: palette.themeDarker },
      icon: { color: palette.themeDarker }
    },
    secondary: {
      root: { background: 'rgba(255,255,255,0.08)', color: white, border: '1px solid rgba(255,255,255,0.55)', borderRadius: 6 },
      rootHovered: { background: 'rgba(255,255,255,0.18)', color: white },
      rootPressed: { background: 'rgba(255,255,255,0.26)', color: white },
      rootDisabled: { background: 'transparent', color: 'rgba(255,255,255,0.5)', borderColor: 'rgba(255,255,255,0.3)' },
      icon: { color: white },
      iconDisabled: { color: 'rgba(255,255,255,0.5)' }
    }
  };
}

interface INotice {
  type: MessageBarType;
  text: string;
}

export const FileTypeAnalyser: React.FC<IFileTypeAnalyserProps> = (props) => {
  const serviceRef = useRef<SharePointService>();
  if (!serviceRef.current) {
    serviceRef.current = new SharePointService(props.context);
  }
  const storeRef = useRef<ResultsStore>();
  if (!storeRef.current) {
    storeRef.current = new ResultsStore(props.context);
  }

  // Scanning writes the shared results file, so only people who manage the
  // site get the button; everyone else sees the last saved scan.
  const canScan = useMemo(() => {
    const ctx = props.context.pageContext;
    const legacy = ctx.legacyPageContext as { isSiteAdmin?: boolean } | undefined;
    return !!(legacy && legacy.isSiteAdmin) || ctx.web.permissions.hasPermission(SPPermission.manageWeb);
  }, [props.context]);

  const [overview, setOverview] = useState<ISiteCollectionOverview | undefined>(undefined);
  const overviewRef = useRef<ISiteCollectionOverview | undefined>(undefined);
  overviewRef.current = overview;

  const [isLoadingSaved, setIsLoadingSaved] = useState(true);
  const [liveRootWeb, setLiveRootWeb] = useState<IWebNode | undefined>(undefined);
  const [progress, setProgress] = useState<IScanProgress>(INITIAL_PROGRESS);
  const [isScanning, setIsScanning] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | undefined>(undefined);
  const [notice, setNotice] = useState<INotice | undefined>(undefined);
  const heroButtons = useMemo(heroButtonStyles, []);

  useEffect(() => {
    let active = true;
    const load = async (): Promise<void> => {
      try {
        const saved = await (storeRef.current as ResultsStore).load();
        if (active && saved) {
          setOverview(saved);
        }
      } catch (err) {
        if (active) {
          setNotice({
            type: MessageBarType.warning,
            text: `The last saved scan could not be loaded (${err instanceof Error ? err.message : 'unknown error'}).`
          });
        }
      } finally {
        if (active) {
          setIsLoadingSaved(false);
        }
      }
    };
    load().catch(() => undefined);
    return () => {
      active = false;
    };
  }, []);

  const startScan = useCallback(() => {
    const service = serviceRef.current;
    const store = storeRef.current;
    if (!service || !store) {
      return;
    }

    setErrorMessage(undefined);
    setNotice(undefined);
    setLiveRootWeb(undefined);
    setIsScanning(true);
    setProgress(INITIAL_PROGRESS);

    const run = async (): Promise<void> => {
      let result: ISiteCollectionOverview;
      try {
        result = await service.scanSiteCollection(
          (p) => setProgress({ ...p }),
          (rootWeb) => setLiveRootWeb({ ...rootWeb })
        );
      } catch (err) {
        // The previously shown results stay on screen; only the new scan is discarded.
        setErrorMessage(
          err instanceof ScanCancelledError
            ? 'Scan cancelled. The previous results are still shown.'
            : err instanceof Error
            ? err.message
            : 'The scan failed unexpectedly.'
        );
        setIsScanning(false);
        return;
      }

      const baseline = overviewRef.current;
      result.previous = baseline ? summarize(baseline) : undefined;
      setOverview(result);
      setLiveRootWeb(undefined);
      setIsScanning(false);

      setIsSaving(true);
      try {
        await store.save(result);
        setNotice({
          type: MessageBarType.success,
          text: 'Scan saved. Everyone who opens this page will see these results until the next scan.'
        });
      } catch (err) {
        setNotice({
          type: MessageBarType.warning,
          text:
            'The scan finished, but its results could not be saved for other users ' +
            `(${err instanceof Error ? err.message : 'unknown error'}). They are shown here until you leave the page.`
        });
      } finally {
        setIsSaving(false);
      }
    };

    run().catch(() => undefined);
  }, []);

  const cancelScan = useCallback(() => {
    if (serviceRef.current) {
      serviceRef.current.cancel();
    }
  }, []);

  const handleExport = useCallback(() => {
    if (overview) {
      exportOverviewToCsv(overview);
    }
  }, [overview]);

  const progressLabel = useMemo(() => {
    switch (progress.phase) {
      case 'discovering-structure':
        return `Discovering sites and libraries… (${progress.websDiscovered} sites, ${progress.librariesDiscovered} libraries found)`;
      case 'aggregating-file-types':
        return `Counting file types… (${progress.librariesScanned}/${progress.librariesDiscovered} libraries)`;      case 'starting':
        return 'Starting scan…';
      default:
        return '';
    }
  }, [progress]);

  const progressPercent =
    progress.phase === 'aggregating-file-types' && progress.librariesDiscovered > 0
      ? progress.librariesScanned / progress.librariesDiscovered
      : undefined;

  const lastScanLine = overview && overview.scanCompletedAt
    ? `Last scanned ${formatDate(overview.scanCompletedAt)}` +
      (overview.scannedBy ? ` by ${overview.scannedBy}` : '') +
      ` · took ${formatDuration(overview.scanCompletedAt.getTime() - overview.scanStartedAt.getTime())}`
    : undefined;

  const treeRoot = isScanning ? liveRootWeb : overview ? overview.rootWeb : liveRootWeb;

  let activeStep = 0;
  if (isSaving) {
    activeStep = 2;
  } else if (progress.phase === 'aggregating-file-types') {
    activeStep = 1;
  }
  const scanSteps: { label: string; state: 'done' | 'active' | 'pending' }[] = [
    'Find sites & libraries',
    'Count file types',
    'Save for everyone'
  ].map((label, i) => ({ label, state: i < activeStep ? 'done' : i === activeStep ? 'active' : 'pending' }));

  return (
    <div className={styles.fileTypeAnalyser}>
      <div className={styles.hero}>
        <div className={styles.heroText}>
          <div className={styles.heroEyebrow}>
            <Icon iconName="FabricFolderSearch" /> File inventory
          </div>
          <h2 className={styles.heroTitle}>{props.description || 'File Type Analyser'}</h2>
          <div className={styles.heroMeta}>
            {lastScanLine || (isLoadingSaved ? 'Loading the last scan\u2026' : 'No scan yet')}
            {isSaving && ' \u00b7 saving\u2026'}
          </div>
        </div>
        <Stack horizontal wrap tokens={{ childrenGap: 8 }} className={styles.heroActions}>
          {canScan && !isScanning && (
            <PrimaryButton
              text={overview ? 'Run new scan' : 'Start scan'}
              iconProps={{ iconName: 'ScanView' }}
              onClick={startScan}
              disabled={isLoadingSaved || isSaving}
              styles={heroButtons.primary}
            />
          )}
          {isScanning && (
            <DefaultButton text="Cancel" iconProps={{ iconName: 'Cancel' }} onClick={cancelScan} styles={heroButtons.secondary} />
          )}
          <DefaultButton
            text="Export CSV"
            iconProps={{ iconName: 'ExcelDocument' }}
            disabled={!overview || isScanning}
            onClick={handleExport}
            styles={heroButtons.secondary}
          />
        </Stack>
      </div>

      {errorMessage && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErrorMessage(undefined)}>
          {errorMessage}
        </MessageBar>
      )}
      {notice && (
        <MessageBar messageBarType={notice.type} onDismiss={() => setNotice(undefined)}>
          {notice.text}
        </MessageBar>
      )}

      {(isScanning || isSaving) && (
        <div className={styles.progressCard}>
          <ol className={styles.steps}>
            {scanSteps.map((step, i) => (
              <li key={step.label} className={`${styles.step} ${styles[step.state]}`}>
                <span className={styles.stepDot}>{step.state === 'done' ? <Icon iconName="CheckMark" /> : i + 1}</span>
                <span className={styles.stepLabel}>{step.label}</span>
              </li>
            ))}
          </ol>
          {isScanning && (
            <>
              <ProgressIndicator label={progressLabel} percentComplete={progressPercent} />
              <Text variant="small" block className={styles.currentItem}>
                {progress.currentItem}
              </Text>
            </>
          )}
        </div>
      )}

      {isLoadingSaved && (
        <div className={styles.loading}>
          <Spinner size={SpinnerSize.medium} label="Loading the last scan…" />
        </div>
      )}

      {!isScanning && overview && <Dashboard overview={overview} />}

      {treeRoot && (
        <div className={styles.treeWrap}>
          <Text variant="large" block className={styles.sectionTitle}>
            Site tree
          </Text>
          <WebNodeTree node={treeRoot} depth={0} />
        </div>
      )}

      {!isLoadingSaved && !isScanning && !overview && !liveRootWeb && (
        <Stack horizontalAlign="center" className={styles.emptyState}>
          <span className={styles.emptyBadge}>
            <Icon iconName="FabricFolderSearch" />
          </span>
          <Text variant="large" block className={styles.emptyTitle}>
            Nothing scanned yet
          </Text>
          <Text variant="medium" className={styles.emptyText}>
            {canScan
              ? 'No scan has been run yet. Click “Start scan” to build a file type overview for this site collection.'
              : 'No scan results yet. A site owner needs to run the first scan.'}
          </Text>
        </Stack>
      )}
    </div>
  );
};
