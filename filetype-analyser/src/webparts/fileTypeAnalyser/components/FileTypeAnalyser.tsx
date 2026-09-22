import * as React from 'react';
import { useCallback, useMemo, useRef, useState } from 'react';
import {
  DefaultButton,
  Icon,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  ProgressIndicator,
  Stack,
  Text
} from '@fluentui/react';

import { IFileTypeAnalyserProps } from './IFileTypeAnalyserProps';
import styles from './FileTypeAnalyser.module.scss';
import { WebNodeTree } from './WebNodeTree';
import { IScanProgress } from '../models/IScanProgress';
import { ISiteCollectionOverview } from '../models/ISiteCollectionOverview';
import { IWebNode } from '../models/IWebNode';
import { exportOverviewToCsv } from '../services/ExportService';
import { formatBytes } from '../services/formatBytes';
import { ScanCancelledError, SharePointService } from '../services/SharePointService';

const INITIAL_PROGRESS: IScanProgress = {
  phase: 'idle',
  currentItem: '',
  websDiscovered: 0,
  librariesDiscovered: 0,
  librariesScanned: 0
};

const SummaryCard: React.FC<{ icon: string; label: string; value: string }> = ({ icon, label, value }) => (
  <div className={styles.summaryCard}>
    <Icon iconName={icon} className={styles.summaryCardIcon} />
    <div>
      <Text variant="xLarge" block>
        {value}
      </Text>
      <Text variant="small" block className={styles.summaryCardLabel}>
        {label}
      </Text>
    </div>
  </div>
);

export const FileTypeAnalyser: React.FC<IFileTypeAnalyserProps> = (props) => {
  const serviceRef = useRef<SharePointService>();
  if (!serviceRef.current) {
    serviceRef.current = new SharePointService(props.context);
  }

  const [overview, setOverview] = useState<ISiteCollectionOverview | undefined>(undefined);
  const [liveRootWeb, setLiveRootWeb] = useState<IWebNode | undefined>(undefined);
  const [progress, setProgress] = useState<IScanProgress>(INITIAL_PROGRESS);
  const [isScanning, setIsScanning] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | undefined>(undefined);

  const startScan = useCallback(() => {
    const service = serviceRef.current;
    if (!service) {
      return;
    }

    setErrorMessage(undefined);
    setOverview(undefined);
    setLiveRootWeb(undefined);
    setIsScanning(true);
    setProgress(INITIAL_PROGRESS);

    const run = async (): Promise<void> => {
      try {
        const result = await service.scanSiteCollection(
          (p) => setProgress({ ...p }),
          (rootWeb) => setLiveRootWeb({ ...rootWeb })
        );
        setOverview(result);
      } catch (err) {
        if (err instanceof ScanCancelledError) {
          setErrorMessage('Scan cancelled.');
        } else {
          setErrorMessage(err instanceof Error ? err.message : 'The scan failed unexpectedly.');
        }
      } finally {
        setIsScanning(false);
      }
    };

    run().catch(() => {
      /* errors are already handled and surfaced via errorMessage */
    });
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

  const treeRoot = overview ? overview.rootWeb : liveRootWeb;

  const progressLabel = useMemo(() => {
    switch (progress.phase) {
      case 'discovering-structure':
        return `Discovering sites and libraries... (${progress.websDiscovered} webs, ${progress.librariesDiscovered} libraries found)`;
      case 'aggregating-file-types':
        return `Aggregating file types... (${progress.librariesScanned}/${progress.librariesDiscovered} libraries)`;
      case 'starting':
        return 'Starting scan...';
      case 'completed':
        return 'Scan complete.';
      default:
        return '';
    }
  }, [progress]);

  const progressPercent =
    progress.phase === 'aggregating-file-types' && progress.librariesDiscovered > 0
      ? progress.librariesScanned / progress.librariesDiscovered
      : undefined;

  return (
    <div className={styles.fileTypeAnalyser}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" className={styles.header} wrap>
        <Text variant="xLarge" block>
          {props.description || 'File Type Analyser'}
        </Text>
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          {!isScanning && (
            <PrimaryButton text="Start scan" iconProps={{ iconName: 'ScanView' }} onClick={startScan} />
          )}
          {isScanning && <DefaultButton text="Cancel" iconProps={{ iconName: 'Cancel' }} onClick={cancelScan} />}
          <DefaultButton
            text="Export CSV"
            iconProps={{ iconName: 'ExcelDocument' }}
            disabled={!overview}
            onClick={handleExport}
          />
        </Stack>
      </Stack>

      {errorMessage && (
        <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErrorMessage(undefined)}>
          {errorMessage}
        </MessageBar>
      )}

      {isScanning && (
        <div className={styles.progressWrap}>
          <ProgressIndicator label={progressLabel} percentComplete={progressPercent} />
          <Text variant="small" block className={styles.currentItem}>
            {progress.currentItem}
          </Text>
        </div>
      )}

      {overview && (
        <Stack horizontal wrap tokens={{ childrenGap: 12 }} className={styles.summaryRow}>
          <SummaryCard
            icon="Database"
            label="Site storage used"
            value={overview.storage.available ? formatBytes(overview.storage.usedBytes) : 'n/a'}
          />
          <SummaryCard icon="FabricFolder" label="Webs scanned" value={String(overview.totalWebs)} />
          <SummaryCard icon="DocLibrary" label="Libraries scanned" value={String(overview.totalLibraries)} />
          <SummaryCard icon="Page" label="Files indexed" value={overview.totalFiles.toLocaleString()} />
          <SummaryCard
            icon="FileTemplate"
            label="Distinct file types"
            value={String(overview.totalFileTypeStats.length)}
          />
        </Stack>
      )}

      {overview && overview.totalFileTypeStats.length > 0 && (
        <div className={styles.typeBreakdown}>
          <Text variant="large" block className={styles.sectionTitle}>
            File types across the site collection
          </Text>
          <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
            {overview.totalFileTypeStats.map((stat) => (
              <div key={stat.extension} className={styles.typeChip}>
                <span className={styles.typeChipExt}>.{stat.extension}</span>
                <span className={styles.typeChipCount}>{stat.count.toLocaleString()}</span>
              </div>
            ))}
          </Stack>
        </div>
      )}

      {treeRoot && (
        <div className={styles.treeWrap}>
          <Text variant="large" block className={styles.sectionTitle}>
            Site tree
          </Text>
          <WebNodeTree node={treeRoot} depth={0} />
        </div>
      )}

      {!isScanning && !overview && !liveRootWeb && (
        <Stack horizontalAlign="center" className={styles.emptyState}>
          <Icon iconName="FolderSearch" className={styles.emptyIcon} />
          <Text variant="medium">Click &quot;Start scan&quot; to build a file type overview for this site collection.</Text>
        </Stack>
      )}
    </div>
  );
};
