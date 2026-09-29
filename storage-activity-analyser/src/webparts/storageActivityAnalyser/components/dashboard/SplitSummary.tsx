import * as React from 'react';
import { useMemo } from 'react';
import { Icon } from '@fluentui/react';

import { IScanResult } from '../../models/IScanResult';
import { IActivitySplit, newestFileAge, thresholdLabel } from '../../services/activity';
import { formatBytes } from '../../services/formatBytes';
import styles from './Dashboard.module.scss';
import { formatCompact, formatPercent } from './format';
import { useCountUp, useEntrance } from './motion';

interface ISplitSummaryProps {
  result: IScanResult;
  split: IActivitySplit;
  thresholdMonths: number;
}

const Figure: React.FC<{ kind: 'inactive' | 'active'; bytes: number; files: number; total: number; caption: string }> = ({
  kind,
  bytes,
  files,
  total,
  caption
}) => {
  const animated = useCountUp(bytes);
  return (
    <div className={styles.figure}>
      <div className={styles.figureHead}>
        <span className={`${styles.swatch} ${styles[kind]}`} />
        <span className={styles.figureLabel}>{kind === 'inactive' ? 'Inactive files' : 'Active files'}</span>
      </div>
      <div className={styles.figureValue} title={`${bytes.toLocaleString()} bytes`}>
        {formatBytes(animated)}
      </div>
      <div className={styles.figureMeta}>
        <strong>{formatPercent(bytes, total)}</strong> of file storage · {formatCompact(files)} files
      </div>
      <div className={styles.figureCaption}>{caption}</div>
    </div>
  );
};

export const SplitSummary: React.FC<ISplitSummaryProps> = ({ result, split, thresholdMonths }) => {
  const entering = useEntrance(result.scanCompletedAt);
  const period = thresholdLabel(thresholdMonths);

  const facts = useMemo(() => {
    let versionBytes = 0;
    let librariesWithMetrics = 0;
    let dormantLibraries = 0;
    let dormantBytes = 0;
    let failed = 0;
    for (const lib of result.libraries) {
      if (lib.metrics) {
        librariesWithMetrics++;
        versionBytes += Math.max(0, lib.metrics.totalSize - lib.metrics.fileStreamSize);
      }
      const newest = newestFileAge(lib.histogram);
      if (newest !== undefined && newest >= thresholdMonths) {
        dormantLibraries++;
        dormantBytes += lib.bytes;
      }
      if (lib.error) {
        failed++;
      }
    }
    failed += result.webs.filter((w) => w.error).length;
    return { versionBytes, librariesWithMetrics, dormantLibraries, dormantBytes, failed };
  }, [result, thresholdMonths]);

  const inactiveShare = split.totalBytes ? split.inactiveBytes / split.totalBytes : 0;

  return (
    <section className={entering ? `${styles.card} ${styles.enter}` : styles.card} aria-label="Active and inactive storage">
      <div className={styles.figures}>
        <Figure
          kind="inactive"
          bytes={split.inactiveBytes}
          files={split.inactiveFiles}
          total={split.totalBytes}
          caption={`Not modified for ${period} or more`}
        />
        <Figure
          kind="active"
          bytes={split.activeBytes}
          files={split.activeFiles}
          total={split.totalBytes}
          caption={`Modified in the last ${period}`}
        />
        <div className={styles.figure}>
          <div className={styles.figureHead}>
            <span className={styles.figureIcon}>
              <Icon iconName="Database" />
            </span>
            <span className={styles.figureLabel}>All files</span>
          </div>
          <div className={styles.figureValue}>{formatBytes(split.totalBytes)}</div>
          <div className={styles.figureMeta}>
            {formatCompact(split.totalFiles)} files in {result.libraries.length} libraries
          </div>
          <div className={styles.figureCaption}>
            {result.siteStorageBytes !== undefined
              ? `Site collection storage used: ${formatBytes(result.siteStorageBytes)}`
              : 'Current version of every file'}
          </div>
        </div>
      </div>

      <div
        className={styles.splitBar}
        role="img"
        aria-label={`${formatPercent(split.inactiveBytes, split.totalBytes)} of file storage is inactive, ${formatPercent(
          split.activeBytes,
          split.totalBytes
        )} is active`}
      >
        {split.inactiveBytes > 0 && (
          <span
            className={`${styles.splitSegment} ${styles.inactive}`}
            style={{ flexGrow: split.inactiveBytes }}
            title={`Inactive: ${formatBytes(split.inactiveBytes)} (${formatPercent(split.inactiveBytes, split.totalBytes)})`}
          />
        )}
        {split.activeBytes > 0 && (
          <span
            className={`${styles.splitSegment} ${styles.active}`}
            style={{ flexGrow: split.activeBytes }}
            title={`Active: ${formatBytes(split.activeBytes)} (${formatPercent(split.activeBytes, split.totalBytes)})`}
          />
        )}
      </div>

      <ul className={styles.insights}>
        {split.totalBytes > 0 && (
          <li>
            <Icon iconName="Lightbulb" className={styles.insightIcon} />
            {inactiveShare >= 0.5
              ? `Most of the file storage here (${formatPercent(split.inactiveBytes, split.totalBytes)}) has not been touched for ${period} or more. Archiving or cleaning it up would free ${formatBytes(split.inactiveBytes)}.`
              : `${formatPercent(split.inactiveBytes, split.totalBytes)} of the file storage (${formatBytes(split.inactiveBytes)}) has not been touched for ${period} or more.`}
          </li>
        )}
        {facts.dormantLibraries > 0 && (
          <li>
            <Icon iconName="DocLibrary" className={styles.insightIcon} />
            {facts.dormantLibraries === 1 ? '1 library has' : `${facts.dormantLibraries} libraries have`} had no file
            changed for {period} or more ({formatBytes(facts.dormantBytes)}). They are marked “Dormant” below.
          </li>
        )}
        {facts.librariesWithMetrics > 0 && facts.versionBytes > 0 && (
          <li>
            <Icon iconName="History" className={styles.insightIcon} />
            Version history takes a further {formatBytes(facts.versionBytes)} in the libraries SharePoint reported it for.
            It is not included in the figures above; trimming version limits can reduce it.
          </li>
        )}
        {facts.failed > 0 && (
          <li>
            <Icon iconName="Warning" className={styles.insightIcon} />
            {facts.failed === 1 ? '1 site or library' : `${facts.failed} sites or libraries`} could not be read fully
            with your access, so the totals may be lower than the real figure. See the table below or the CSV.
          </li>
        )}
      </ul>
    </section>
  );
};
