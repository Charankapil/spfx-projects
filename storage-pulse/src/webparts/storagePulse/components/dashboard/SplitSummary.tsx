import * as React from 'react';
import { useMemo } from 'react';
import { Icon } from '@fluentui/react';
import * as strings from 'StoragePulseWebPartStrings';

import { IScanResult } from '../../models/IScanResult';
import { IActivitySplit, newestFileAge } from '../../services/activity';
import { formatBytes } from '../../services/formatBytes';
import { format, thresholdLabel } from '../text';
import styles from './Dashboard.module.scss';
import { formatCompact, formatPercent } from './format';
import { useCountUp, useEntrance } from './motion';

interface ISplitSummaryProps {
  result: IScanResult;
  split: IActivitySplit;
  thresholdMonths: number;
}

const RADIUS = 70;
const CIRCUMFERENCE = 2 * Math.PI * RADIUS;
/** Surface gap between the two arcs, in the same units as the circumference (arcs use butt caps). */
const GAP = 4;

/** Donut gauge: inactive share (orange) then active share (blue), with the inactive % in the middle. */
const Gauge: React.FC<{ split: IActivitySplit }> = ({ split }) => {
  const share = split.totalBytes ? split.inactiveBytes / split.totalBytes : 0;
  const percent = useCountUp(Math.round(share * 100));
  const both = split.inactiveBytes > 0 && split.activeBytes > 0;
  const inactiveLength = Math.max(0, share * CIRCUMFERENCE - (both ? GAP : 0));
  const activeLength = Math.max(0, (1 - share) * CIRCUMFERENCE - (both ? GAP : 0));
  const label = format(strings.GaugeAria, {
    inactive: formatPercent(split.inactiveBytes, split.totalBytes),
    active: formatPercent(split.activeBytes, split.totalBytes)
  });
  return (
    <div className={styles.gauge} role="img" aria-label={label}>
      <svg viewBox="0 0 180 180" width="180" height="180" aria-hidden="true">
        <circle className={styles.gaugeTrack} cx="90" cy="90" r={RADIUS} />
        {split.totalBytes > 0 && inactiveLength > 0 && (
          <circle
            className={`${styles.gaugeArc} ${styles.gaugeInactive}`}
            cx="90"
            cy="90"
            r={RADIUS}
            strokeDasharray={`${inactiveLength} ${CIRCUMFERENCE}`}
            strokeDashoffset={both ? -GAP / 2 : 0}
            transform="rotate(-90 90 90)"
          />
        )}
        {split.totalBytes > 0 && activeLength > 0 && (
          <circle
            className={`${styles.gaugeArc} ${styles.gaugeActive}`}
            cx="90"
            cy="90"
            r={RADIUS}
            strokeDasharray={`${activeLength} ${CIRCUMFERENCE}`}
            strokeDashoffset={-(share * CIRCUMFERENCE + (both ? GAP / 2 : 0))}
            transform="rotate(-90 90 90)"
          />
        )}
      </svg>
      <div className={styles.gaugeCenter}>
        <span className={styles.gaugeValue}>{split.totalBytes ? `${percent}%` : '–'}</span>
        <span className={styles.gaugeLabel}>{strings.GaugeLabel}</span>
      </div>
    </div>
  );
};

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
        <span className={styles.figureLabel}>{kind === 'inactive' ? strings.InactiveFiles : strings.ActiveFiles}</span>
      </div>
      <div className={styles.figureValue} title={`${bytes.toLocaleString()} bytes`}>
        {formatBytes(animated)}
      </div>
      <div className={styles.figureMeta}>
        {format(strings.ShareOfStorage, { percent: formatPercent(bytes, total), files: formatCompact(files) })}
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
    let hiddenItems = 0;
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
      hiddenItems += lib.unreadItems || 0;
    }
    failed += result.webs.filter((w) => w.error).length;
    return { versionBytes, librariesWithMetrics, dormantLibraries, dormantBytes, failed, hiddenItems };
  }, [result, thresholdMonths]);

  const inactiveShare = split.totalBytes ? split.inactiveBytes / split.totalBytes : 0;
  const values = {
    percent: formatPercent(split.inactiveBytes, split.totalBytes),
    size: formatBytes(split.inactiveBytes),
    period
  };

  return (
    <section className={`${styles.card} ${entering ? styles.enter : ''}`} aria-label={strings.InactiveFiles}>
      <div className={styles.summaryTop}>
        <Gauge split={split} />
        <div className={styles.figures}>
          <Figure
            kind="inactive"
            bytes={split.inactiveBytes}
            files={split.inactiveFiles}
            total={split.totalBytes}
            caption={format(strings.NotModifiedFor, { period })}
          />
          <Figure
            kind="active"
            bytes={split.activeBytes}
            files={split.activeFiles}
            total={split.totalBytes}
            caption={format(strings.ModifiedInLast, { period })}
          />
          <div className={styles.figure}>
            <div className={styles.figureHead}>
              <Icon iconName="Database" className={styles.figureIcon} />
              <span className={styles.figureLabel}>{strings.AllFiles}</span>
            </div>
            <div className={styles.figureValue}>{formatBytes(split.totalBytes)}</div>
            <div className={styles.figureMeta}>
              {format(strings.FilesInLibraries, {
                files: formatCompact(split.totalFiles),
                libraries: result.libraries.length
              })}
            </div>
            <div className={styles.figureCaption}>
              {result.siteStorageBytes !== undefined
                ? format(strings.SiteCollectionStorage, { size: formatBytes(result.siteStorageBytes) })
                : strings.CurrentVersions}
            </div>
          </div>
        </div>
      </div>

      <ul className={styles.insights}>
        {split.totalBytes > 0 && (
          <li>
            <Icon iconName="Lightbulb" className={styles.insightIcon} />
            <span>{format(inactiveShare >= 0.5 ? strings.InsightMostlyInactive : strings.InsightSomeInactive, values)}</span>
          </li>
        )}
        {facts.dormantLibraries > 0 && (
          <li>
            <Icon iconName="DocLibrary" className={styles.insightIcon} />
            <span>
              {format(facts.dormantLibraries === 1 ? strings.InsightDormantOne : strings.InsightDormantMany, {
                count: facts.dormantLibraries,
                period,
                size: formatBytes(facts.dormantBytes)
              })}
            </span>
          </li>
        )}
        {facts.librariesWithMetrics > 0 && facts.versionBytes > 0 && (
          <li>
            <Icon iconName="History" className={styles.insightIcon} />
            <span>{format(strings.InsightVersions, { size: formatBytes(facts.versionBytes) })}</span>
          </li>
        )}
        {facts.failed > 0 && (
          <li>
            <Icon iconName="Warning" className={styles.insightIcon} />
            <span>{format(facts.failed === 1 ? strings.InsightFailedOne : strings.InsightFailedMany, { count: facts.failed })}</span>
          </li>
        )}
        {facts.hiddenItems > 0 && (
          <li>
            <Icon iconName="Info" className={styles.insightIcon} />
            <span>
              {format(strings.InsightHidden, {
                count: formatCompact(facts.hiddenItems),
                name: result.scannedBy || strings.TheScanner
              })}
            </span>
          </li>
        )}
      </ul>
    </section>
  );
};
