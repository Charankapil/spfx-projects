import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { Icon } from '@fluentui/react';
import * as strings from 'StoragePulseWebPartStrings';

import { IScanProgress } from '../models/IScanResult';
import { formatBytes } from '../services/formatBytes';
import { formatCompact, formatDuration } from './dashboard/format';
import { PulseMark } from './PulseMark';
import styles from './StoragePulse.module.scss';
import { format } from './text';

interface IScanPanelProps {
  progress: IScanProgress;
  isSaving: boolean;
}

function readable(url: string): string {
  try {
    return decodeURIComponent(url);
  } catch {
    return url;
  }
}

const RATE_WINDOW_MS = 60000;

const STEPS = [strings.ScanStepDiscover, strings.ScanStepRead, strings.ScanStepAnalyse, strings.ScanStepSave];

/** Live view of a running scan: radar, step tracker, progress and speed / time-left figures. */
export const ScanPanel: React.FC<IScanPanelProps> = ({ progress, isSaving }) => {
  const [now, setNow] = useState(Date.now());
  const startedAt = useRef(Date.now());
  const readStartedAt = useRef<number | undefined>(undefined);
  const lastAnnounced = useRef(-1);
  const rateSamples = useRef<{ t: number; items: number }[]>([]);
  const [announcement, setAnnouncement] = useState('');

  useEffect(() => {
    const timer = window.setInterval(() => setNow(Date.now()), 1000);
    return () => window.clearInterval(timer);
  }, []);

  if (progress.phase === 'reading-files' && readStartedAt.current === undefined) {
    readStartedAt.current = Date.now();
  }

  let step = 0;
  if (isSaving) {
    step = 3;
  } else if (progress.phase === 'completed') {
    step = 2;
  } else if (progress.phase === 'reading-files') {
    step = 1;
  }

  const reading = progress.phase === 'reading-files';
  const fraction =
    reading && progress.itemsExpected > 0 ? Math.min(1, progress.itemsRead / progress.itemsExpected) : undefined;
  // Speed over the last minute, not the whole run: the first libraries are
  // mostly per-library overhead and would make the estimate far too long.
  const samples = rateSamples.current;
  if (reading && (samples.length === 0 || samples[samples.length - 1].items !== progress.itemsRead)) {
    samples.push({ t: Date.now(), items: progress.itemsRead });
  }
  while (samples.length > 2 && now - samples[1].t > RATE_WINDOW_MS) {
    samples.shift();
  }
  const first = samples[0];
  const windowSeconds = first ? (now - first.t) / 1000 : 0;
  const readSeconds = readStartedAt.current ? (now - readStartedAt.current) / 1000 : 0;
  const rate = readSeconds >= 5 && windowSeconds > 0 ? (progress.itemsRead - first.items) / windowSeconds : 0;
  const remaining = Math.max(0, progress.itemsExpected - progress.itemsRead);
  const eta = rate > 0 && progress.itemsRead >= 1000 ? (remaining / rate) * 1000 : undefined;

  // Screen readers hear progress in 10% steps rather than every page.
  useEffect(() => {
    if (fraction === undefined) {
      return;
    }
    const decile = Math.floor(fraction * 10);
    if (decile !== lastAnnounced.current) {
      lastAnnounced.current = decile;
      setAnnouncement(format(strings.LiveProgress, { percent: `${decile * 10}%` }));
    }
  }, [fraction]);

  // SharePoint has asked the scan to slow down: say what is happening and that nothing is lost.
  const throttle = progress.throttle;
  const secondsLeft = throttle ? Math.max(0, Math.ceil((throttle.pausedUntil - now) / 1000)) : 0;
  const throttledNow = !!throttle && secondsLeft > 0;

  const stats: { label: string; value: string }[] = reading
    ? [
        { label: strings.StatFiles, value: formatCompact(progress.itemsRead) },
        { label: strings.StatData, value: formatBytes(progress.bytesRead) },
        { label: strings.StatRate, value: rate > 0 ? format(strings.RateValue, { rate: formatCompact(Math.round(rate)) }) : '–' },
        { label: strings.StatEta, value: eta !== undefined ? formatDuration(eta) : strings.EtaCalculating },
        {
          label: strings.StatLibraries,
          value: format(strings.LibrariesValue, {
            done: Math.min(progress.librariesDone + 1, progress.librariesFound),
            total: progress.librariesFound
          })
        }
      ]
    : [
        {
          label: strings.StatLibraries,
          value: format(strings.DiscoveringValue, { webs: progress.websFound, libraries: progress.librariesFound })
        },
        { label: strings.StatElapsed, value: formatDuration(now - startedAt.current) }
      ];

  return (
    <section className={styles.scanPanel} aria-label={strings.ScanningTitle} aria-busy="true">
      <div className={styles.radar} aria-hidden="true">
        <span className={styles.radarSweep} />
        <span className={styles.radarRing} />
        <PulseMark size={56} busy />
      </div>

      <div className={styles.scanBody}>
        <ol className={styles.steps}>
          {STEPS.map((label, i) => (
            <li
              key={label}
              className={`${styles.step} ${i < step ? styles.stepDone : i === step ? styles.stepActive : ''}`}
              aria-current={i === step ? 'step' : undefined}
            >
              <span className={styles.stepDot}>{i < step ? <Icon iconName="CheckMark" /> : i + 1}</span>
              <span>{label}</span>
            </li>
          ))}
        </ol>

        <div
          className={`${styles.track} ${fraction === undefined ? styles.trackIndeterminate : ''}`}
          role="progressbar"
          aria-valuemin={0}
          aria-valuemax={100}
          aria-valuenow={fraction === undefined ? undefined : Math.round(fraction * 100)}
        >
          <span className={styles.trackFill} style={fraction === undefined ? undefined : { width: `${fraction * 100}%` }} />
        </div>

        <dl className={styles.stats}>
          {stats.map((s) => (
            <div key={s.label} className={styles.stat}>
              <dt>{s.label}</dt>
              <dd>{s.value}</dd>
            </div>
          ))}
        </dl>

        {throttle && (
          <div className={`${styles.throttleNote} ${throttledNow ? styles.throttleActive : ''}`} role="status">
            <Icon iconName="Clock" className={styles.throttleIcon} />
            <div>
              <strong>{strings.ThrottledTitle}</strong>
              {throttledNow && <div>{format(strings.ThrottledWaiting, { time: formatDuration(secondsLeft * 1000) })}</div>}
              {throttle.concurrency < throttle.maxConcurrency && (
                <div>{format(strings.ThrottledSlower, { now: throttle.concurrency, max: throttle.maxConcurrency })}</div>
              )}
              <div>{format(strings.ThrottledCount, { count: throttle.throttledCount })}</div>
            </div>
          </div>
        )}
        <div className={styles.currentItem} title={readable(progress.currentItem)}>
          {readable(progress.currentItem)}
        </div>
        <div className={styles.scanHint}>{strings.ScanHint}</div>
        <div className={styles.srOnly} aria-live="polite">
          {announcement}
        </div>
      </div>
    </section>
  );
};
