import * as React from 'react';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

import { IScanProgress, IScorecardRun, ISiteMetrics, ISnapshot } from '../models/IScorecard';
import { aggregate, formatBytes, scoreOf, siteScore, WEIGHTS } from '../services/metrics';
import { ScorecardScanner } from '../services/ScorecardScanner';
import { SnapshotStore } from '../services/SnapshotStore';
import { FieldCoverage } from './FieldCoverage';
import styles from './GovernanceScorecard.module.scss';
import { HubTable } from './HubTable';
import { KpiTiles } from './KpiTiles';
import { ITrendPoint, TrendChart } from './TrendChart';

export interface IGovernanceScorecardProps {
  title: string;
  scanner: ScorecardScanner;
  store: SnapshotStore;
  staleYears: number;
}

const CURRENT = 'current';
const ALL = '__all';

function shortDate(iso: string, withYear: boolean): string {
  const d = new Date(iso);
  return d.toLocaleDateString(undefined, withYear ? { day: 'numeric', month: 'short', year: 'numeric' } : { day: 'numeric', month: 'short' });
}

function csvCell(value: string): string {
  return /[",\r\n]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

function toCsv(run: IScorecardRun, sites: ISiteMetrics[]): string {
  const head = ['Hub', 'Site', 'URL', 'Score', 'Documents', 'Scanned', 'Metadata complete %', 'Retention labelled %', 'Archive candidates %', 'Stale bytes', 'Duplicate copies', 'Duplicate bytes', 'Storage bytes', 'Owners', 'Missing columns', 'Error']
    .concat(run.requiredFields.map((f) => `${f} filled %`));
  const fmt = (v: number | undefined): string => (v === undefined ? '' : v.toFixed(1));
  const rows = sites.map((s) => {
    const p = siteScore(s, run.requiredFields);
    return [s.hub, s.title, s.url, fmt(p.score), String(s.documents), String(s.scanned), fmt(p.metadata), fmt(p.labelled), fmt(p.stale), String(s.staleBytes), String(s.duplicates), String(s.duplicateBytes), s.storageBytes === undefined ? '' : String(s.storageBytes), s.ownership, s.missingColumns.join('; '), s.error || '']
      .concat(run.requiredFields.map((f) => fmt(s.scanned ? ((s.fieldFilled[f] || 0) / s.scanned) * 100 : undefined)))
      .map(csvCell)
      .join(',');
  });
  return [head.map(csvCell).join(',')].concat(rows).join('\r\n');
}

function download(text: string, name: string): void {
  const blob = new Blob(['﻿' + text], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export const GovernanceScorecard = (props: IGovernanceScorecardProps): JSX.Element => {
  const { scanner, store } = props;
  const [snapshots, setSnapshots] = React.useState<ISnapshot[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [loadError, setLoadError] = React.useState<string | undefined>(undefined);
  const [current, setCurrent] = React.useState<IScorecardRun | undefined>(undefined);
  const [progress, setProgress] = React.useState<IScanProgress | undefined>(undefined);
  const [scanError, setScanError] = React.useState<string | undefined>(undefined);
  const [view, setView] = React.useState<string>(CURRENT);
  const [hub, setHub] = React.useState<string>(ALL);
  const [saving, setSaving] = React.useState(false);
  const [saveMessage, setSaveMessage] = React.useState<{ ok: boolean; text: string } | undefined>(undefined);

  const loadSnapshots = React.useCallback(async () => {
    setLoading(true);
    try {
      setSnapshots(await store.load());
      setLoadError(undefined);
    } catch (err) {
      setLoadError((err as Error).message);
    } finally {
      setLoading(false);
    }
  }, [store]);

  React.useEffect(() => {
    loadSnapshots().catch(() => undefined);
  }, [loadSnapshots]);

  const runScan = async (): Promise<void> => {
    setScanError(undefined);
    setSaveMessage(undefined);
    setProgress({ done: 0, total: 0, current: '' });
    try {
      const run = await scanner.scan(setProgress, () => undefined);
      setCurrent(run);
      setView(CURRENT);
    } catch (err) {
      setScanError((err as Error).message);
    } finally {
      setProgress(undefined);
    }
  };

  const save = async (): Promise<void> => {
    if (!current) {
      return;
    }
    setSaving(true);
    setSaveMessage(undefined);
    try {
      await store.save(current);
      setCurrent(undefined);
      setView('0');
      await loadSnapshots();
      setSaveMessage({ ok: true, text: 'Snapshot saved. Everyone who opens this page now sees these results.' });
    } catch (err) {
      const status = (err as { status?: number }).status;
      setSaveMessage({
        ok: false,
        text:
          status === 401 || status === 403
            ? 'You cannot save snapshots on this site. A site owner of the hub can, or can give you edit access to the "ECM Scorecard Snapshots" list.'
            : (err as Error).message
      });
    } finally {
      setSaving(false);
    }
  };

  // Oldest first, with an unsaved scan at the end.
  const chronological: { run: IScorecardRun; key: string; unsaved: boolean }[] = snapshots
    .map((s, i) => ({ run: s as IScorecardRun, key: String(i), unsaved: false }))
    .reverse();
  if (current) {
    chronological.push({ run: current, key: CURRENT, unsaved: true });
  }
  const displayedEntry = chronological.filter((c) => c.key === view)[0] || chronological[chronological.length - 1];
  const displayed = displayedEntry?.run;

  const hubs: string[] = [];
  (displayed?.sites || []).forEach((s) => {
    if (hubs.indexOf(s.hub) < 0) {
      hubs.push(s.hub);
    }
  });
  const scope = (run: IScorecardRun): ISiteMetrics[] => (hub === ALL ? run.sites : run.sites.filter((s) => s.hub === hub));

  const history = chronological.map((c) => {
    const total = aggregate(scope(c.run), c.run.requiredFields);
    return { total, parts: scoreOf(total) };
  });
  const historyIndex = displayedEntry ? chronological.indexOf(displayedEntry) : -1;
  const years = chronological.map((c) => new Date(c.run.runAt).getFullYear());
  const withYear = years.some((y) => y !== years[0]);
  const trendPoints: ITrendPoint[] = chronological.map((c, i) => ({
    label: shortDate(c.run.runAt, withYear),
    score: history[i].parts.score,
    metadata: history[i].parts.metadata,
    labelled: history[i].parts.labelled,
    unsaved: c.unsaved
  }));

  const baselineEntry = chronological.length > 1 && historyIndex > 0 ? chronological[0] : undefined;
  const sites = displayed ? scope(displayed) : [];
  const total = displayed ? aggregate(sites, displayed.requiredFields) : undefined;
  const parts = total ? scoreOf(total) : undefined;
  const failedSites = sites.filter((s) => s.error).length;
  const sampledSites = sites.filter((s) => s.sampled).length;

  return (
    <div className={styles.scorecard}>
      <header className={styles.header}>
        <div>
          <h2 className={styles.title}>{props.title}</h2>
          <p className={styles.subtitle}>
            {displayed
              ? `${displayedEntry.unsaved ? 'Unsaved scan' : 'Snapshot'} from ${new Date(displayed.runAt).toLocaleString()}${
                  !displayedEntry.unsaved && (displayed as ISnapshot).savedBy ? ` · saved by ${(displayed as ISnapshot).savedBy}` : ''
                }`
              : 'Measures the ECM pilot targets: metadata, retention labels, stale content, storage and ownership.'}
          </p>
        </div>
        <div className={styles.actions}>
          <PrimaryButton text={progress ? 'Scanning…' : 'Run scan'} disabled={!!progress} onClick={() => runScan()} />
          {current && <DefaultButton text={saving ? 'Saving…' : 'Save snapshot'} disabled={saving || !!progress} onClick={() => save()} />}
          {displayed && (
            <DefaultButton
              text="Export CSV"
              disabled={!!progress}
              onClick={() => download(toCsv(displayed, sites), `governance-scorecard-${displayed.runAt.substring(0, 10)}.csv`)}
            />
          )}
        </div>
      </header>

      {progress && (
        <div className={styles.progress} role="status" aria-live="polite">
          <div className={styles.progressText}>
            {progress.total ? `Scanning ${progress.done} of ${progress.total} sites${progress.current ? `: ${progress.current}` : ''}` : 'Reading the navigation list…'}
          </div>
          <div className={styles.progressTrack} aria-hidden="true">
            <span style={{ width: `${progress.total ? (progress.done / progress.total) * 100 : 3}%` }} />
          </div>
        </div>
      )}

      {scanError && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          The scan could not start: {scanError}. Check that this web part points at the site with the ECM Navigation list.
        </MessageBar>
      )}
      {loadError && (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline>
          Saved snapshots could not be read: {loadError}
        </MessageBar>
      )}
      {saveMessage && (
        <MessageBar messageBarType={saveMessage.ok ? MessageBarType.success : MessageBarType.error} isMultiline>
          {saveMessage.text}
        </MessageBar>
      )}
      {current && !progress && (
        <MessageBar messageBarType={MessageBarType.info} isMultiline>
          These results are not saved yet. Save a snapshot to keep them, share them with everyone who opens this page and track progress against
          the baseline.
        </MessageBar>
      )}

      {!displayed && !loading && !progress && (
        <section className={styles.intro}>
          <h3>No scans yet</h3>
          <p>
            Run a scan to measure every site in the ECM navigation. The scorecard reads each site&apos;s document libraries and reports, per hub and
            site:
          </p>
          <ul>
            <li>
              <b>Metadata complete</b>: documents with every required field filled in ({WEIGHTS.metadata}% of the score).
            </li>
            <li>
              <b>Retention labelled</b>: documents with a Purview retention label ({WEIGHTS.labelled}%).
            </li>
            <li>
              <b>Archive candidates</b>: documents not changed in {props.staleYears} years (counts against {WEIGHTS.fresh}%).
            </li>
            <li>
              <b>Ownership</b>: owners assigned through a group, or at least two people ({WEIGHTS.ownership}%).
            </li>
            <li>
              <b>Storage</b>: site storage and likely duplicate copies.
            </li>
          </ul>
          <p>The first saved snapshot becomes the baseline for the pilot&apos;s value targets.</p>
        </section>
      )}

      {displayed && total && parts && (
        <div className={progress ? styles.stale : undefined}>
          <div className={styles.filters} role="group" aria-label="Filters">
            <label className={styles.filter}>
              <span>Hub</span>
              <select value={hub} onChange={(e) => setHub(e.target.value)}>
                <option value={ALL}>All hubs</option>
                {hubs.map((h) => (
                  <option key={h} value={h}>
                    {h}
                  </option>
                ))}
              </select>
            </label>
            <label className={styles.filter}>
              <span>Scan</span>
              <select value={displayedEntry.key} onChange={(e) => setView(e.target.value)}>
                {chronological
                  .slice()
                  .reverse()
                  .map((c) => (
                    <option key={c.key} value={c.key}>
                      {new Date(c.run.runAt).toLocaleString()}
                      {c.unsaved ? ' (not saved)' : c === chronological[0] && chronological.length > 1 ? ' (baseline)' : ''}
                    </option>
                  ))}
              </select>
            </label>
            <span className={styles.filterNote}>
              Required metadata: {displayed.requiredFields.join(', ') || 'none set'} · {total.scannedSites} of {total.sites} sites scanned
              {failedSites ? ` · ${failedSites} could not be scanned` : ''}
              {sampledSites ? ` · ${sampledSites} sampled` : ''}
            </span>
          </div>

          <KpiTiles
            parts={parts}
            total={total}
            baseline={
              baselineEntry
                ? { parts: history[0].parts, total: history[0].total, label: `baseline (${shortDate(baselineEntry.run.runAt, true)})` }
                : undefined
            }
            history={history}
            historyIndex={historyIndex}
            staleYears={displayed.staleYears}
          />

          <div className={styles.twoCol}>
            <TrendChart points={trendPoints} />
            <FieldCoverage total={total} requiredFields={displayed.requiredFields} />
          </div>

          <HubTable sites={sites} requiredFields={displayed.requiredFields} />

          <p className={styles.footnote}>
            Score = {WEIGHTS.metadata}% metadata complete + {WEIGHTS.labelled}% retention labelled + {WEIGHTS.fresh}% not stale +{' '}
            {WEIGHTS.ownership}% ownership; 80 and above is on track, 60 to 79 needs attention, below 60 is at risk. Measured as{' '}
            {displayedEntry.unsaved ? 'you' : (displayed as ISnapshot).savedBy || 'the person who ran the scan'}: sites they cannot open are listed as
            not scanned. Likely duplicates share a file name and size within a site ({formatBytes(total.duplicateBytes)} found).
          </p>
        </div>
      )}
    </div>
  );
};
