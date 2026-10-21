import * as React from 'react';

import { IAggregate } from '../models/IScorecard';
import { formatBytes, formatPct, IScoreParts, statusOf } from '../services/metrics';
import styles from './GovernanceScorecard.module.scss';
import { Sparkline } from './Sparkline';
import { StatusChip } from './StatusChip';
import { INK } from './vizTokens';

export interface IKpiHistoryPoint {
  parts: IScoreParts;
  total: IAggregate;
}

export interface IKpiTilesProps {
  parts: IScoreParts;
  total: IAggregate;
  /** Oldest saved snapshot, for "since baseline" deltas; undefined when the displayed run is the baseline. */
  baseline?: { parts: IScoreParts; total: IAggregate; label: string };
  history: IKpiHistoryPoint[];
  historyIndex: number;
  staleYears: number;
}

/** Signed change with its direction in words and an arrow; colour only reinforces it. */
const Delta = (props: { now?: number; then?: number; upIsGood: boolean; unit: 'pts' | 'bytes'; since: string }): JSX.Element | null => {
  if (props.now === undefined || props.then === undefined) {
    return null;
  }
  const diff = props.now - props.then;
  const flat = props.unit === 'pts' ? Math.abs(diff) < 0.5 : diff === 0;
  if (flat) {
    return <span className={styles.delta}>No change since {props.since}</span>;
  }
  const good = diff > 0 === props.upIsGood;
  const amount = props.unit === 'pts' ? `${Math.abs(diff).toFixed(0)} pts` : formatBytes(Math.abs(diff));
  return (
    <span className={styles.delta} style={{ color: good ? INK.deltaGood : INK.deltaBad }}>
      {diff > 0 ? '▲' : '▼'} {diff > 0 ? 'Up' : 'Down'} {amount} since {props.since}
    </span>
  );
};

export const KpiTiles = (props: IKpiTilesProps): JSX.Element => {
  const { parts, total, baseline, history, historyIndex } = props;
  const since = baseline?.label || '';
  const series = (pick: (p: IKpiHistoryPoint) => number | undefined): (number | undefined)[] => history.map(pick);

  return (
    <div className={styles.kpis}>
      <section className={`${styles.tile} ${styles.heroTile}`} aria-label="Governance score">
        <div className={styles.tileLabel}>Governance score</div>
        <div className={styles.heroRow}>
          <span className={styles.heroValue}>{parts.score === undefined ? '–' : Math.round(parts.score)}</span>
          <span className={styles.heroOutOf}>/ 100</span>
        </div>
        <StatusChip status={statusOf(parts.score)} />
        {baseline && <Delta now={parts.score} then={baseline.parts.score} upIsGood unit="pts" since={since} />}
        <Sparkline values={series((h) => h.parts.score)} highlight={historyIndex} label="Score over saved snapshots" />
      </section>

      <section className={styles.tile} aria-label="Metadata complete">
        <div className={styles.tileLabel}>Metadata complete</div>
        <div className={styles.tileValue}>{formatPct(parts.metadata)}</div>
        <div className={styles.tileSub}>
          {total.complete.toLocaleString()} of {total.scanned.toLocaleString()} documents have every required field
        </div>
        {baseline && <Delta now={parts.metadata} then={baseline.parts.metadata} upIsGood unit="pts" since={since} />}
        <Sparkline values={series((h) => h.parts.metadata)} highlight={historyIndex} label="Metadata complete over saved snapshots" />
      </section>

      <section className={styles.tile} aria-label="Retention labelled">
        <div className={styles.tileLabel}>Retention labelled</div>
        <div className={styles.tileValue}>{formatPct(parts.labelled)}</div>
        <div className={styles.tileSub}>{total.labelled.toLocaleString()} documents carry a Purview retention label</div>
        {baseline && <Delta now={parts.labelled} then={baseline.parts.labelled} upIsGood unit="pts" since={since} />}
        <Sparkline values={series((h) => h.parts.labelled)} highlight={historyIndex} label="Retention labelled over saved snapshots" />
      </section>

      <section className={styles.tile} aria-label="Archive candidates">
        <div className={styles.tileLabel}>Archive candidates</div>
        <div className={styles.tileValue}>{formatPct(parts.stale)}</div>
        <div className={styles.tileSub}>
          {total.stale.toLocaleString()} documents ({formatBytes(total.staleBytes)}) not changed in {props.staleYears} years
        </div>
        {baseline && <Delta now={parts.stale} then={baseline.parts.stale} upIsGood={false} unit="pts" since={since} />}
        <Sparkline values={series((h) => h.parts.stale)} highlight={historyIndex} label="Archive candidates over saved snapshots" />
      </section>

      <section className={styles.tile} aria-label="Storage">
        <div className={styles.tileLabel}>Storage used</div>
        <div className={styles.tileValue}>{total.storageKnown ? formatBytes(total.storageBytes) : '–'}</div>
        <div className={styles.tileSub}>
          {total.duplicates.toLocaleString()} likely duplicate copies ({formatBytes(total.duplicateBytes)}) could be removed
        </div>
        {baseline && baseline.total.storageKnown && total.storageKnown && (
          <Delta now={total.storageBytes} then={baseline.total.storageBytes} upIsGood={false} unit="bytes" since={since} />
        )}
        <Sparkline
          values={series((h) => (h.total.storageKnown ? h.total.storageBytes : undefined))}
          highlight={historyIndex}
          label="Storage over saved snapshots"
        />
      </section>
    </div>
  );
};
