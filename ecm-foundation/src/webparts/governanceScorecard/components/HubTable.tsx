import * as React from 'react';

import { IAggregate, ISiteMetrics, OwnershipKind } from '../models/IScorecard';
import { aggregate, formatBytes, formatPct, IScoreParts, scoreOf, siteScore, statusOf } from '../services/metrics';
import styles from './GovernanceScorecard.module.scss';
import { StatusChip } from './StatusChip';
import { STATUS_COLOR } from './vizTokens';

const OWNERSHIP_LABEL: { [k in OwnershipKind]: string } = {
  group: 'Group',
  people: 'People',
  single: '1 person',
  none: 'None',
  unknown: '–'
};

/** A percentage as a meter coloured by its status, with the number beside it (colour is never the only cue). */
const Meter = (props: { value?: number; lowerIsBetter?: boolean }): JSX.Element => {
  if (props.value === undefined) {
    return <span className={styles.muted}>–</span>;
  }
  const status = statusOf(props.lowerIsBetter ? 100 - props.value : props.value);
  return (
    <span className={styles.meterCell}>
      <span className={styles.meter} aria-hidden="true">
        <span style={{ width: `${Math.max(2, props.value)}%`, background: status ? STATUS_COLOR[status] : undefined }} />
      </span>
      <span className={styles.meterValue}>{formatPct(props.value)}</span>
    </span>
  );
};

const MetricCells = (props: { parts: IScoreParts; total: IAggregate; ownership?: string }): JSX.Element => (
  <>
    <td>
      <StatusChip status={statusOf(props.parts.score)} />
    </td>
    <td className={styles.num}>{props.parts.score === undefined ? '–' : Math.round(props.parts.score)}</td>
    <td className={styles.num}>{props.total.documents.toLocaleString()}</td>
    <td>
      <Meter value={props.parts.metadata} />
    </td>
    <td>
      <Meter value={props.parts.labelled} />
    </td>
    <td>
      <Meter value={props.parts.stale} lowerIsBetter />
    </td>
    <td className={styles.num}>{props.total.storageKnown ? formatBytes(props.total.storageBytes) : '–'}</td>
    <td>{props.ownership}</td>
  </>
);

export interface IHubTableProps {
  sites: ISiteMetrics[];
  requiredFields: string[];
}

/** Hubs with their totals; open a hub to see its sites, weakest first. */
export const HubTable = (props: IHubTableProps): JSX.Element => {
  const hubs: string[] = [];
  props.sites.forEach((s) => {
    if (hubs.indexOf(s.hub) < 0) {
      hubs.push(s.hub);
    }
  });
  const [open, setOpen] = React.useState<{ [hub: string]: boolean }>({});

  return (
    <section className={styles.card} aria-labelledby="ecm-hub-table">
      <h3 id="ecm-hub-table" className={styles.cardTitle}>
        By hub and site
      </h3>
      <p className={styles.cardSub}>Open a hub to see its sites, weakest first.</p>
      <div className={styles.tableWrap}>
        <table className={styles.table}>
          <thead>
            <tr>
              <th>Hub / site</th>
              <th>Status</th>
              <th className={styles.num}>Score</th>
              <th className={styles.num}>Documents</th>
              <th>Metadata complete</th>
              <th>Retention labelled</th>
              <th>Archive candidates</th>
              <th className={styles.num}>Storage</th>
              <th>Owners</th>
            </tr>
          </thead>
          <tbody>
            {hubs.map((hub) => {
              const sites = props.sites.filter((s) => s.hub === hub);
              const total = aggregate(sites, props.requiredFields);
              const parts = scoreOf(total);
              const expanded = !!open[hub];
              const ranked = sites
                .map((s) => ({ s, parts: siteScore(s, props.requiredFields) }))
                .sort((a, b) => (a.parts.score ?? 101) - (b.parts.score ?? 101));
              return (
                <React.Fragment key={hub}>
                  <tr className={styles.hubRow}>
                    <td>
                      <button type="button" className={styles.expand} aria-expanded={expanded} onClick={() => setOpen({ ...open, [hub]: !expanded })}>
                        <span className={`${styles.chevron} ${expanded ? styles.chevronOpen : ''}`} aria-hidden="true">
                          ›
                        </span>
                        {hub}
                        <span className={styles.siteCount}>
                          {sites.length} {sites.length === 1 ? 'site' : 'sites'}
                        </span>
                      </button>
                    </td>
                    <MetricCells
                      parts={parts}
                      total={total}
                      ownership={total.ownersKnown ? `${total.ownersOk} of ${total.ownersKnown} OK` : '–'}
                    />
                  </tr>
                  {expanded &&
                    ranked.map(({ s, parts: sp }) => (
                      <tr key={s.url} className={styles.siteRow}>
                        <td>
                          <a href={s.url} target="_blank" rel="noopener noreferrer" data-interception="off">
                            {s.title}
                          </a>
                          {s.sampled && (
                            <span className={styles.tag} title={`Measured on the first ${s.scanned.toLocaleString()} of ${s.documents.toLocaleString()} items`}>
                              Sampled
                            </span>
                          )}
                          {s.missingColumns.length > 0 && !s.error && (
                            <span className={styles.tag} title={`Missing columns: ${s.missingColumns.join(', ')}`}>
                              {s.missingColumns.length} {s.missingColumns.length === 1 ? 'column' : 'columns'} missing
                            </span>
                          )}
                        </td>
                        {s.error ? (
                          <td colSpan={8} className={styles.errorCell}>
                            Couldn&apos;t scan: {s.error}
                          </td>
                        ) : (
                          <MetricCells
                            parts={sp}
                            total={aggregate([s], props.requiredFields)}
                            ownership={s.ownership === 'people' || s.ownership === 'single' ? `${OWNERSHIP_LABEL[s.ownership]} (${s.owners})` : OWNERSHIP_LABEL[s.ownership]}
                          />
                        )}
                      </tr>
                    ))}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    </section>
  );
};
