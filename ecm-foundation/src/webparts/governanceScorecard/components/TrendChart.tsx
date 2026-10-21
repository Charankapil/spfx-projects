import * as React from 'react';

import { formatPct } from '../services/metrics';
import styles from './GovernanceScorecard.module.scss';
import { INK, SERIES } from './vizTokens';

export interface ITrendPoint {
  label: string;
  score?: number;
  metadata?: number;
  labelled?: number;
  unsaved?: boolean;
}

type SeriesKey = 'score' | 'metadata' | 'labelled';

const SERIES_DEFS: { key: SeriesKey; name: string; color: string }[] = [
  { key: 'score', name: 'Governance score', color: SERIES.score },
  { key: 'metadata', name: 'Metadata complete', color: SERIES.metadata },
  { key: 'labelled', name: 'Retention labelled', color: SERIES.labelled }
];

const HEIGHT = 240;
const MARGIN = { top: 14, right: 150, bottom: 30, left: 40 };
const TICKS = [0, 25, 50, 75, 100];

/**
 * Three percentages on one 0–100 axis over the saved snapshots. Legend always
 * shown; end labels only where they do not collide; a crosshair tooltip on
 * hover and keyboard focus; and a table view with every value.
 */
export const TrendChart = (props: { points: ITrendPoint[] }): JSX.Element => {
  const { points } = props;
  const wrapRef = React.useRef<HTMLDivElement>(null);
  const [width, setWidth] = React.useState(720);
  const [hover, setHover] = React.useState<number | undefined>(undefined);
  const [asTable, setAsTable] = React.useState(false);

  React.useEffect(() => {
    const el = wrapRef.current;
    if (!el) {
      return undefined;
    }
    const measure = (): void => setWidth(Math.max(320, el.clientWidth));
    measure();
    const RO = (window as unknown as { ResizeObserver?: new (cb: () => void) => { observe: (e: Element) => void; disconnect: () => void } }).ResizeObserver;
    if (RO) {
      const ro = new RO(measure);
      ro.observe(el);
      return () => ro.disconnect();
    }
    window.addEventListener('resize', measure);
    return () => window.removeEventListener('resize', measure);
  }, [asTable]);

  const plotW = width - MARGIN.left - MARGIN.right;
  const plotH = HEIGHT - MARGIN.top - MARGIN.bottom;
  const x = (i: number): number => MARGIN.left + (points.length <= 1 ? plotW / 2 : (i / (points.length - 1)) * plotW);
  const y = (v: number): number => MARGIN.top + plotH - (v / 100) * plotH;

  const paths = SERIES_DEFS.map((s) => {
    let d = '';
    let pen = false;
    points.forEach((p, i) => {
      const v = p[s.key];
      if (v === undefined) {
        pen = false;
        return;
      }
      d += `${pen ? 'L' : 'M'}${x(i).toFixed(1)},${y(v).toFixed(1)} `;
      pen = true;
    });
    return { ...s, d };
  });

  // End labels: drop them all if any two would sit closer than one line of text.
  const last = points.length - 1;
  const ends = SERIES_DEFS.map((s) => ({ ...s, v: last >= 0 ? points[last][s.key] : undefined })).filter((e) => e.v !== undefined) as {
    key: SeriesKey;
    name: string;
    color: string;
    v: number;
  }[];
  const endYs = ends.map((e) => y(e.v)).sort((a, b) => a - b);
  const endLabelsFit = endYs.every((v, i) => i === 0 || v - endYs[i - 1] >= 15);

  // Thin out x labels so they never overlap (about 90px each).
  const labelEvery = Math.max(1, Math.ceil(points.length / Math.max(1, Math.floor(plotW / 90))));

  const pick = (clientX: number, rect: DOMRect): void => {
    if (!points.length) {
      return;
    }
    const rel = clientX - rect.left - MARGIN.left;
    const i = points.length <= 1 ? 0 : Math.round((rel / plotW) * (points.length - 1));
    setHover(Math.max(0, Math.min(points.length - 1, i)));
  };

  const onKey = (e: React.KeyboardEvent): void => {
    if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
      e.preventDefault();
      const start = hover === undefined ? last : hover;
      setHover(Math.max(0, Math.min(last, start + (e.key === 'ArrowRight' ? 1 : -1))));
    }
  };

  const hovered = hover !== undefined ? points[hover] : undefined;
  const tooltipLeft = hover !== undefined ? Math.min(x(hover) + 12, width - 200) : 0;

  return (
    <section className={styles.card} aria-labelledby="ecm-trend-title">
      <div className={styles.cardHead}>
        <div>
          <h3 id="ecm-trend-title" className={styles.cardTitle}>
            Trend across saved snapshots
          </h3>
          <p className={styles.cardSub}>Percent of scanned documents (score out of 100).</p>
        </div>
        <button type="button" className={styles.linkButton} onClick={() => setAsTable(!asTable)} aria-pressed={asTable}>
          {asTable ? 'Show as chart' : 'Show as table'}
        </button>
      </div>

      <ul className={styles.legend} aria-label="Legend">
        {SERIES_DEFS.map((s) => (
          <li key={s.key}>
            <svg width="16" height="4" aria-hidden="true">
              <line x1="0" y1="2" x2="16" y2="2" stroke={s.color} strokeWidth="2" strokeLinecap="round" />
            </svg>
            {s.name}
          </li>
        ))}
      </ul>

      {points.length < 2 && (
        <p className={styles.hint}>The trend line appears once there are two scans. Save a snapshot after each scan to build the baseline.</p>
      )}

      {asTable ? (
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th>Scan</th>
                {SERIES_DEFS.map((s) => (
                  <th key={s.key} className={styles.num}>
                    {s.name}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {points.map((p, i) => (
                <tr key={i}>
                  <td>
                    {p.label}
                    {p.unsaved ? ' (not saved)' : ''}
                  </td>
                  {SERIES_DEFS.map((s) => (
                    <td key={s.key} className={styles.num}>
                      {s.key === 'score' ? (p.score === undefined ? '–' : Math.round(p.score)) : formatPct(p[s.key])}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <div className={styles.chartWrap} ref={wrapRef}>
          <svg width={width} height={HEIGHT} role="img" aria-label="Line chart of governance score, metadata complete and retention labelled per scan">
            {TICKS.map((t) => (
              <g key={t}>
                <line x1={MARGIN.left} x2={MARGIN.left + plotW} y1={y(t)} y2={y(t)} stroke={t === 0 ? INK.axis : INK.grid} strokeWidth="1" />
                <text x={MARGIN.left - 8} y={y(t) + 4} textAnchor="end" className={styles.axisText}>
                  {t}
                </text>
              </g>
            ))}
            {points.map((p, i) =>
              i % labelEvery === 0 || i === last ? (
                <text key={i} x={x(i)} y={HEIGHT - 8} textAnchor="middle" className={styles.axisText}>
                  {p.label}
                </text>
              ) : null
            )}
            {hover !== undefined && <line x1={x(hover)} x2={x(hover)} y1={MARGIN.top} y2={MARGIN.top + plotH} stroke={INK.axis} strokeWidth="1" />}
            {paths.map((s) => (
              <path key={s.key} d={s.d} fill="none" stroke={s.color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
            ))}
            {ends.map((e) => (
              <circle key={e.key} cx={x(last)} cy={y(e.v)} r="4" fill={e.color} stroke="#fff" strokeWidth="2" />
            ))}
            {hover !== undefined &&
              hover !== last &&
              SERIES_DEFS.map((s) => {
                const v = points[hover][s.key];
                return v === undefined ? null : <circle key={s.key} cx={x(hover)} cy={y(v)} r="4" fill={s.color} stroke="#fff" strokeWidth="2" />;
              })}
            {endLabelsFit &&
              ends.map((e) => (
                <text key={e.key} x={x(last) + 10} y={y(e.v) + 4} className={styles.endLabel}>
                  <tspan className={styles.endValue}>{e.key === 'score' ? Math.round(e.v) : `${Math.round(e.v)}%`}</tspan> {e.name}
                </text>
              ))}
            <rect
              x={MARGIN.left}
              y={MARGIN.top}
              width={plotW}
              height={plotH}
              fill="transparent"
              tabIndex={0}
              aria-label="Use the left and right arrow keys to read each scan"
              onPointerMove={(e) => pick(e.clientX, (e.currentTarget.ownerSVGElement as SVGSVGElement).getBoundingClientRect())}
              onPointerLeave={() => setHover(undefined)}
              onFocus={() => setHover(last)}
              onBlur={() => setHover(undefined)}
              onKeyDown={onKey}
              className={styles.hitArea}
            />
          </svg>
          {hovered && (
            <div className={styles.tooltip} style={{ left: tooltipLeft, top: MARGIN.top }} role="status">
              <div className={styles.tooltipTitle}>
                {hovered.label}
                {hovered.unsaved ? ' · not saved' : ''}
              </div>
              {SERIES_DEFS.map((s) => (
                <div key={s.key} className={styles.tooltipRow}>
                  <svg width="12" height="4" aria-hidden="true">
                    <line x1="0" y1="2" x2="12" y2="2" stroke={s.color} strokeWidth="2" strokeLinecap="round" />
                  </svg>
                  <strong>{s.key === 'score' ? (hovered.score === undefined ? '–' : Math.round(hovered.score)) : formatPct(hovered[s.key])}</strong>
                  <span>{s.name}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      )}
    </section>
  );
};
