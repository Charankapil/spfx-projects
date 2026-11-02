import * as React from 'react';
import { IconButton, Text } from '@fluentui/react';
import { useApp } from '../AppContext';
import { chartInk, otherColor, seriesColor } from '../palette';
import { fmtCompact, fmtNumber } from '../../logic/format';
import { downloadSvgAsPng } from '../../export/svgToPng';
import { Card, Empty } from './Layout';

/* ---------- helpers ---------- */

function useWidth<T extends HTMLElement>(): [React.RefObject<T>, number] {
  const ref = React.useRef<T>(null);
  const [width, setWidth] = React.useState(0);
  React.useEffect(() => {
    const el = ref.current;
    if (!el) return undefined;
    const update = (): void => setWidth(Math.floor(el.getBoundingClientRect().width));
    update();
    const RO = (window as unknown as { ResizeObserver?: typeof ResizeObserver }).ResizeObserver;
    if (RO) {
      const ro = new RO(() => update());
      ro.observe(el);
      return () => ro.disconnect();
    }
    window.addEventListener('resize', update);
    return () => window.removeEventListener('resize', update);
  }, []);
  return [ref, width];
}

/** Rounds `max` up to a readable axis maximum and returns evenly spaced ticks. */
export function niceTicks(max: number, count: number = 4): number[] {
  if (max <= 0) return [0, 1];
  const raw = max / count;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const norm = raw / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
  const ticks: number[] = [];
  for (let v = 0; v <= max + step * 0.001; v += step) ticks.push(Math.round(v * 1000) / 1000);
  if (ticks[ticks.length - 1] < max) ticks.push(ticks[ticks.length - 1] + step);
  return ticks;
}

const FONT = 'Segoe UI, -apple-system, BlinkMacSystemFont, Roboto, Helvetica Neue, sans-serif';

interface ITooltip {
  x: number;
  y: number;
  title: string;
  rows: { label: string; value: string; color?: string }[];
}

const Tooltip: React.FC<{ tip: ITooltip | undefined; width: number }> = ({ tip, width }) => {
  const { theme } = useApp();
  if (!tip) return null;
  const left = Math.min(Math.max(tip.x + 12, 0), Math.max(0, width - 200));
  return (
    <div
      role="status"
      style={{
        position: 'absolute', left, top: Math.max(0, tip.y - 10), pointerEvents: 'none', zIndex: 2, minWidth: 140, maxWidth: 240,
        background: theme.semanticColors.bodyBackground, color: theme.semanticColors.bodyText,
        border: `1px solid ${theme.palette.neutralQuaternary}`, borderRadius: 4, padding: '6px 8px',
        boxShadow: theme.effects.elevation8, fontSize: 12
      }}
    >
      <div style={{ fontWeight: 600, marginBottom: 2 }}>{tip.title}</div>
      {tip.rows.map((r) => (
        <div key={r.label} style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          {r.color && <span style={{ width: 8, height: 8, borderRadius: 2, background: r.color, flex: 'none' }} />}
          <span style={{ flex: 1 }}>{r.label}</span>
          <span style={{ fontVariantNumeric: 'tabular-nums', fontWeight: 600 }}>{r.value}</span>
        </div>
      ))}
    </div>
  );
};

export const Legend: React.FC<{ items: { label: string; color: string; value?: string }[] }> = ({ items }) => {
  const { theme } = useApp();
  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px 14px', fontSize: 12, color: theme.semanticColors.bodySubtext }}>
      {items.map((i) => (
        <span key={i.label} style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
          <span style={{ width: 10, height: 10, borderRadius: 2, background: i.color }} />
          <span style={{ color: theme.semanticColors.bodyText }}>{i.label}</span>
          {i.value && <span style={{ fontVariantNumeric: 'tabular-nums' }}>{i.value}</span>}
        </span>
      ))}
    </div>
  );
};

/* ---------- Chart card with PNG export ---------- */

export interface IChartCardProps {
  title: string;
  subtitle?: string;
  span?: number;
  children?: React.ReactNode;
}

export const ChartCard: React.FC<IChartCardProps> = (props) => {
  const { theme, notify } = useApp();
  const ref = React.useRef<HTMLDivElement>(null);
  const exportPng = (): void => {
    const svg = ref.current ? (ref.current.querySelector('svg[data-chart]') as SVGSVGElement | null) : null;
    if (!svg) return;
    downloadSvgAsPng(svg, props.title, theme.semanticColors.bodyBackground).catch(() => notify('The chart image could not be created.', 'error'));
  };
  return (
    <Card
      title={props.title}
      subtitle={props.subtitle}
      span={props.span}
      actions={<IconButton iconProps={{ iconName: 'Photo2' }} title="Download chart as PNG" ariaLabel={`Download ${props.title} as PNG`} onClick={exportPng} />}
    >
      <div ref={ref} data-chart-card={props.title}>{props.children}</div>
    </Card>
  );
};

/* ---------- Column chart (grouped or stacked) ---------- */

export interface ISeries {
  name: string;
  values: number[];
}

export interface IColumnChartProps {
  title: string;
  categories: string[];
  /** Full labels for tooltips (defaults to categories). */
  categoryTitles?: string[];
  series: ISeries[];
  stacked?: boolean;
  height?: number;
  format?: (n: number) => string;
}

export const ColumnChart: React.FC<IColumnChartProps> = (props) => {
  const { theme, isDark } = useApp();
  const ink = chartInk(theme);
  const [ref, width] = useWidth<HTMLDivElement>();
  const [hover, setHover] = React.useState<number | undefined>(undefined);
  const height = props.height || 220;
  const fmt = props.format || fmtNumber;
  const n = props.categories.length;
  const padL = 40;
  const padR = 8;
  const padT = 10;
  const padB = 26;
  const plotW = Math.max(10, width - padL - padR);
  const plotH = height - padT - padB;
  const totals = props.categories.map((_, i) => props.series.reduce((s, se) => s + (se.values[i] || 0), 0));
  const max = props.stacked ? Math.max(0, ...totals) : Math.max(0, ...props.series.map((s) => Math.max(0, ...s.values)));
  const ticks = niceTicks(max);
  const top = ticks[ticks.length - 1] || 1;
  const y = (v: number): number => padT + plotH - (v / top) * plotH;
  const band = n ? plotW / n : plotW;
  const groupW = Math.max(2, band * 0.72);
  const barW = props.stacked ? groupW : Math.max(1, (groupW - (props.series.length - 1) * 2) / Math.max(1, props.series.length));
  const labelEvery = Math.max(1, Math.ceil(n / Math.max(1, Math.floor(plotW / 56))));
  const colors = props.series.map((_, i) => seriesColor(i, isDark));
  const titles = props.categoryTitles || props.categories;

  const tip: ITooltip | undefined =
    hover !== undefined && width
      ? {
          x: padL + band * hover + band / 2,
          y: padT,
          title: titles[hover],
          rows: props.series
            .map((s, i) => ({ label: s.name, value: fmt(s.values[hover] || 0), color: colors[i] }))
            .concat(props.stacked && props.series.length > 1 ? [{ label: 'Total', value: fmt(totals[hover]), color: undefined as unknown as string }] : [])
        }
      : undefined;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
      {props.series.length > 1 && <Legend items={props.series.map((s, i) => ({ label: s.name, color: colors[i] }))} />}
      <div ref={ref} style={{ position: 'relative', width: '100%', height }} onMouseLeave={() => setHover(undefined)}>
        {width > 0 && (
          <svg data-chart={props.title} width={width} height={height} role="img" aria-label={props.title} fontFamily={FONT}>
            {ticks.map((t) => (
              <g key={t}>
                <line x1={padL} x2={width - padR} y1={y(t)} y2={y(t)} stroke={ink.grid} strokeWidth={1} />
                <text x={padL - 6} y={y(t) + 4} textAnchor="end" fontSize={11} fill={ink.muted}>{fmtCompact(t)}</text>
              </g>
            ))}
            {props.categories.map((c, i) => {
              const gx = padL + band * i + (band - groupW) / 2;
              let acc = 0;
              return (
                <g key={i} opacity={hover === undefined || hover === i ? 1 : 0.55}>
                  {props.series.map((s, si) => {
                    const v = s.values[i] || 0;
                    if (v <= 0) return null;
                    let x: number;
                    let y0: number;
                    let h: number;
                    if (props.stacked) {
                      x = gx;
                      y0 = y(acc + v);
                      h = y(acc) - y(acc + v);
                      acc += v;
                      // 2px surface gap between stacked segments.
                      if (si > 0) h = Math.max(0, h - 2);
                    } else {
                      x = gx + si * (barW + 2);
                      y0 = y(v);
                      h = y(0) - y(v);
                    }
                    return <rect key={si} x={x} y={y0} width={barW} height={Math.max(0, h)} rx={Math.min(3, barW / 2)} fill={colors[si]} />;
                  })}
                  {i % labelEvery === 0 && (
                    <text x={padL + band * i + band / 2} y={height - 8} textAnchor="middle" fontSize={11} fill={ink.muted}>{c}</text>
                  )}
                </g>
              );
            })}
            <line x1={padL} x2={width - padR} y1={y(0)} y2={y(0)} stroke={ink.muted} strokeWidth={1} />
            {props.categories.map((_, i) => (
              <rect
                key={`hit${i}`}
                x={padL + band * i}
                y={padT}
                width={band}
                height={plotH}
                fill="transparent"
                onMouseEnter={() => setHover(i)}
                onFocus={() => setHover(i)}
              />
            ))}
          </svg>
        )}
        <Tooltip tip={tip} width={width} />
      </div>
    </div>
  );
};

/* ---------- Line chart with crosshair ---------- */

export interface ILineChartProps {
  title: string;
  dates: string[];
  series: ISeries[];
  height?: number;
  format?: (n: number) => string;
  /** Formats axis ticks; defaults to compact numbers. */
  axisFormat?: (n: number) => string;
  area?: boolean;
}

function shortDate(iso: string): string {
  const d = new Date(Date.parse(iso.length === 10 ? iso + 'T00:00:00' : iso));
  const m = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][d.getMonth()];
  return `${d.getDate()} ${m}`;
}

export const LineChart: React.FC<ILineChartProps> = (props) => {
  const { theme, isDark } = useApp();
  const ink = chartInk(theme);
  const [ref, width] = useWidth<HTMLDivElement>();
  const [hover, setHover] = React.useState<number | undefined>(undefined);
  const height = props.height || 220;
  const fmt = props.format || fmtNumber;
  const axisFmt = props.axisFormat || fmtCompact;
  const n = props.dates.length;
  const padL = 48;
  const padR = 12;
  const padT = 10;
  const padB = 26;
  const plotW = Math.max(10, width - padL - padR);
  const plotH = height - padT - padB;
  const max = Math.max(0, ...props.series.map((s) => Math.max(0, ...s.values)));
  const ticks = niceTicks(max);
  const top = ticks[ticks.length - 1] || 1;
  const x = (i: number): number => padL + (n <= 1 ? plotW / 2 : (i / (n - 1)) * plotW);
  const y = (v: number): number => padT + plotH - (v / top) * plotH;
  const colors = props.series.map((_, i) => seriesColor(i, isDark));
  const labelCount = Math.max(2, Math.floor(plotW / 80));
  const labelIdx: number[] = [];
  for (let k = 0; k < labelCount; k++) labelIdx.push(Math.round((k / (labelCount - 1)) * (n - 1)));

  const onMove = (e: React.MouseEvent<SVGRectElement>): void => {
    const rect = (e.currentTarget as SVGRectElement).getBoundingClientRect();
    const rel = (e.clientX - rect.left) / Math.max(1, rect.width);
    setHover(Math.max(0, Math.min(n - 1, Math.round(rel * (n - 1)))));
  };

  const tip: ITooltip | undefined =
    hover !== undefined && width
      ? { x: x(hover), y: padT, title: shortDate(props.dates[hover]), rows: props.series.map((s, i) => ({ label: s.name, value: fmt(s.values[hover] || 0), color: colors[i] })) }
      : undefined;

  if (!n) return <Empty text="No data for this period." />;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
      {props.series.length > 1 && <Legend items={props.series.map((s, i) => ({ label: s.name, color: colors[i] }))} />}
      <div ref={ref} style={{ position: 'relative', width: '100%', height }} onMouseLeave={() => setHover(undefined)}>
        {width > 0 && (
          <svg data-chart={props.title} width={width} height={height} role="img" aria-label={props.title} fontFamily={FONT}>
            {ticks.map((t) => (
              <g key={t}>
                <line x1={padL} x2={width - padR} y1={y(t)} y2={y(t)} stroke={ink.grid} strokeWidth={1} />
                <text x={padL - 6} y={y(t) + 4} textAnchor="end" fontSize={11} fill={ink.muted}>{axisFmt(t)}</text>
              </g>
            ))}
            {labelIdx.map((i) => (
              <text key={`l${i}`} x={x(i)} y={height - 8} textAnchor={i === 0 ? 'start' : i === n - 1 ? 'end' : 'middle'} fontSize={11} fill={ink.muted}>
                {shortDate(props.dates[i])}
              </text>
            ))}
            {props.series.map((s, si) => {
              const pts = s.values.map((v, i) => `${x(i).toFixed(1)},${y(v || 0).toFixed(1)}`);
              return (
                <g key={s.name}>
                  {props.area && props.series.length === 1 && (
                    <path d={`M${x(0)},${y(0)} L${pts.join(' L')} L${x(n - 1)},${y(0)} Z`} fill={colors[si]} fillOpacity={0.12} stroke="none" />
                  )}
                  <polyline points={pts.join(' ')} fill="none" stroke={colors[si]} strokeWidth={2} strokeLinejoin="round" strokeLinecap="round" />
                  <circle cx={x(n - 1)} cy={y(s.values[n - 1] || 0)} r={3.5} fill={colors[si]} stroke={ink.surface} strokeWidth={2} />
                </g>
              );
            })}
            {hover !== undefined && (
              <g>
                <line x1={x(hover)} x2={x(hover)} y1={padT} y2={padT + plotH} stroke={ink.muted} strokeDasharray="3 3" />
                {props.series.map((s, si) => (
                  <circle key={s.name} cx={x(hover)} cy={y(s.values[hover] || 0)} r={4} fill={colors[si]} stroke={ink.surface} strokeWidth={2} />
                ))}
              </g>
            )}
            <rect x={padL} y={padT} width={plotW} height={plotH} fill="transparent" onMouseMove={onMove} />
          </svg>
        )}
        <Tooltip tip={tip} width={width} />
      </div>
    </div>
  );
};

/* ---------- Donut ---------- */

export interface IDonutProps {
  title: string;
  segments: { label: string; value: number }[];
  format?: (n: number) => string;
  centerLabel?: string;
}

export const Donut: React.FC<IDonutProps> = (props) => {
  const { theme, isDark } = useApp();
  const ink = chartInk(theme);
  const [hover, setHover] = React.useState<number | undefined>(undefined);
  const fmt = props.format || fmtNumber;
  const total = props.segments.reduce((s, x) => s + x.value, 0);
  const size = 160;
  const r = 70;
  const inner = 46;
  const cx = size / 2;
  const cy = size / 2;
  const colorFor = (i: number, label: string): string => (label === 'Other' ? otherColor(theme) : seriesColor(i, isDark));
  let angle = -Math.PI / 2;
  const arcs = props.segments.map((s, i) => {
    const sweep = total ? (s.value / total) * Math.PI * 2 : 0;
    const a0 = angle;
    const a1 = angle + sweep;
    angle = a1;
    const large = sweep > Math.PI ? 1 : 0;
    const p = (a: number, rad: number): string => `${(cx + rad * Math.cos(a)).toFixed(2)},${(cy + rad * Math.sin(a)).toFixed(2)}`;
    const d =
      sweep >= Math.PI * 2 - 0.0001
        ? `M${cx - r},${cy} A${r},${r} 0 1 1 ${cx + r},${cy} A${r},${r} 0 1 1 ${cx - r},${cy} M${cx - inner},${cy} A${inner},${inner} 0 1 0 ${cx + inner},${cy} A${inner},${inner} 0 1 0 ${cx - inner},${cy} Z`
        : `M${p(a0, r)} A${r},${r} 0 ${large} 1 ${p(a1, r)} L${p(a1, inner)} A${inner},${inner} 0 ${large} 0 ${p(a0, inner)} Z`;
    return { d, color: colorFor(i, s.label), label: s.label, value: s.value, sweep };
  });
  if (!total) return <Empty text="No data." />;
  const focus = hover !== undefined ? arcs[hover] : undefined;
  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 16 }}>
      <svg data-chart={props.title} width={size} height={size} role="img" aria-label={props.title} fontFamily={FONT} style={{ flex: 'none' }}>
        {arcs.map((a, i) =>
          a.sweep > 0 ? (
            <path
              key={a.label}
              d={a.d}
              fill={a.color}
              fillRule="evenodd"
              stroke={ink.surface}
              strokeWidth={2}
              opacity={hover === undefined || hover === i ? 1 : 0.45}
              onMouseEnter={() => setHover(i)}
              onMouseLeave={() => setHover(undefined)}
            />
          ) : null
        )}
        <text x={cx} y={cy - 2} textAnchor="middle" fontSize={20} fontWeight={600} fill={ink.text}>{fmtCompact(focus ? focus.value : total)}</text>
        <text x={cx} y={cy + 16} textAnchor="middle" fontSize={11} fill={ink.muted}>{focus ? focus.label.slice(0, 18) : props.centerLabel || 'Total'}</text>
      </svg>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 4, minWidth: 140, flex: 1 }}>
        {arcs.map((a, i) => (
          <div
            key={a.label}
            onMouseEnter={() => setHover(i)}
            onMouseLeave={() => setHover(undefined)}
            style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 12, opacity: hover === undefined || hover === i ? 1 : 0.6 }}
          >
            <span style={{ width: 10, height: 10, borderRadius: 2, background: a.color, flex: 'none' }} />
            <span style={{ flex: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{a.label}</span>
            <span style={{ fontVariantNumeric: 'tabular-nums', fontWeight: 600 }}>{fmt(a.value)}</span>
            <span style={{ fontVariantNumeric: 'tabular-nums', color: theme.semanticColors.bodySubtext, width: 40, textAlign: 'right' }}>
              {Math.round((a.value / total) * 100)}%
            </span>
          </div>
        ))}
      </div>
    </div>
  );
};

/* ---------- Horizontal bar list ---------- */

export interface IBarRow {
  label: string;
  value: number;
  hint?: string;
  onClick?: () => void;
}

export const BarList: React.FC<{ title: string; rows: IBarRow[]; format?: (n: number) => string; colorIndex?: number; max?: number; emptyText?: string }> = (props) => {
  const { theme, isDark } = useApp();
  const fmt = props.format || fmtNumber;
  const max = props.max || Math.max(1, ...props.rows.map((r) => r.value));
  const color = seriesColor(props.colorIndex || 0, isDark);
  if (!props.rows.length) return <Empty text={props.emptyText || 'No data.'} />;
  return (
    <div role="list" aria-label={props.title} style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
      {props.rows.map((r) => (
        <div
          role="listitem"
          key={r.label}
          title={`${r.label}: ${fmt(r.value)}${r.hint ? ` (${r.hint})` : ''}`}
          onClick={r.onClick}
          style={{ display: 'grid', gridTemplateColumns: 'minmax(90px, 38%) 1fr auto', alignItems: 'center', gap: 10, cursor: r.onClick ? 'pointer' : 'default' }}
        >
          <Text variant="small" style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{r.label}</Text>
          <div style={{ height: 10, background: theme.palette.neutralLighter, borderRadius: 3, overflow: 'hidden' }}>
            <div style={{ width: `${Math.max(1, (r.value / max) * 100)}%`, height: '100%', background: r.label === 'Other' ? otherColor(theme) : color, borderRadius: 3 }} />
          </div>
          <Text variant="small" style={{ fontVariantNumeric: 'tabular-nums', fontWeight: 600, minWidth: 48, textAlign: 'right' }}>
            {fmt(r.value)}
          </Text>
        </div>
      ))}
    </div>
  );
};

/* ---------- Meter (used vs available) ---------- */

export const Meter: React.FC<{ used: number; total: number; label?: string }> = ({ used, total, label }) => {
  const { theme, isDark } = useApp();
  const pct = total > 0 ? Math.min(1, used / total) : 0;
  const color = pct >= 0.98 ? (isDark ? '#f07b7b' : '#c50f1f') : pct >= 0.9 ? (isDark ? '#e0a000' : '#a35c00') : seriesColor(0, isDark);
  return (
    <div title={label} style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 120 }}>
      <div style={{ flex: 1, height: 8, background: theme.palette.neutralLighter, borderRadius: 3, overflow: 'hidden' }}>
        <div style={{ width: `${pct * 100}%`, height: '100%', background: color, borderRadius: 3 }} />
      </div>
      <span style={{ fontSize: 12, fontVariantNumeric: 'tabular-nums', width: 38, textAlign: 'right' }}>{Math.round(pct * 100)}%</span>
    </div>
  );
};
