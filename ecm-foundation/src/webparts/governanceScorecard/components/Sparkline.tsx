import * as React from 'react';

import { INK, SERIES } from './vizTokens';

/** Snapshot history in the de-emphasis grey, the displayed run marked in the accent. */
export const Sparkline = (props: { values: (number | undefined)[]; highlight: number; label: string }): JSX.Element | null => {
  const points = props.values.map((v, i) => ({ v, i })).filter((p) => p.v !== undefined) as { v: number; i: number }[];
  if (points.length < 2) {
    return null;
  }
  const w = 96;
  const h = 28;
  const pad = 4;
  const min = Math.min(...points.map((p) => p.v));
  const max = Math.max(...points.map((p) => p.v));
  const span = max - min || 1;
  const x = (i: number): number => pad + (i / Math.max(1, props.values.length - 1)) * (w - pad * 2);
  const y = (v: number): number => h - pad - ((v - min) / span) * (h - pad * 2);
  const d = points.map((p, k) => `${k ? 'L' : 'M'}${x(p.i).toFixed(1)},${y(p.v).toFixed(1)}`).join(' ');
  const current = points.filter((p) => p.i === props.highlight)[0];
  return (
    <svg width={w} height={h} viewBox={`0 0 ${w} ${h}`} role="img" aria-label={props.label}>
      <path d={d} fill="none" stroke={INK.deEmphasis} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
      {current && <circle cx={x(current.i)} cy={y(current.v)} r="4" fill={SERIES.score} stroke="#fff" strokeWidth="2" />}
    </svg>
  );
};
