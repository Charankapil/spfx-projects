import * as React from 'react';

import styles from './StoragePulse.module.scss';

let instance = 0;

interface IPulseMarkProps {
  size: number;
  /** Spins the ring and runs the pulse line while a scan is running. */
  busy?: boolean;
  className?: string;
}

/** The Storage Pulse logo as inline SVG, so it can animate and needs no extra request. */
export const PulseMark: React.FC<IPulseMarkProps> = ({ size, busy, className }) => {
  // Gradient ids must be unique when two web parts sit on one page.
  const id = React.useMemo(() => `sp${++instance}`, []);
  return (
    <svg
      viewBox="0 0 128 128"
      width={size}
      height={size}
      aria-hidden="true"
      focusable="false"
      className={`${styles.mark} ${busy ? styles.markBusy : ''} ${className || ''}`}
    >
      <defs>
        <linearGradient id={`${id}bg`} x1="0" y1="0" x2="1" y2="1">
          <stop offset="0" stopColor="#0b1530" />
          <stop offset="1" stopColor="#18336a" />
        </linearGradient>
        <linearGradient id={`${id}amber`} x1="0" y1="0" x2="1" y2="1">
          <stop offset="0" stopColor="#ffc062" />
          <stop offset="1" stopColor="#ff8a3d" />
        </linearGradient>
        <linearGradient id={`${id}cyan`} x1="0" y1="1" x2="0" y2="0">
          <stop offset="0" stopColor="#2aa7ff" />
          <stop offset="1" stopColor="#4af0ff" />
        </linearGradient>
      </defs>
      <rect x="4" y="4" width="120" height="120" rx="28" fill={`url(#${id}bg)`} />
      <rect x="4.5" y="4.5" width="119" height="119" rx="27.5" fill="none" stroke="#ffffff" strokeOpacity="0.16" />
      <circle cx="64" cy="64" r="38" fill="none" stroke="#ffffff" strokeOpacity="0.09" strokeWidth="10" />
      <g className={styles.markRing}>
        <circle
          cx="64"
          cy="64"
          r="38"
          fill="none"
          stroke={`url(#${id}amber)`}
          strokeWidth="10"
          strokeLinecap="round"
          strokeDasharray="140 239"
          transform="rotate(-90 64 64)"
        />
        <circle
          cx="64"
          cy="64"
          r="38"
          fill="none"
          stroke={`url(#${id}cyan)`}
          strokeWidth="10"
          strokeLinecap="round"
          strokeDasharray="76 239"
          strokeDashoffset="-153"
          transform="rotate(-90 64 64)"
        />
      </g>
      <path
        className={styles.markLine}
        d="M35 66h14l5-12 8 25 7-31 6 23 4-5h14"
        fill="none"
        stroke="#eef8ff"
        strokeWidth="5.5"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
};
