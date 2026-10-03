import * as React from 'react';
import { DefaultButton, Icon, MessageBar, MessageBarType } from '@fluentui/react';
import styles from '../AdminCenter.module.scss';

export const PALETTE = ['#2563eb', '#e07b00', '#0d9488', '#9333ea', '#d6336c', '#65a30d', '#b45309', '#64748b'];

export const Loading: React.FC<{ text?: string }> = ({ text }) => (
  <div className={styles.spinnerWrap} role="status" aria-live="polite">
    <div className={styles.spinner} />
    <span>{text || 'Loading…'}</span>
  </div>
);

export const ErrorBar: React.FC<{ error: string; onRetry?: () => void }> = ({ error, onRetry }) => (
  <MessageBar
    messageBarType={MessageBarType.error}
    isMultiline
    actions={onRetry ? <DefaultButton onClick={onRetry}>Try again</DefaultButton> : undefined}
  >
    {error}
  </MessageBar>
);

export const Empty: React.FC<{ text: string }> = ({ text }) => <div className={styles.empty}>{text}</div>;

export const ViewHeader: React.FC<{ title: string; hint?: string; children?: React.ReactNode }> = ({ title, hint, children }) => (
  <div className={styles.viewHead}>
    <div>
      <h2 className={styles.viewTitle}>{title}</h2>
      {hint && <p className={styles.viewHint}>{hint}</p>}
    </div>
    <div className={styles.actions}>{children}</div>
  </div>
);

export const Card: React.FC<{ title?: string; right?: React.ReactNode; children: React.ReactNode }> = ({ title, right, children }) => (
  <section className={styles.card}>
    {title && (
      <h3 className={styles.cardTitle}>
        <span>{title}</span>
        {right}
      </h3>
    )}
    {children}
  </section>
);

export const Kpi: React.FC<{ label: string; value: string; sub?: string; onClick?: () => void }> = ({ label, value, sub, onClick }) => {
  const inner = (
    <>
      <div className={styles.kpiLabel}>{label}</div>
      <div className={styles.kpiValue}>{value}</div>
      {sub && <div className={styles.kpiSub}>{sub}</div>}
    </>
  );
  return onClick ? (
    <button type="button" className={`${styles.kpi} ${styles.kpiClickable}`} onClick={onClick}>
      {inner}
    </button>
  ) : (
    <div className={styles.kpi}>{inner}</div>
  );
};

export const Pill: React.FC<{ kind?: 'critical' | 'warning' | 'good' | 'info'; children: React.ReactNode }> = ({ kind, children }) => {
  const cls = kind === 'critical' ? styles.pillCritical : kind === 'warning' ? styles.pillWarning : kind === 'good' ? styles.pillGood : kind === 'info' ? styles.pillInfo : '';
  return <span className={`${styles.pill} ${cls}`}>{children}</span>;
};

export interface ISlice {
  label: string;
  value: number;
  color?: string;
}

/** Donut chart in plain SVG with a legend; accessible via the legend text. */
export const Donut: React.FC<{ slices: ISlice[]; centre?: string; centreSub?: string; size?: number; format?: (n: number) => string }> = ({ slices, centre, centreSub, size = 140, format }) => {
  const total = slices.reduce((s, x) => s + x.value, 0);
  const r = size / 2 - 12;
  const c = 2 * Math.PI * r;
  let offset = 0;
  return (
    <div className={styles.donutWrap}>
      <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} role="img" aria-label={slices.map((s) => `${s.label} ${s.value}`).join(', ')}>
        <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke="rgba(0,0,0,0.06)" strokeWidth={16} />
        {total > 0 &&
          slices.map((s, i) => {
            const len = (s.value / total) * c;
            const el = (
              <circle
                key={s.label}
                cx={size / 2}
                cy={size / 2}
                r={r}
                fill="none"
                stroke={s.color || PALETTE[i % PALETTE.length]}
                strokeWidth={16}
                strokeDasharray={`${Math.max(0, len - 1)} ${c - Math.max(0, len - 1)}`}
                strokeDashoffset={-offset}
                transform={`rotate(-90 ${size / 2} ${size / 2})`}
              />
            );
            offset += len;
            return el;
          })}
        {centre && (
          <text x="50%" y={centreSub ? '47%' : '52%'} textAnchor="middle" fontSize="22" fontWeight="600" fill="#242424">
            {centre}
          </text>
        )}
        {centreSub && (
          <text x="50%" y="62%" textAnchor="middle" fontSize="11" fill="#616161">
            {centreSub}
          </text>
        )}
      </svg>
      <div className={styles.legend}>
        {slices.map((s, i) => (
          <div key={s.label} className={styles.legendItem}>
            <span className={styles.swatch} style={{ background: s.color || PALETTE[i % PALETTE.length] }} />
            <span>
              {s.label} <span className={styles.muted}>· {format ? format(s.value) : s.value.toLocaleString()}{total > 0 ? ` (${Math.round((s.value / total) * 100)}%)` : ''}</span>
            </span>
          </div>
        ))}
      </div>
    </div>
  );
};

export const BarList: React.FC<{ rows: Array<{ label: string; value: number; display?: string }>; color?: string }> = ({ rows, color }) => {
  const max = rows.reduce((m, r) => Math.max(m, r.value), 0) || 1;
  return (
    <div>
      {rows.map((r) => (
        <div className={styles.barRow} key={r.label}>
          <span className={styles.barLabel} title={r.label}>
            {r.label}
          </span>
          <div className={styles.barTrack}>
            <div className={styles.barFill} style={{ width: `${Math.max(2, (r.value / max) * 100)}%`, background: color }} />
          </div>
          <span className={styles.barValue}>{r.display || r.value.toLocaleString()}</span>
        </div>
      ))}
    </div>
  );
};

export const ScoreRing: React.FC<{ score: number }> = ({ score }) => {
  const size = 132;
  const r = 54;
  const c = 2 * Math.PI * r;
  const color = score >= 85 ? '#107c41' : score >= 60 ? '#b87400' : '#c4314b';
  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`} role="img" aria-label={`Health score ${score} out of 100`}>
      <circle cx={size / 2} cy={size / 2} r={r} fill="none" stroke="rgba(0,0,0,0.07)" strokeWidth={12} />
      <circle
        cx={size / 2}
        cy={size / 2}
        r={r}
        fill="none"
        stroke={color}
        strokeWidth={12}
        strokeLinecap="round"
        strokeDasharray={`${(score / 100) * c} ${c}`}
        transform={`rotate(-90 ${size / 2} ${size / 2})`}
      />
      <text x="50%" y="52%" textAnchor="middle" fontSize="32" fontWeight="600" fill="#242424">
        {score}
      </text>
      <text x="50%" y="68%" textAnchor="middle" fontSize="11" fill="#616161">
        out of 100
      </text>
    </svg>
  );
};

export const severityIcon = (s: string): { icon: string; color: string } =>
  s === 'critical'
    ? { icon: 'ErrorBadge', color: '#c4314b' }
    : s === 'warning'
    ? { icon: 'Warning', color: '#b87400' }
    : s === 'good'
    ? { icon: 'CompletedSolid', color: '#107c41' }
    : { icon: 'Info', color: '#0b6a9e' };

export const FindingRow: React.FC<{ severity: string; title: string; detail?: string; onGo?: () => void }> = ({ severity, title, detail, onGo }) => {
  const si = severityIcon(severity);
  return (
    <div className={styles.finding}>
      <Icon iconName={si.icon} className={styles.findingIcon} style={{ color: si.color }} aria-hidden="true" />
      <div>
        <div className={styles.findingTitle}>{title}</div>
        {detail && <div className={styles.findingDetail}>{detail}</div>}
      </div>
      {onGo && (
        <DefaultButton className={styles.findingGo} onClick={onGo}>
          Review
        </DefaultButton>
      )}
    </div>
  );
};


/** Tiny trend line for a history of values. */
export const Sparkline: React.FC<{ values: number[]; color?: string; width?: number; height?: number }> = ({ values, color = '#2563eb', width = 110, height = 28 }) => {
  if (values.length < 2) {
    return <span className={styles.muted}>no trend yet</span>;
  }
  const min = Math.min.apply(null, values);
  const max = Math.max.apply(null, values);
  const range = max - min || 1;
  const pts = values.map((v, i) => `${((i / (values.length - 1)) * (width - 4) + 2).toFixed(1)},${(height - 3 - ((v - min) / range) * (height - 6)).toFixed(1)}`);
  const last = pts[pts.length - 1].split(',');
  return (
    <svg width={width} height={height} viewBox={`0 0 ${width} ${height}`} role="img" aria-label={`Trend over ${values.length} snapshots`}>
      <polyline points={pts.join(' ')} fill="none" stroke={color} strokeWidth={1.6} strokeLinejoin="round" strokeLinecap="round" />
      <circle cx={last[0]} cy={last[1]} r={2.4} fill={color} />
    </svg>
  );
};
