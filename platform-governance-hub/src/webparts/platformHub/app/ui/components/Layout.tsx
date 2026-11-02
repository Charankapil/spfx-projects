import * as React from 'react';
import { DefaultButton, Icon, MessageBar, MessageBarType, Spinner, SpinnerSize, Text, mergeStyles } from '@fluentui/react';
import { useApp } from '../AppContext';
import { describeError } from '../../http/ApiClient';
import { statusColors } from '../palette';

/* ---------- Card ---------- */

export interface ICardProps {
  title?: string;
  subtitle?: string;
  actions?: React.ReactNode;
  /** Grid column span on the 12-column page grid. */
  span?: number;
  children?: React.ReactNode;
  pad?: boolean;
}

export const Card: React.FC<ICardProps> = (props) => {
  const { theme, compact } = useApp();
  const span = compact ? 12 : props.span || 12;
  const cls = mergeStyles({
    gridColumn: `span ${span}`,
    minWidth: 0,
    background: theme.semanticColors.bodyBackground,
    border: `1px solid ${theme.palette.neutralLight}`,
    borderRadius: 6,
    padding: props.pad === false ? 0 : '14px 16px',
    display: 'flex',
    flexDirection: 'column',
    gap: 10
  });
  return (
    <section className={cls}>
      {(props.title || props.actions) && (
        <div style={{ display: 'flex', alignItems: 'flex-start', gap: 8, padding: props.pad === false ? '14px 16px 0' : 0 }}>
          <div style={{ flex: 1, minWidth: 0 }}>
            {props.title && <Text variant="mediumPlus" style={{ fontWeight: 600, display: 'block' }}>{props.title}</Text>}
            {props.subtitle && <Text variant="small" style={{ color: theme.semanticColors.bodySubtext, display: 'block' }}>{props.subtitle}</Text>}
          </div>
          {props.actions}
        </div>
      )}
      {props.children}
    </section>
  );
};

export const PageGrid: React.FC<{ children?: React.ReactNode }> = ({ children }) => (
  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(12, minmax(0, 1fr))', gap: 12, alignItems: 'start' }}>{children}</div>
);

/* ---------- Page header ---------- */

export interface IPageHeaderProps {
  title: string;
  description?: string;
  actions?: React.ReactNode;
}

export const PageHeader: React.FC<IPageHeaderProps> = (props) => {
  const { theme } = useApp();
  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'flex-end', gap: 12, marginBottom: 12 }}>
      <div style={{ flex: '1 1 320px', minWidth: 0 }}>
        <Text as="h2" variant="xLarge" style={{ fontWeight: 600, margin: 0, display: 'block' }}>{props.title}</Text>
        {props.description && (
          <Text variant="small" style={{ color: theme.semanticColors.bodySubtext, display: 'block', maxWidth: 760 }}>{props.description}</Text>
        )}
      </div>
      {props.actions && <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>{props.actions}</div>}
    </div>
  );
};

/* ---------- KPI tiles ---------- */

export type Tone = 'neutral' | 'good' | 'warning' | 'critical';

export interface IKpi {
  label: string;
  value: string;
  sub?: string;
  tone?: Tone;
  onClick?: () => void;
}

export const KpiRow: React.FC<{ items: IKpi[] }> = ({ items }) => {
  const { theme, isDark } = useApp();
  const status = statusColors(isDark);
  const toneColor = (t: Tone | undefined): string | undefined =>
    t === 'critical' ? status.critical : t === 'warning' ? status.warning : t === 'good' ? status.good : undefined;
  return (
    <div style={{ gridColumn: 'span 12', display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: 12 }}>
      {items.map((k) => {
        const color = toneColor(k.tone);
        const cls = mergeStyles({
          background: theme.semanticColors.bodyBackground,
          border: `1px solid ${theme.palette.neutralLight}`,
          borderRadius: 6,
          padding: '12px 14px',
          display: 'flex',
          flexDirection: 'column',
          gap: 2,
          minWidth: 0,
          textAlign: 'left',
          font: 'inherit',
          color: 'inherit',
          cursor: k.onClick ? 'pointer' : 'default',
          selectors: k.onClick ? { ':hover': { borderColor: theme.palette.themePrimary }, ':focus-visible': { outline: `2px solid ${theme.palette.themePrimary}` } } : undefined
        });
        const body = (
          <>
            <Text variant="small" style={{ color: theme.semanticColors.bodySubtext, textTransform: 'uppercase', letterSpacing: '0.04em', fontWeight: 600 }}>{k.label}</Text>
            <Text style={{ fontSize: 26, fontWeight: 600, lineHeight: '32px', fontVariantNumeric: 'tabular-nums', color: theme.semanticColors.bodyText }}>{k.value}</Text>
            {k.sub && (
              <Text variant="small" style={{ color: color || theme.semanticColors.bodySubtext, display: 'flex', alignItems: 'center', gap: 4 }}>
                {k.tone && k.tone !== 'neutral' && <Icon iconName={k.tone === 'good' ? 'CompletedSolid' : 'WarningSolid'} />}
                {k.sub}
              </Text>
            )}
          </>
        );
        return k.onClick ? (
          <button key={k.label} type="button" className={cls} onClick={k.onClick}>{body}</button>
        ) : (
          <div key={k.label} className={cls}>{body}</div>
        );
      })}
    </div>
  );
};

/* ---------- Status pill (never colour alone: icon + text) ---------- */

export const Pill: React.FC<{ tone: Tone; text: string }> = ({ tone, text }) => {
  const { theme, isDark } = useApp();
  const status = statusColors(isDark);
  const color = tone === 'critical' ? status.critical : tone === 'warning' ? status.warning : tone === 'good' ? status.good : theme.semanticColors.bodySubtext;
  const icon = tone === 'critical' ? 'StatusErrorFull' : tone === 'warning' ? 'WarningSolid' : tone === 'good' ? 'CompletedSolid' : 'CircleRing';
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4, color, fontWeight: 600, fontSize: 12, whiteSpace: 'nowrap' }}>
      <Icon iconName={icon} style={{ fontSize: 11 }} />
      <span>{text}</span>
    </span>
  );
};

export function severityTone(s: string): Tone {
  return s === 'High' ? 'critical' : s === 'Medium' ? 'warning' : 'neutral';
}

/* ---------- Loading / error gate ---------- */

export interface IGateProps {
  loading: boolean;
  error: unknown;
  onRetry?: () => void;
  label?: string;
  children?: React.ReactNode;
}

export const Gate: React.FC<IGateProps> = (props) => {
  const { provider } = useApp();
  if (props.error) {
    return (
      <MessageBar
        messageBarType={MessageBarType.error}
        isMultiline
        actions={props.onRetry ? <DefaultButton text="Try again" onClick={props.onRetry} /> : undefined}
      >
        <strong>{props.label || 'This data'} could not be loaded.</strong> {describeError(props.error)}
        {provider.mode === 'live' && ' Open "Data sources" in the toolbar to see which permission or role is missing.'}
      </MessageBar>
    );
  }
  if (props.loading) {
    const progress = provider.getProgress();
    return (
      <div style={{ padding: 32, display: 'flex', justifyContent: 'center' }}>
        <Spinner
          size={SpinnerSize.large}
          label={progress ? `${progress.label} (${progress.done} of ${progress.total})` : `Loading ${props.label ? props.label.toLowerCase() : 'data'}…`}
        />
      </div>
    );
  }
  return <>{props.children}</>;
};

export const Notice: React.FC<{ kind?: 'info' | 'warning'; children?: React.ReactNode }> = ({ kind, children }) => (
  <div style={{ gridColumn: 'span 12' }}>
    <MessageBar messageBarType={kind === 'warning' ? MessageBarType.warning : MessageBarType.info} isMultiline>
      {children}
    </MessageBar>
  </div>
);

export const Empty: React.FC<{ text: string }> = ({ text }) => {
  const { theme } = useApp();
  return <Text variant="small" style={{ color: theme.semanticColors.bodySubtext, padding: '12px 0', display: 'block' }}>{text}</Text>;
};
