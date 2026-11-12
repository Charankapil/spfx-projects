import * as React from 'react';
import {
  DefaultButton, Dialog, DialogFooter, DialogType, MessageBarType, NormalPeoplePicker, type IPersonaProps, PrimaryButton
} from '@fluentui/react';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import styles from './VantageApp.module.scss';
import { SharePointService } from '../services/SharePointService';
import type { ArchiveService } from '../services/ArchiveService';
import type { PermissionService } from '../services/PermissionService';
import type { ProvisioningService } from '../services/ProvisioningService';
import type { WorkService } from '../services/WorkService';
import type { ICurrentUser, IWorkspaceSummary, WorkspaceRole } from '../models/Models';

// ---------------------------------------------------------------------------
// App context
// ---------------------------------------------------------------------------

export interface IAppContext {
  sp: SharePointService;
  provisioner: ProvisioningService;
  work: WorkService;
  archiver: ArchiveService;
  permissions: PermissionService;
  user: ICurrentUser;
  webServerRelativeUrl: string;
  siteOwner: boolean;
  workspace: IWorkspaceSummary | undefined;
  role: WorkspaceRole;
  /** Show a message bar at the top of the app. */
  notify: (type: MessageBarType, text: string) => void;
  /** Run an action with a busy indicator and error reporting. Resolves false on error. */
  perform: (label: string, fn: () => Promise<void>) => Promise<boolean>;
  busy: boolean;
  /** Ask other tabs to reload workspace data (tasks, counts). */
  refreshToken: number;
  bumpRefresh: () => void;
  /** Theme CSS variables; apply inside Panels/Dialogs, which render outside the app root. */
  themeStyle: React.CSSProperties;
}

export const AppContext = React.createContext<IAppContext>(undefined as unknown as IAppContext);
export const useApp = (): IAppContext => React.useContext(AppContext);

export const canEdit = (role: WorkspaceRole): boolean => role === 'owner' || role === 'admin' || role === 'member';
export const canAdmin = (role: WorkspaceRole): boolean => role === 'owner' || role === 'admin';

// ---------------------------------------------------------------------------
// Formatting
// ---------------------------------------------------------------------------

export const fmt = (n?: number): string => (n === undefined || n === null ? '—' : Number(n).toLocaleString());
export const formatDate = (d?: Date | string): string => {
  if (!d) return '';
  const date = typeof d === 'string' ? new Date(d) : d;
  return isNaN(date.getTime()) ? '' : date.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
};
export const formatDateTime = (iso?: string): string =>
  iso ? new Date(iso).toLocaleString(undefined, { year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' }) : '—';

export const relativeTime = (iso?: string): string => {
  if (!iso) return '';
  const mins = Math.round((Date.now() - new Date(iso).getTime()) / 60000);
  if (mins < 1) return 'just now';
  if (mins < 60) return `${mins} min ago`;
  const hours = Math.round(mins / 60);
  if (hours < 24) return `${hours} h ago`;
  const days = Math.round(hours / 24);
  return days < 30 ? `${days} d ago` : formatDate(iso);
};

/** A calendar date chosen in a picker → UTC midnight ISO string (timezone independent). */
export const toUtcIso = (d: Date): string => new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())).toISOString();
/** A calendar date chosen in a picker → local noon ISO string (stable for due dates in any timezone). */
export const toNoonIso = (d: Date): string => new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12).toISOString();
/** ISO string saved as UTC midnight → local Date representing the same calendar date. */
export const fromUtcIso = (iso?: string): Date | undefined => {
  if (!iso) return undefined;
  const d = new Date(iso);
  return isNaN(d.getTime()) ? undefined : new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
};

/** Accepts ISO (2024-06-01) or locale-formatted input. Fluent UI's DatePicker API requires `null` for "invalid". */
// eslint-disable-next-line @rushstack/no-new-null
export const parseDate = (text: string): Date | null => {
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text.trim());
  const d = iso ? new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])) : new Date(text);
  return isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
};

export const isOverdue = (due?: string, status?: string): boolean => {
  if (!due || status === 'Completed' || status === 'Cancelled') return false;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  return new Date(due) < today;
};

// ---------------------------------------------------------------------------
// Theme
// ---------------------------------------------------------------------------

export function themeVars(theme: IReadonlyTheme | undefined): React.CSSProperties {
  const s = theme?.semanticColors;
  const p = theme?.palette;
  const vars: Record<string, string | undefined> = {
    '--vnt-bg': s?.bodyStandoutBackground,
    '--vnt-card': s?.bodyBackground,
    '--vnt-border': s?.bodyDivider,
    '--vnt-text': s?.bodyText,
    '--vnt-subtle': s?.bodySubtext,
    '--vnt-primary': p?.themePrimary,
    '--vnt-primary-soft': p?.themeLighterAlt,
    '--vnt-on-primary': s?.primaryButtonText,
    '--vnt-success': s?.successIcon,
    '--vnt-warning': s?.warningIcon,
    '--vnt-error': s?.errorText
  };
  Object.keys(vars).forEach(k => { if (!vars[k]) delete vars[k]; });
  return vars as React.CSSProperties;
}

// ---------------------------------------------------------------------------
// Small presentational components
// ---------------------------------------------------------------------------

const STATUS_TONE: Record<string, string> = {
  // Task statuses
  'Not Started': styles.neutral, 'In Progress': styles.info, 'On Hold': styles.warn, 'In Review': styles.purple,
  Completed: styles.ok, Cancelled: styles.neutral,
  // Job statuses
  Running: styles.info, CompletedWithErrors: styles.warn, Failed: styles.bad, Abandoned: styles.neutral,
  // Priorities
  Low: styles.neutral, Medium: styles.info, High: styles.warn, Critical: styles.bad
};
const STATUS_LABEL: Record<string, string> = { CompletedWithErrors: 'Completed with errors', Abandoned: 'Stopped' };

export const Badge: React.FC<{ value?: string; tone?: string }> = ({ value, tone }) =>
  value ? <span className={`${styles.badge} ${tone || STATUS_TONE[value] || styles.neutral}`}>{STATUS_LABEL[value] || value}</span> : <>—</>;

export const Kpi: React.FC<{ label: string; value: React.ReactNode; note?: React.ReactNode; tone?: string }> = ({ label, value, note, tone }) => (
  <div className={styles.card} role="group" aria-label={label}>
    <div className={styles.kpiLabel}>{label}</div>
    <div className={`${styles.kpiValue} ${tone || ''}`}>{value}</div>
    {note && <div className={styles.kpiNote}>{note}</div>}
  </div>
);

export const Card: React.FC<{ title?: React.ReactNode; actions?: React.ReactNode; children?: React.ReactNode; className?: string }> = ({ title, actions, children, className }) => (
  <section className={`${styles.card} ${className || ''}`}>
    {(title || actions) && (
      <div className={styles.cardHeader}>
        {title && <h3 className={styles.cardTitle}>{title}</h3>}
        {actions && <div className={styles.cardActions}>{actions}</div>}
      </div>
    )}
    {children}
  </section>
);

export const EmptyState: React.FC<{ title: string; text?: React.ReactNode; action?: React.ReactNode }> = ({ title, text, action }) => (
  <div className={styles.emptyState}>
    <h3 className={styles.cardTitle}>{title}</h3>
    {text && <p className={styles.muted}>{text}</p>}
    {action}
  </div>
);

export interface IConfirmOptions {
  title: string;
  text: React.ReactNode;
  confirmText: string;
  danger?: boolean;
}

/** Promise-based confirm dialog: `const ok = await confirm({...})`. */
export function useConfirm(): [(o: IConfirmOptions) => Promise<boolean>, React.ReactElement] {
  const [state, setState] = React.useState<(IConfirmOptions & { resolve: (v: boolean) => void }) | undefined>(undefined);
  const confirm = React.useCallback((o: IConfirmOptions) => new Promise<boolean>(resolve => setState({ ...o, resolve })), []);
  const close = (v: boolean): void => { if (state) state.resolve(v); setState(undefined); };
  const dialog = (
    <Dialog hidden={!state} onDismiss={() => close(false)} minWidth={400} modalProps={{ isBlocking: true }}
      dialogContentProps={{ type: DialogType.normal, title: state?.title }}>
      <div className={styles.dialogText}>{state?.text}</div>
      <DialogFooter>
        <PrimaryButton text={state?.confirmText} onClick={() => close(true)} className={state?.danger ? styles.dangerButton : undefined} />
        <DefaultButton text="Cancel" onClick={() => close(false)} />
      </DialogFooter>
    </Dialog>
  );
  return [confirm, dialog];
}

// ---------------------------------------------------------------------------
// People picker (SharePoint people search)
// ---------------------------------------------------------------------------

export interface IPersonValue {
  key: string;
  name: string;
}

export const PeoplePicker: React.FC<{
  label: string; value?: IPersonValue; onChange: (v: IPersonValue | undefined) => void; disabled?: boolean; placeholder?: string;
}> = ({ label, value, onChange, disabled, placeholder }) => {
  const { sp } = useApp();
  const id = React.useMemo(() => `vnt-pp-${Math.random().toString(36).slice(2)}`, []);
  const selected: IPersonaProps[] = value && value.key ? [{ key: value.key, text: value.name, secondaryText: value.key.split('|').pop() }] : [];
  const resolve = async (filter: string): Promise<IPersonaProps[]> => {
    try {
      const people = await sp.searchPeople(filter);
      return people.map(p => ({ key: p.key, text: p.displayName, secondaryText: p.email || p.key.split('|').pop() }));
    } catch {
      return [];
    }
  };
  return (
    <div>
      <label htmlFor={id} className={styles.fieldLabel}>{label}</label>
      <NormalPeoplePicker
        onResolveSuggestions={resolve}
        selectedItems={selected}
        itemLimit={1}
        disabled={disabled}
        resolveDelay={300}
        inputProps={{ id, placeholder: selected.length ? undefined : (placeholder || 'Search people'), 'aria-label': label }}
        pickerSuggestionsProps={{ noResultsFoundText: 'No people found', loadingText: 'Searching…', suggestionsHeaderText: 'People' }}
        onChange={items => {
          const p = items && items[0];
          onChange(p ? { key: String(p.key), name: p.text || String(p.key) } : undefined);
        }}
      />
    </div>
  );
};

export const listUrl = (webServerRelativeUrl: string, title: string): string =>
  `${webServerRelativeUrl.replace(/\/$/, '')}/Lists/${encodeURIComponent(title)}`;

export const errorText = (e: unknown): string => SharePointService.errorMessage(e);
