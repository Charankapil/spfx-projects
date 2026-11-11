import * as React from 'react';
import {
  ActionButton, DatePicker, DayOfWeek, DefaultButton, DetailsList, DetailsListLayoutMode, Dialog, DialogFooter,
  DialogType, Dropdown, type IColumn, type IDropdownOption, type ITheme, Link, MessageBar, MessageBarType, Pivot,
  PivotItem, PrimaryButton, ProgressIndicator, SelectionMode, Spinner, SpinnerSize, TextField, ThemeProvider, Toggle,
  TooltipHost
} from '@fluentui/react';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import styles from './VantageApp.module.scss';
import { LISTS, ProvisioningService, type ISetupState } from '../services/ProvisioningService';
import { ArchiveService } from '../services/ArchiveService';
import { SharePointService, type SPEntity } from '../services/SharePointService';
import {
  type ArchiveMode, DEFAULT_ELIGIBLE_STATUSES, type IArchiveJob, type IArchivePolicy, type IArchiveProgress,
  type IWorkspaceSummary, MAX_BATCH_SIZE
} from '../models/Models';

export interface IVantageAppProps {
  sp: SharePointService;
  webUrl: string;
  webServerRelativeUrl: string;
  currentUser: string;
  canManageLists: boolean;
  canArchive: boolean;
  theme: IReadonlyTheme | undefined;
  title: string;
  subtitle: string;
}

interface IPolicyForm {
  enabled: boolean;
  mode: ArchiveMode;
  maxItems: string;
  archiveBefore: Date | undefined;
  batchSize: string;
}

interface IStats {
  activeItems?: number;
  archivedItems?: number;
  lastJob?: IArchiveJob;
}

type Notice = { type: MessageBarType; text: string } | undefined;

const MODE_OPTIONS: IDropdownOption[] = [
  { key: 'date', text: 'Archive items completed before a date' },
  { key: 'count', text: 'Keep the active list under a maximum size' },
  { key: 'both', text: 'Both: date first, then size limit' }
];

const DEFAULT_FORM: IPolicyForm = { enabled: false, mode: 'date', maxItems: '100000', archiveBefore: undefined, batchSize: '100' };

const STATUS_STYLE: Record<string, string> = {
  Completed: styles.ok,
  CompletedWithErrors: styles.warn,
  Running: styles.info,
  Failed: styles.bad,
  Abandoned: styles.neutral
};

const STATUS_LABEL: Record<string, string> = {
  Completed: 'Completed',
  CompletedWithErrors: 'Completed with errors',
  Running: 'Running',
  Failed: 'Failed',
  Abandoned: 'Stopped'
};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** A calendar date chosen in the picker → UTC midnight ISO string (timezone independent). */
const toUtcIso = (d: Date): string => new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())).toISOString();
/** ISO string saved as UTC midnight → local Date representing the same calendar date. */
const fromUtcIso = (iso: string): Date | undefined => {
  const d = new Date(iso);
  return isNaN(d.getTime()) ? undefined : new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
};
const formatDate = (d?: Date): string =>
  d ? d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' }) : '';
const formatDateTime = (iso?: string): string =>
  iso ? new Date(iso).toLocaleString(undefined, { year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' }) : '—';
/** Accepts ISO (2024-06-01) or locale-formatted input. Fluent UI's DatePicker API requires `null` for "invalid". */
// eslint-disable-next-line @rushstack/no-new-null
const parseDate = (text: string): Date | null => {
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text.trim());
  const d = iso ? new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])) : new Date(text);
  return isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
};
const fmt = (n?: number): string => (n === undefined || n === null ? '—' : Number(n).toLocaleString());

const formToPolicy = (workspaceId: string, f: IPolicyForm): IArchivePolicy => ({
  WorkspaceId: workspaceId,
  Enabled: f.enabled,
  Mode: f.mode,
  MaxActiveItems: Math.floor(Number(f.maxItems)),
  ArchiveBefore: f.archiveBefore ? toUtcIso(f.archiveBefore) : undefined,
  EligibleStatuses: DEFAULT_ELIGIBLE_STATUSES,
  BatchSize: Math.floor(Number(f.batchSize))
});

const formsEqual = (a: IPolicyForm, b: IPolicyForm): boolean =>
  a.enabled === b.enabled && a.mode === b.mode && a.maxItems === b.maxItems && a.batchSize === b.batchSize &&
  (a.archiveBefore ? a.archiveBefore.getTime() : 0) === (b.archiveBefore ? b.archiveBefore.getTime() : 0);

function themeVars(theme: IReadonlyTheme | undefined): React.CSSProperties {
  const s = theme?.semanticColors;
  const p = theme?.palette;
  const vars: Record<string, string | undefined> = {
    '--vnt-bg': s?.bodyStandoutBackground,
    '--vnt-card': s?.bodyBackground,
    '--vnt-border': s?.bodyDivider,
    '--vnt-text': s?.bodyText,
    '--vnt-subtle': s?.bodySubtext,
    '--vnt-primary': p?.themePrimary,
    '--vnt-on-primary': s?.primaryButtonText,
    '--vnt-success': s?.successIcon,
    '--vnt-warning': s?.warningIcon,
    '--vnt-error': s?.errorText
  };
  Object.keys(vars).forEach(k => { if (!vars[k]) delete vars[k]; });
  return vars as React.CSSProperties;
}

const StatusBadge: React.FC<{ status?: string }> = ({ status }) =>
  status ? <span className={`${styles.badge} ${STATUS_STYLE[status] || styles.neutral}`}>{STATUS_LABEL[status] || status}</span> : <>—</>;

const Kpi: React.FC<{ label: string; value: React.ReactNode; note?: React.ReactNode }> = ({ label, value, note }) => (
  <div className={styles.card} role="group" aria-label={label}>
    <div className={styles.kpiLabel}>{label}</div>
    <div className={styles.kpiValue}>{value}</div>
    {note && <div className={styles.kpiNote}>{note}</div>}
  </div>
);

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export const VantageApp: React.FC<IVantageAppProps> = props => {
  const { sp, webServerRelativeUrl, canManageLists, canArchive, theme } = props;

  const provisioner = React.useMemo(() => new ProvisioningService(sp), [sp]);
  const archiver = React.useMemo(() => new ArchiveService(sp), [sp]);

  const [loading, setLoading] = React.useState(true);
  const [setup, setSetup] = React.useState<ISetupState | undefined>(undefined);
  const [busy, setBusy] = React.useState<string>('');
  const [notice, setNotice] = React.useState<Notice>(undefined);
  const [tab, setTab] = React.useState<string>('overview');

  const [workspaces, setWorkspaces] = React.useState<IWorkspaceSummary[]>([]);
  const [workspace, setWorkspace] = React.useState<IWorkspaceSummary | undefined>(undefined);
  const [newName, setNewName] = React.useState('');

  const [form, setForm] = React.useState<IPolicyForm>(DEFAULT_FORM);
  const [savedForm, setSavedForm] = React.useState<IPolicyForm>(DEFAULT_FORM);
  const [policyId, setPolicyId] = React.useState<number | undefined>(undefined);
  const [policyLoading, setPolicyLoading] = React.useState(false);

  const [stats, setStats] = React.useState<IStats>({});
  const [jobs, setJobs] = React.useState<IArchiveJob[]>([]);
  const [confirmOpen, setConfirmOpen] = React.useState(false);
  const [running, setRunning] = React.useState(false);
  const [progress, setProgress] = React.useState<IArchiveProgress | undefined>(undefined);

  const dirty = !formsEqual(form, savedForm);
  const initialized = !!setup?.initialized;

  // --- Initial load --------------------------------------------------------
  React.useEffect(() => {
    let active = true;
    (async () => {
      try {
        const state = await provisioner.getSetupState();
        if (!active) return;
        setSetup(state);
        if (state.initialized) {
          const list = await provisioner.listWorkspaces();
          if (!active) return;
          setWorkspaces(list);
          setWorkspace(list[0]);
        }
      } catch (e) {
        if (active) setNotice({ type: MessageBarType.error, text: SharePointService.errorMessage(e, 'Could not load VANTAGE configuration.') });
      } finally {
        if (active) setLoading(false);
      }
    })().catch(() => undefined);
    return () => { active = false; };
  }, [provisioner]);

  // --- Workspace change: load policy, stats and job history ----------------
  const refreshWorkspaceData = React.useCallback(async (ws: IWorkspaceSummary): Promise<void> => {
    const [active, archived, recent] = await Promise.all([
      sp.getListItemCount(ws.workItemsList).catch(() => undefined),
      sp.getListItemCount(ws.archiveList).catch(() => undefined),
      archiver.getRecentJobs(ws.id).catch(() => [] as IArchiveJob[])
    ]);
    setStats({ activeItems: active, archivedItems: archived, lastJob: recent[0] });
    setJobs(recent);
  }, [sp, archiver]);

  React.useEffect(() => {
    if (!workspace || !initialized) return;
    let active = true;
    setPolicyLoading(true);
    (async () => {
      try {
        const result = await sp.getCollection<SPEntity>(
          `${sp.listPath(LISTS.policies)}/items?$select=Id,Enabled,ArchiveMode,MaxActiveItems,ArchiveBefore,BatchSize` +
          `&$filter=${encodeURIComponent(`WorkspaceId eq '${SharePointService.odataString(workspace.id)}'`)}&$top=1`
        );
        if (!active) return;
        const saved = result.value[0];
        const loaded: IPolicyForm = saved ? {
          enabled: !!saved.Enabled,
          mode: (['count', 'date', 'both'].indexOf(saved.ArchiveMode as string) >= 0 ? saved.ArchiveMode : 'date') as ArchiveMode,
          maxItems: String(saved.MaxActiveItems || 100000),
          archiveBefore: saved.ArchiveBefore ? fromUtcIso(saved.ArchiveBefore as string) : undefined,
          batchSize: String(saved.BatchSize || 100)
        } : DEFAULT_FORM;
        setPolicyId(saved ? (saved.Id as number) : undefined);
        setForm(loaded);
        setSavedForm(loaded);
        await refreshWorkspaceData(workspace);
      } catch (e) {
        if (active) setNotice({ type: MessageBarType.error, text: SharePointService.errorMessage(e, 'Could not load the archive policy.') });
      } finally {
        if (active) setPolicyLoading(false);
      }
    })().catch(() => undefined);
    return () => { active = false; };
  }, [workspace, initialized, sp, refreshWorkspaceData]);

  // --- Warn before leaving the page while an archive run is in progress -----
  React.useEffect(() => {
    if (!running) return;
    const handler = (e: BeforeUnloadEvent): string => {
      e.preventDefault();
      e.returnValue = 'An archive run is in progress. Leaving this page will stop it.';
      return e.returnValue;
    };
    window.addEventListener('beforeunload', handler);
    return () => window.removeEventListener('beforeunload', handler);
  }, [running]);

  // --- Actions ---------------------------------------------------------------
  const perform = async (label: string, fn: () => Promise<void>): Promise<void> => {
    setBusy(label);
    setNotice(undefined);
    try {
      await fn();
    } catch (e) {
      setNotice({ type: MessageBarType.error, text: SharePointService.errorMessage(e) });
    } finally {
      setBusy('');
    }
  };

  const initialize = (): Promise<void> => perform('Setting up VANTAGE…', async () => {
    await provisioner.initialize(step => setBusy(step));
    const state = await provisioner.getSetupState();
    setSetup(state);
    const list = await provisioner.listWorkspaces();
    setWorkspaces(list);
    if (!workspace && list.length) setWorkspace(list[0]);
    setNotice({ type: MessageBarType.success, text: 'VANTAGE is set up on this site. Create your first workspace to get started.' });
    setTab(list.length ? 'overview' : 'workspaces');
  });

  const createWorkspace = (): Promise<void> => perform('Creating workspace lists…', async () => {
    const created = await provisioner.createWorkspace(newName, workspaces);
    setWorkspaces(current => current.concat([created]).sort((a, b) => a.title.localeCompare(b.title)));
    setWorkspace(created);
    setNewName('');
    setNotice({ type: MessageBarType.success, text: `Workspace “${created.title}” was created with its active and archive lists.` });
  });

  const validationError = ((): string | undefined => {
    if (!workspace) return undefined;
    try {
      ArchiveService.validatePolicy(formToPolicy(workspace.id, form));
      return undefined;
    } catch (e) {
      return SharePointService.errorMessage(e);
    }
  })();

  const savePolicy = (): Promise<void> => perform('Saving policy…', async () => {
    if (!workspace) throw new Error('Select a workspace first.');
    if (validationError) throw new Error(validationError);
    const p = formToPolicy(workspace.id, form);
    const fields: SPEntity = {
      Title: `${workspace.title} archive policy`,
      WorkspaceId: workspace.id,
      Enabled: p.Enabled,
      ArchiveMode: p.Mode,
      MaxActiveItems: p.MaxActiveItems,
      ArchiveBefore: p.ArchiveBefore || null,
      EligibleStatuses: p.EligibleStatuses,
      BatchSize: p.BatchSize
    };
    if (policyId) {
      await sp.updateItem(LISTS.policies, policyId, fields);
    } else {
      const created = await sp.createItem(LISTS.policies, fields);
      setPolicyId(created.Id);
    }
    setSavedForm(form);
    setNotice({ type: MessageBarType.success, text: 'Archive policy saved.' });
  });

  const runArchive = async (): Promise<void> => {
    setConfirmOpen(false);
    if (!workspace) return;
    const ws = workspace;
    setRunning(true);
    setNotice(undefined);
    setProgress({ processed: 0, archived: 0, failed: 0, skipped: 0, message: 'Starting archive run…' });
    try {
      const result = await archiver.run(ws, formToPolicy(ws.id, savedForm), p => setProgress(p));
      setNotice({
        type: result.failed ? MessageBarType.warning : MessageBarType.success,
        text: `Archive run finished: ${fmt(result.archived)} archived, ${fmt(result.skipped)} skipped (attachments), ${fmt(result.failed)} failed. See Job history and the ${LISTS.log} list for detail.`
      });
    } catch (e) {
      setNotice({ type: MessageBarType.error, text: SharePointService.errorMessage(e) });
    } finally {
      setRunning(false);
      await refreshWorkspaceData(ws).catch(() => undefined);
    }
  };

  // --- Derived UI ------------------------------------------------------------
  const listUrl = (title: string): string => `${webServerRelativeUrl.replace(/\/$/, '')}/Lists/${encodeURIComponent(title)}`;
  const locked = !!busy || running;

  const policySummary = (): React.ReactNode[] => {
    const lines: React.ReactNode[] = [];
    if (!form.enabled) lines.push(<>The policy is <strong>disabled</strong>; nothing will be archived.</>);
    if (form.mode !== 'count') {
      lines.push(form.archiveBefore
        ? <>Completed or cancelled tasks finished <strong>before {formatDate(form.archiveBefore)}</strong> are moved to the archive.</>
        : <>Choose a cutoff date.</>);
    }
    if (form.mode !== 'date') {
      lines.push(<>If the active list still holds more than <strong>{fmt(Number(form.maxItems) || 0)}</strong> items, the oldest completed or cancelled tasks are archived until it does not.</>);
    }
    lines.push(<>Open work is never archived. Items with attachments are skipped and kept.</>);
    lines.push(<>Each item is copied, verified and audited before it is removed from the active list.</>);
    return lines;
  };

  const workspaceColumns: IColumn[] = [
    { key: 'title', name: 'Workspace', fieldName: 'title', minWidth: 140, maxWidth: 260, isRowHeader: true,
      onRender: (w: IWorkspaceSummary) => <strong>{w.title}</strong> },
    { key: 'active', name: 'Active list', minWidth: 160, maxWidth: 260,
      onRender: (w: IWorkspaceSummary) => <Link href={listUrl(w.workItemsList)} target="_blank" data-interception="off">{w.workItemsList}</Link> },
    { key: 'archive', name: 'Archive list', minWidth: 160, maxWidth: 260,
      onRender: (w: IWorkspaceSummary) => <Link href={listUrl(w.archiveList)} target="_blank" data-interception="off">{w.archiveList}</Link> },
    { key: 'select', name: '', minWidth: 110,
      onRender: (w: IWorkspaceSummary) => workspace && workspace.id === w.id
        ? <span className={`${styles.badge} ${styles.info}`}>Current</span>
        : <ActionButton iconProps={{ iconName: 'Switch' }} disabled={locked} onClick={() => setWorkspace(w)} ariaLabel={`Switch to ${w.title}`}>Select</ActionButton> }
  ];

  const jobColumns: IColumn[] = [
    { key: 'started', name: 'Started', minWidth: 140, maxWidth: 180, onRender: (j: IArchiveJob) => formatDateTime(j.Started) },
    { key: 'status', name: 'Status', minWidth: 150, maxWidth: 180, onRender: (j: IArchiveJob) => <StatusBadge status={j.Status} /> },
    { key: 'mode', name: 'Method', fieldName: 'Mode', minWidth: 60, maxWidth: 80 },
    { key: 'archived', name: 'Archived', minWidth: 70, maxWidth: 90, onRender: (j: IArchiveJob) => fmt(j.Archived) },
    { key: 'skipped', name: 'Skipped', minWidth: 70, maxWidth: 90, onRender: (j: IArchiveJob) => fmt(j.Skipped) },
    { key: 'failed', name: 'Failed', minWidth: 60, maxWidth: 80, onRender: (j: IArchiveJob) => fmt(j.FailedCount) },
    { key: 'error', name: 'Notes', minWidth: 160, isMultiline: true,
      onRender: (j: IArchiveJob) => j.LastError
        ? <TooltipHost content={j.LastError}><span>{j.LastError.length > 90 ? `${j.LastError.slice(0, 90)}…` : j.LastError}</span></TooltipHost>
        : '' }
  ];

  // --- Render ------------------------------------------------------------------
  const header = (
    <header className={styles.header}>
      <div className={styles.brand}>
        <div className={styles.logo} aria-hidden="true">V</div>
        <div>
          <h2 className={styles.title}>{props.title}</h2>
          <p className={styles.subtitle}>{props.subtitle}</p>
        </div>
      </div>
      {initialized && workspaces.length > 0 && (
        <div className={styles.headerActions}>
          <Dropdown
            className={styles.workspacePicker}
            ariaLabel="Current workspace"
            placeholder="Choose a workspace"
            selectedKey={workspace ? workspace.id : undefined}
            disabled={locked}
            options={workspaces.map(w => ({ key: w.id, text: w.title }))}
            onChange={(_, o) => setWorkspace(workspaces.find(w => w.id === String(o?.key)))}
          />
        </div>
      )}
    </header>
  );

  const busyIndicator = (busy || policyLoading) && !running
    ? <div className={styles.messages} aria-live="polite"><ProgressIndicator label={busy || 'Loading workspace…'} /></div>
    : undefined;

  const messages = notice && (
    <div className={styles.messages}>
      <MessageBar messageBarType={notice.type} isMultiline onDismiss={() => setNotice(undefined)} dismissButtonAriaLabel="Close">
        {notice.text}
      </MessageBar>
    </div>
  );

  let content: React.ReactNode;
  if (loading) {
    content = <div className={styles.emptyState}><Spinner size={SpinnerSize.large} label="Loading VANTAGE…" /></div>;
  } else if (!initialized) {
    const partial = setup && setup.missingLists.length > 0 && setup.missingLists.length < 5;
    content = (
      <div className={`${styles.card} ${styles.emptyState}`} style={{ marginTop: 16 }}>
        <h3 className={styles.cardTitle}>Set up VANTAGE on this site</h3>
        <p className={styles.muted}>
          A one-time setup creates five control lists (settings, workspace registry,
          archive policies, job history and audit log) in this site. It uses your SharePoint permissions only.
        </p>
        {partial && (
          <MessageBar messageBarType={MessageBarType.warning}>
            A previous setup did not finish. Missing: {setup?.missingLists.join(', ')}. Run setup again to repair it.
          </MessageBar>
        )}
        {!canManageLists && (
          <MessageBar messageBarType={MessageBarType.info}>
            You need permission to manage lists on this site (Site Owner or Designer) to run setup. Ask a site owner to open this page.
          </MessageBar>
        )}
        <div className={styles.actions} style={{ justifyContent: 'center' }}>
          <PrimaryButton text={partial ? 'Repair setup' : 'Set up VANTAGE'} iconProps={{ iconName: 'Settings' }}
            onClick={initialize} disabled={locked || !canManageLists} />
        </div>
      </div>
    );
  } else {
    content = (
      <Pivot aria-label="VANTAGE sections" selectedKey={tab} onLinkClick={item => item && setTab(item.props.itemKey as string)} overflowBehavior="menu">
        <PivotItem headerText="Overview" itemKey="overview" itemIcon="ViewDashboard">
          <div className={styles.section}>
            {workspaces.length === 0 ? (
              <div className={`${styles.card} ${styles.emptyState}`}>
                <h3 className={styles.cardTitle}>No workspaces yet</h3>
                <p className={styles.muted}>A workspace gets its own active work-item list and archive list.</p>
                <PrimaryButton text="Create a workspace" onClick={() => setTab('workspaces')} />
              </div>
            ) : (
              <div className={styles.grid}>
                <Kpi label="Workspaces" value={fmt(workspaces.length)} note="Active on this site" />
                <Kpi label="Active items" value={fmt(stats.activeItems)} note={workspace ? workspace.title : ''} />
                <Kpi label="Archived items" value={fmt(stats.archivedItems)} note="In the workspace archive list" />
                <Kpi label="Last archive run" value={stats.lastJob ? <StatusBadge status={stats.lastJob.Status} /> : 'Never'}
                  note={stats.lastJob ? formatDateTime(stats.lastJob.Started) : 'No runs for this workspace yet'} />
              </div>
            )}
          </div>
        </PivotItem>

        <PivotItem headerText="Workspaces" itemKey="workspaces" itemIcon="Org">
          <div className={styles.section}>
            <div className={styles.card}>
              <h3 className={styles.cardTitle}>Create a workspace</h3>
              <p className={styles.muted}>Creates an active work-item list and a matching archive list, then registers the workspace.</p>
              <form className={styles.inlineForm} onSubmit={e => { e.preventDefault(); if (newName.trim() && !locked && canManageLists) createWorkspace().catch(() => undefined); }}>
                <TextField label="Workspace name" value={newName} maxLength={120} placeholder="e.g. IT Operations"
                  onChange={(_, v) => setNewName(v || '')} disabled={!canManageLists} />
                <PrimaryButton type="submit" text="Create workspace" iconProps={{ iconName: 'Add' }}
                  disabled={locked || !newName.trim() || !canManageLists} />
              </form>
              {!canManageLists && <p className={styles.muted} style={{ marginTop: 8 }}>Only users who can manage lists on this site can create workspaces.</p>}
            </div>
            {workspaces.length > 0 && (
              <div className={styles.card} style={{ marginTop: 16 }}>
                <h3 className={styles.cardTitle}>All workspaces</h3>
                <DetailsList items={workspaces} columns={workspaceColumns} selectionMode={SelectionMode.none}
                  layoutMode={DetailsListLayoutMode.justified} getKey={(w: IWorkspaceSummary) => w.id} compact />
              </div>
            )}
          </div>
        </PivotItem>

        <PivotItem headerText="Archive & retention" itemKey="archive" itemIcon="Archive">
          <div className={styles.section}>
            {!workspace ? (
              <MessageBar>Create or select a workspace to configure its archive policy.</MessageBar>
            ) : (
              <div className={styles.twoCol}>
                <div className={styles.card}>
                  <h3 className={styles.cardTitle}>Archive policy · {workspace.title}</h3>
                  <p className={styles.muted}>Only tasks with status Completed or Cancelled are eligible.</p>
                  <div className={styles.formStack}>
                    <Toggle label="Policy enabled" inlineLabel checked={form.enabled} disabled={locked}
                      onChange={(_, v) => setForm({ ...form, enabled: !!v })} />
                    <Dropdown label="Archive method" selectedKey={form.mode} options={MODE_OPTIONS} disabled={locked}
                      onChange={(_, o) => setForm({ ...form, mode: o?.key as ArchiveMode })} />
                    {form.mode !== 'count' && (
                      <DatePicker label="Archive items completed before" value={form.archiveBefore} maxDate={new Date()}
                        firstDayOfWeek={DayOfWeek.Monday} formatDate={formatDate} placeholder="Select a date" disabled={locked}
                        isRequired allowTextInput parseDateFromString={parseDate} onSelectDate={d => setForm({ ...form, archiveBefore: d || undefined })} />
                    )}
                    {form.mode !== 'date' && (
                      <TextField label="Maximum active items" type="number" min={1} value={form.maxItems} disabled={locked}
                        description="Total item count of the active list, including non-task records."
                        onChange={(_, v) => setForm({ ...form, maxItems: v || '' })} />
                    )}
                    <TextField label="Batch size" type="number" min={1} max={MAX_BATCH_SIZE} value={form.batchSize} disabled={locked}
                      description={`Items read per request (1–${MAX_BATCH_SIZE}). Lower values reduce throttling risk.`}
                      onChange={(_, v) => setForm({ ...form, batchSize: v || '' })} />
                  </div>
                  {validationError && <div className={styles.messages}><MessageBar messageBarType={MessageBarType.severeWarning}>{validationError}</MessageBar></div>}
                  <div className={styles.actions}>
                    <PrimaryButton text="Save policy" iconProps={{ iconName: 'Save' }} onClick={savePolicy}
                      disabled={locked || !dirty || !!validationError} />
                    {dirty && <DefaultButton text="Discard changes" onClick={() => setForm(savedForm)} disabled={locked} />}
                    <TooltipHost content={
                      !canArchive ? 'You need permission to add, edit and delete items in this site.'
                        : dirty ? 'Save the policy before running it.'
                          : !savedForm.enabled ? 'Enable and save the policy first.' : undefined}>
                      <DefaultButton text="Run archive now" iconProps={{ iconName: 'Play' }} onClick={() => setConfirmOpen(true)}
                        disabled={locked || dirty || !savedForm.enabled || !canArchive || !!validationError} />
                    </TooltipHost>
                  </div>
                </div>

                <div>
                  <div className={styles.card}>
                    <h3 className={styles.cardTitle}>What this policy does</h3>
                    <ul className={styles.summaryList}>{policySummary().map((l, i) => <li key={i}>{l}</li>)}</ul>
                  </div>
                  {progress && (
                    <div className={styles.card} style={{ marginTop: 16 }} aria-live="polite">
                      <h3 className={styles.cardTitle}>{running ? 'Archive run in progress' : 'Last run in this session'}</h3>
                      {running && <ProgressIndicator description="Keep this page open until the run finishes." />}
                      <p className={styles.muted}>{progress.message}</p>
                      <div className={styles.counters}>
                        <div className={styles.counter}><strong>{fmt(progress.processed)}</strong><span>Processed</span></div>
                        <div className={styles.counter}><strong className={styles.ok}>{fmt(progress.archived)}</strong><span>Archived</span></div>
                        <div className={styles.counter}><strong className={styles.neutral}>{fmt(progress.skipped)}</strong><span>Skipped</span></div>
                        <div className={styles.counter}><strong className={progress.failed ? styles.bad : undefined}>{fmt(progress.failed)}</strong><span>Failed</span></div>
                      </div>
                      {running && <div className={styles.actions}><DefaultButton text="Stop after current item" iconProps={{ iconName: 'Stop' }} onClick={() => archiver.cancel()} /></div>}
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        </PivotItem>

        <PivotItem headerText="Job history" itemKey="jobs" itemIcon="History">
          <div className={styles.section}>
            <div className={styles.card}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8 }}>
                <h3 className={styles.cardTitle}>Recent archive runs{workspace ? ` · ${workspace.title}` : ''}</h3>
                <ActionButton iconProps={{ iconName: 'Refresh' }} disabled={!workspace || locked}
                  onClick={() => { if (workspace) refreshWorkspaceData(workspace).catch(() => undefined); }}>Refresh</ActionButton>
              </div>
              {jobs.length === 0
                ? <p className={styles.muted}>No archive runs recorded for this workspace.</p>
                : <DetailsList items={jobs} columns={jobColumns} selectionMode={SelectionMode.none}
                    layoutMode={DetailsListLayoutMode.justified} getKey={(j: IArchiveJob) => String(j.Id)} compact />}
              <p className={styles.muted} style={{ marginTop: 12 }}>
                Per-item outcomes are recorded in the <Link href={listUrl(LISTS.log)} target="_blank" data-interception="off">{LISTS.log}</Link> list.
              </p>
            </div>
          </div>
        </PivotItem>
      </Pivot>
    );
  }

  return (
    <ThemeProvider theme={theme as ITheme | undefined} applyTo="none">
      <section className={styles.root} style={themeVars(theme)}>
        {header}
        <div className={styles.body}>
          {busyIndicator}
          {messages}
          {content}
        </div>
      </section>

      <Dialog
        hidden={!confirmOpen}
        onDismiss={() => setConfirmOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Run archive now?',
          subText: workspace ? `Eligible items in “${workspace.title}” will be moved from ${workspace.workItemsList} to ${workspace.archiveList}.` : ''
        }}
        modalProps={{ isBlocking: true }}
        minWidth={420}
      >
        <ul className={styles.summaryList}>{policySummary().slice(form.enabled ? 0 : 1).map((l, i) => <li key={i}>{l}</li>)}</ul>
        <p className={styles.muted} style={{ marginTop: 12 }}>
          Keep this page open while the run is in progress. Test on non-production data first.
        </p>
        <DialogFooter>
          <PrimaryButton text="Run archive" onClick={() => { runArchive().catch(() => undefined); }} />
          <DefaultButton text="Cancel" onClick={() => setConfirmOpen(false)} />
        </DialogFooter>
      </Dialog>
    </ThemeProvider>
  );
};
