import * as React from 'react';
import {
  ActionButton, DatePicker, DayOfWeek, DefaultButton, DetailsList, DetailsListLayoutMode, Dropdown, type IColumn, Link,
  MessageBar, MessageBarType, Pivot, PivotItem, PrimaryButton, ProgressIndicator, SearchBox, SelectionMode, TextField, Toggle, TooltipHost
} from '@fluentui/react';

import styles from './VantageApp.module.scss';
import {
  Badge, Card, fmt, formatDate, formatDateTime, fromUtcIso, listUrl, parseDate, toUtcIso, useApp, useConfirm
} from './common';
import { ArchiveService } from '../services/ArchiveService';
import { LISTS } from '../services/ProvisioningService';
import { SharePointService, type SPEntity } from '../services/SharePointService';
import {
  type ArchiveMode, DEFAULT_ELIGIBLE_STATUSES, type IArchivedTask, type IArchiveJob, type IArchiveLogEntry, type IArchivePolicy,
  type IArchiveProgress, MAX_ATTACHMENT_BYTES, MAX_BATCH_SIZE
} from '../models/Models';

interface IPolicyForm {
  enabled: boolean;
  mode: ArchiveMode;
  maxItems: string;
  archiveBefore: Date | undefined;
  batchSize: string;
}

const MODE_OPTIONS = [
  { key: 'date', text: 'Archive tasks completed before a date' },
  { key: 'count', text: 'Keep the active list under a maximum size' },
  { key: 'both', text: 'Both: date first, then size limit' }
];
const DEFAULT_FORM: IPolicyForm = { enabled: false, mode: 'date', maxItems: '100000', archiveBefore: undefined, batchSize: '100' };

const formToPolicy = (workspaceId: string, f: IPolicyForm): IArchivePolicy => ({
  WorkspaceId: workspaceId, Enabled: f.enabled, Mode: f.mode, MaxActiveItems: Math.floor(Number(f.maxItems)),
  ArchiveBefore: f.archiveBefore ? toUtcIso(f.archiveBefore) : undefined, EligibleStatuses: DEFAULT_ELIGIBLE_STATUSES,
  BatchSize: Math.floor(Number(f.batchSize))
});
const formsEqual = (a: IPolicyForm, b: IPolicyForm): boolean =>
  a.enabled === b.enabled && a.mode === b.mode && a.maxItems === b.maxItems && a.batchSize === b.batchSize &&
  (a.archiveBefore ? a.archiveBefore.getTime() : 0) === (b.archiveBefore ? b.archiveBefore.getTime() : 0);

export const ArchiveTab: React.FC<{ onRunningChange: (running: boolean) => void }> = ({ onRunningChange }) => {
  const { sp, archiver, workspace, perform, busy, notify, user, webServerRelativeUrl, bumpRefresh } = useApp();
  const [form, setForm] = React.useState<IPolicyForm>(DEFAULT_FORM);
  const [saved, setSaved] = React.useState<IPolicyForm>(DEFAULT_FORM);
  const [policyId, setPolicyId] = React.useState<number | undefined>(undefined);
  const [loaded, setLoaded] = React.useState(false);
  const [running, setRunning] = React.useState(false);
  const [progress, setProgress] = React.useState<IArchiveProgress | undefined>(undefined);
  const [archived, setArchived] = React.useState<IArchivedTask[]>([]);
  const [jobs, setJobs] = React.useState<IArchiveJob[]>([]);
  const [log, setLog] = React.useState<IArchiveLogEntry[]>([]);
  const [archiveSearch, setArchiveSearch] = React.useState('');
  const [confirm, confirmDialog] = useConfirm();

  const refreshLists = React.useCallback(async (): Promise<void> => {
    if (!workspace) return;
    const [a, j, l] = await Promise.all([
      archiver.getArchivedTasks(workspace).catch(() => [] as IArchivedTask[]),
      archiver.getRecentJobs(workspace.id).catch(() => [] as IArchiveJob[]),
      archiver.getLog(workspace.id).catch(() => [] as IArchiveLogEntry[])
    ]);
    setArchived(a); setJobs(j); setLog(l);
  }, [workspace, archiver]);

  React.useEffect(() => {
    if (!workspace) return;
    let live = true;
    setLoaded(false);
    (async () => {
      const r = await sp.getCollection<SPEntity>(
        `${sp.listPath(LISTS.policies)}/items?$select=Id,Enabled,ArchiveMode,MaxActiveItems,ArchiveBefore,BatchSize` +
        `&${SharePointService.filter(`WorkspaceId eq '${SharePointService.odataString(workspace.id)}'`)}&$top=1`
      );
      if (!live) return;
      const p = r.value[0];
      const f: IPolicyForm = p ? {
        enabled: !!p.Enabled,
        mode: (['count', 'date', 'both'].indexOf(p.ArchiveMode as string) >= 0 ? p.ArchiveMode : 'date') as ArchiveMode,
        maxItems: String(p.MaxActiveItems || 100000),
        archiveBefore: fromUtcIso(p.ArchiveBefore as string | undefined),
        batchSize: String(p.BatchSize || 100)
      } : DEFAULT_FORM;
      setPolicyId(p ? (p.Id as number) : undefined);
      setForm(f); setSaved(f); setLoaded(true);
      await refreshLists();
    })().catch(e => { if (live) { setLoaded(true); notify(MessageBarType.error, `Could not load the archive policy: ${(e as Error).message}`); } });
    return () => { live = false; };
  }, [workspace, sp, refreshLists, notify]);

  React.useEffect(() => { onRunningChange(running); }, [running, onRunningChange]);

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

  if (!workspace) return null;
  const dirty = !formsEqual(form, saved);
  const locked = busy || running;
  const validationError = ((): string | undefined => {
    try { ArchiveService.validatePolicy(formToPolicy(workspace.id, form)); return undefined; } catch (e) { return (e as Error).message; }
  })();

  const save = (): Promise<boolean> => perform('Saving policy…', async () => {
    if (validationError) throw new Error(validationError);
    const p = formToPolicy(workspace.id, form);
    const fields: SPEntity = {
      Title: `${workspace.title} archive policy`, WorkspaceId: workspace.id, Enabled: p.Enabled, ArchiveMode: p.Mode,
      MaxActiveItems: p.MaxActiveItems, ArchiveBefore: p.ArchiveBefore || null, EligibleStatuses: p.EligibleStatuses, BatchSize: p.BatchSize
    };
    if (policyId) await sp.updateItem(LISTS.policies, policyId, fields);
    else setPolicyId((await sp.createItem(LISTS.policies, fields)).Id);
    setSaved(form);
    notify(MessageBarType.success, 'Archive policy saved.');
  });

  const summary = (f: IPolicyForm): React.ReactNode[] => {
    const lines: React.ReactNode[] = [];
    if (!f.enabled) lines.push(<>The policy is <strong>disabled</strong>; nothing will be archived.</>);
    if (f.mode !== 'count') {
      lines.push(f.archiveBefore
        ? <>Completed or cancelled tasks finished <strong>before {formatDate(f.archiveBefore)}</strong> are moved to the archive.</>
        : <>Choose a cutoff date.</>);
    }
    if (f.mode !== 'date') {
      lines.push(<>If the active list holds more than <strong>{fmt(Number(f.maxItems) || 0)}</strong> records, the oldest completed or cancelled tasks are archived until it does not.</>);
    }
    lines.push(<>Each task moves together with its comments, checklist items, dependencies, activity and attachments (up to {MAX_ATTACHMENT_BYTES / 1048576} MB each).</>);
    lines.push(<>Everything is copied and verified (including attachment checksums) before anything is removed. Open work is never archived.</>);
    return lines;
  };

  const run = async (): Promise<void> => {
    const ok = await confirm({
      title: 'Run archive now?', confirmText: 'Run archive',
      text: <>
        <p>Eligible tasks in “{workspace.title}” will be moved from {workspace.workItemsList} to {workspace.archiveList}.</p>
        <ul className={styles.summaryList}>{summary(saved).map((l, i) => <li key={i}>{l}</li>)}</ul>
        <p className={styles.muted}>Keep this page open until the run finishes. Test with non-production data first.</p>
      </>
    });
    if (!ok) return;
    setRunning(true);
    setProgress({ processed: 0, archived: 0, records: 0, failed: 0, skipped: 0, message: 'Starting archive run…' });
    try {
      const r = await archiver.run(workspace, formToPolicy(workspace.id, saved), user.displayName, p => setProgress(p));
      notify(r.failed || r.remainingOverLimit ? MessageBarType.warning : MessageBarType.success,
        `Archive run finished: ${fmt(r.archived)} tasks (${fmt(r.records)} records) archived, ${fmt(r.skipped)} skipped, ${fmt(r.failed)} failed.` +
        (r.remainingOverLimit ? ` ${fmt(r.remainingOverLimit)} records remain above the maximum because there were not enough eligible tasks.` : ''));
    } catch (e) {
      notify(MessageBarType.error, (e as Error).message);
    } finally {
      setRunning(false);
      bumpRefresh();
      await refreshLists().catch(() => undefined);
    }
  };

  const restore = async (t: IArchivedTask): Promise<void> => {
    const ok = await confirm({
      title: 'Restore this task?', confirmText: 'Restore',
      text: <>“{t.Title}” and its related records will be moved back to the active list as a new task. If it is still Completed or Cancelled, the next archive run may archive it again — reopen it if you want to keep working on it.</>
    });
    if (!ok) return;
    await perform('Restoring task…', async () => {
      const id = await archiver.restore(workspace, t.Id);
      notify(MessageBarType.success, `“${t.Title}” was restored as task #${id}.`);
      bumpRefresh();
      await refreshLists();
    });
  };

  const archivedColumns: IColumn[] = [
    { key: 'title', name: 'Task', fieldName: 'Title', minWidth: 180, isRowHeader: true, isResizable: true },
    { key: 'status', name: 'Status', minWidth: 90, maxWidth: 110, onRender: (t: IArchivedTask) => <Badge value={t.WorkStatus} /> },
    { key: 'who', name: 'Assigned to', fieldName: 'AssignedToName', minWidth: 110, maxWidth: 180 },
    { key: 'done', name: 'Completed', minWidth: 90, maxWidth: 120, onRender: (t: IArchivedTask) => formatDate(t.CompletedDate) || '—' },
    { key: 'at', name: 'Archived', minWidth: 90, maxWidth: 120, onRender: (t: IArchivedTask) => formatDate(t.ArchivedAt) || '—' },
    { key: 'act', name: '', minWidth: 90,
      onRender: (t: IArchivedTask) => <ActionButton iconProps={{ iconName: 'Undo' }} disabled={locked} onClick={() => { restore(t).catch(() => undefined); }} ariaLabel={`Restore ${t.Title}`}>Restore</ActionButton> }
  ];
  const jobColumns: IColumn[] = [
    { key: 'started', name: 'Started', minWidth: 140, maxWidth: 170, onRender: (j: IArchiveJob) => formatDateTime(j.Started) },
    { key: 'status', name: 'Status', minWidth: 140, maxWidth: 170, onRender: (j: IArchiveJob) => <Badge value={j.Status} /> },
    { key: 'mode', name: 'Method', fieldName: 'Mode', minWidth: 55, maxWidth: 70 },
    { key: 'tasks', name: 'Tasks', minWidth: 55, maxWidth: 70, onRender: (j: IArchiveJob) => fmt(j.Archived) },
    { key: 'records', name: 'Records', minWidth: 60, maxWidth: 75, onRender: (j: IArchiveJob) => fmt(j.Records) },
    { key: 'skipped', name: 'Skipped', minWidth: 60, maxWidth: 75, onRender: (j: IArchiveJob) => fmt(j.Skipped) },
    { key: 'failed', name: 'Failed', minWidth: 55, maxWidth: 70, onRender: (j: IArchiveJob) => fmt(j.FailedCount) },
    { key: 'note', name: 'Notes', minWidth: 160, isMultiline: true,
      onRender: (j: IArchiveJob) => j.LastError ? <TooltipHost content={j.LastError}><span>{j.LastError.length > 90 ? `${j.LastError.slice(0, 90)}…` : j.LastError}</span></TooltipHost> : '' }
  ];
  const logColumns: IColumn[] = [
    { key: 'when', name: 'When', minWidth: 140, maxWidth: 170, onRender: (l: IArchiveLogEntry) => formatDateTime(l.Created) },
    { key: 'outcome', name: 'Outcome', minWidth: 80, maxWidth: 100,
      onRender: (l: IArchiveLogEntry) => <Badge value={l.Outcome} tone={l.Outcome === 'Failed' ? styles.bad : l.Outcome === 'Skipped' ? styles.warn : styles.ok} /> },
    { key: 'item', name: 'Item', minWidth: 60, maxWidth: 80, onRender: (l: IArchiveLogEntry) => l.OriginalItemId ? `#${l.OriginalItemId}` : '' },
    { key: 'detail', name: 'Detail', fieldName: 'Detail', minWidth: 200, isMultiline: true }
  ];

  const term = archiveSearch.trim().toLowerCase();
  const archivedVisible = archived.filter(a => !term || (a.Title || '').toLowerCase().indexOf(term) >= 0);

  return (
    <div className={styles.stack}>
      <div className={styles.twoCol}>
        <Card title={`Archive policy · ${workspace.title}`}>
          <p className={styles.muted}>Only tasks with status Completed or Cancelled are eligible.</p>
          <div className={styles.formStack}>
            <Toggle label="Policy enabled" inlineLabel checked={form.enabled} disabled={locked || !loaded} onChange={(_, v) => setForm({ ...form, enabled: !!v })} />
            <Dropdown label="Archive method" selectedKey={form.mode} options={MODE_OPTIONS} disabled={locked || !loaded}
              onChange={(_, o) => setForm({ ...form, mode: o?.key as ArchiveMode })} />
            {form.mode !== 'count' && (
              <DatePicker label="Archive tasks completed before" value={form.archiveBefore} maxDate={new Date()} firstDayOfWeek={DayOfWeek.Monday}
                formatDate={formatDate} placeholder="Select a date" disabled={locked || !loaded} isRequired allowTextInput parseDateFromString={parseDate}
                onSelectDate={d => setForm({ ...form, archiveBefore: d || undefined })} />
            )}
            {form.mode !== 'date' && (
              <TextField label="Maximum active records" type="number" min={1} value={form.maxItems} disabled={locked || !loaded}
                description="Counts every record in the active list, including comments, checklist items and activity."
                onChange={(_, v) => setForm({ ...form, maxItems: v || '' })} />
            )}
            <TextField label="Batch size" type="number" min={1} max={MAX_BATCH_SIZE} value={form.batchSize} disabled={locked || !loaded}
              description={`Tasks read per request (1–${MAX_BATCH_SIZE}). Lower values reduce throttling risk.`}
              onChange={(_, v) => setForm({ ...form, batchSize: v || '' })} />
          </div>
          {validationError && <div className={styles.messages}><MessageBar messageBarType={MessageBarType.severeWarning}>{validationError}</MessageBar></div>}
          <div className={styles.actions}>
            <PrimaryButton text="Save policy" iconProps={{ iconName: 'Save' }} onClick={() => { save().catch(() => undefined); }} disabled={locked || !dirty || !!validationError} />
            {dirty && <DefaultButton text="Discard changes" onClick={() => setForm(saved)} disabled={locked} />}
            <TooltipHost content={dirty ? 'Save the policy before running it.' : !saved.enabled ? 'Enable and save the policy first.' : undefined}>
              <DefaultButton text="Run archive now" iconProps={{ iconName: 'Play' }} onClick={() => { run().catch(() => undefined); }}
                disabled={locked || dirty || !saved.enabled || !!validationError} />
            </TooltipHost>
          </div>
        </Card>

        <div className={styles.stack}>
          <Card title="What this policy does">
            <ul className={styles.summaryList}>{summary(form).map((l, i) => <li key={i}>{l}</li>)}</ul>
          </Card>
          {progress && (
            <Card title={running ? 'Archive run in progress' : 'Last run in this session'}>
              <div aria-live="polite">
                {running && <ProgressIndicator description="Keep this page open until the run finishes." />}
                <p className={styles.muted}>{progress.message}</p>
                <div className={styles.counters}>
                  <div className={styles.counter}><strong className={styles.ok}>{fmt(progress.archived)}</strong><span>Tasks</span></div>
                  <div className={styles.counter}><strong>{fmt(progress.records)}</strong><span>Records</span></div>
                  <div className={styles.counter}><strong className={styles.neutral}>{fmt(progress.skipped)}</strong><span>Skipped</span></div>
                  <div className={styles.counter}><strong className={progress.failed ? styles.bad : undefined}>{fmt(progress.failed)}</strong><span>Failed</span></div>
                </div>
                {running && <div className={styles.actions}><DefaultButton text="Stop after current task" iconProps={{ iconName: 'Stop' }} onClick={() => archiver.cancel()} /></div>}
              </div>
            </Card>
          )}
        </div>
      </div>

      <Card>
        <Pivot aria-label="Archive records" overflowBehavior="menu">
          <PivotItem headerText="Archived tasks" itemKey="archived" itemCount={archived.length}>
            <div className={styles.section}>
              <div className={styles.filterBar}>
                <SearchBox placeholder="Filter archived tasks" value={archiveSearch} onChange={(_, v) => setArchiveSearch(v || '')} className={styles.search} />
                <ActionButton iconProps={{ iconName: 'Refresh' }} disabled={locked} onClick={() => { refreshLists().catch(() => undefined); }}>Refresh</ActionButton>
              </div>
              {archivedVisible.length
                ? <DetailsList items={archivedVisible} columns={archivedColumns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} compact getKey={(t: IArchivedTask) => String(t.Id)} />
                : <p className={styles.muted}>{archived.length ? 'No archived tasks match the filter.' : 'Nothing has been archived for this workspace yet.'}</p>}
              <p className={styles.muted}>Showing the 100 most recently archived tasks. Open the <Link href={listUrl(webServerRelativeUrl, workspace.archiveList)} target="_blank" data-interception="off">archive list</Link> for everything.</p>
            </div>
          </PivotItem>
          <PivotItem headerText="Job history" itemKey="jobs">
            <div className={styles.section}>
              {jobs.length
                ? <DetailsList items={jobs} columns={jobColumns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} compact getKey={(j: IArchiveJob) => String(j.Id)} />
                : <p className={styles.muted}>No archive runs recorded for this workspace.</p>}
            </div>
          </PivotItem>
          <PivotItem headerText="Audit log" itemKey="log">
            <div className={styles.section}>
              {log.length
                ? <DetailsList items={log} columns={logColumns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} compact getKey={(l: IArchiveLogEntry) => String(l.Id)} />
                : <p className={styles.muted}>No audit entries yet.</p>}
              <p className={styles.muted}>The 50 most recent entries. The full log is in the <Link href={listUrl(webServerRelativeUrl, LISTS.log)} target="_blank" data-interception="off">{LISTS.log}</Link> list.</p>
            </div>
          </PivotItem>
        </Pivot>
      </Card>
      {confirmDialog}
    </div>
  );
};
