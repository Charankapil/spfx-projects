import * as React from 'react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  Checkbox,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  Icon,
  MessageBar,
  MessageBarType,
  Pivot,
  PivotItem,
  PrimaryButton,
  ProgressIndicator,
  Stack,
  Text
} from '@fluentui/react';
import { SPPermission } from '@microsoft/sp-page-context';

import { IReInheritProps } from './IReInheritProps';
import { Logo } from './Logo';
import { OptionsPanel } from './OptionsPanel';
import { ReportHistory } from './ReportHistory';
import { computeStats, KIND_ICONS, ReportView } from './ReportView';
import { ScopePicker } from './ScopePicker';
import { formatCompact } from './format';
import styles from './ReInherit.module.scss';
import { DEFAULT_OPTIONS, IRestoreOptions } from '../models/IRestoreOptions';
import { IDLE_PROGRESS, IRunProgress } from '../models/IRunProgress';
import { IRunReport, KIND_LABELS } from '../models/IRunReport';
import { IScopeNode } from '../models/IScopeNode';
import { IUniqueObject, ObjectKind } from '../models/IUniqueObject';
import { InheritanceRestorer } from '../services/InheritanceRestorer';
import { ReportStore } from '../services/ReportStore';
import { ScopeService } from '../services/ScopeService';
import { CancelledError, newId, SpRest } from '../services/SpRest';
import { UniquePermissionScanner } from '../services/UniquePermissionScanner';

type Stage = 'setup' | 'scanning' | 'review' | 'restoring' | 'done';

interface INotice {
  type: MessageBarType;
  text: string;
}

const STEPS = ['Choose scope', 'Find unique permissions', 'Review', 'Restore inheritance', 'Report'];
const STEP_OF: { [stage in Stage]: number } = { setup: 0, scanning: 1, review: 2, restoring: 3, done: 4 };

/** Found in the scan and not excluded; or failed / left unprocessed by an earlier restore pass. */
function isRestorable(o: IUniqueObject): boolean {
  return o.status === 'found' || o.status === 'failed' || o.status === 'skipped';
}

function errorText(err: unknown, fallback: string): string {
  return err instanceof Error ? err.message : fallback;
}

export const ReInherit: React.FC<IReInheritProps> = (props) => {
  const { context } = props;
  const restRef = useRef<SpRest>();
  if (!restRef.current) {
    restRef.current = new SpRest(context);
  }
  const rest = restRef.current;
  const scopeService = useMemo(() => new ScopeService(context, rest), [context, rest]);
  const store = useMemo(() => new ReportStore(context, rest), [context, rest]);

  // Resetting inheritance needs Manage Permissions; without it every reset would 403.
  const canManage = useMemo(() => {
    const ctx = context.pageContext;
    const legacy = ctx.legacyPageContext as { isSiteAdmin?: boolean } | undefined;
    return !!(legacy && legacy.isSiteAdmin) || ctx.web.permissions.hasPermission(SPPermission.managePermissions);
  }, [context]);

  const [tab, setTab] = useState<string>(canManage ? 'cleanup' : 'reports');
  const [stage, setStage] = useState<Stage>('setup');
  const [selected, setSelected] = useState<Map<string, IScopeNode>>(new Map());
  const [options, setOptions] = useState<IRestoreOptions>(DEFAULT_OPTIONS);
  const [progress, setProgress] = useState<IRunProgress>(IDLE_PROGRESS);
  const [report, setReport] = useState<IRunReport | undefined>(undefined);
  const [notice, setNotice] = useState<INotice | undefined>(undefined);
  const [confirmOpen, setConfirmOpen] = useState(false);
  const [confirmChecked, setConfirmChecked] = useState(false);
  const [historyToken, setHistoryToken] = useState(0);
  const [cancelRequested, setCancelRequested] = useState(false);

  const busy = stage === 'scanning' || stage === 'restoring';

  // A scan or restore runs in this browser tab; leaving the page stops it.
  useEffect(() => {
    if (!busy) {
      return undefined;
    }
    const handler = (e: BeforeUnloadEvent): string => {
      e.preventDefault();
      e.returnValue = 'ReInherit is still running. Leaving the page stops it.';
      return e.returnValue;
    };
    window.addEventListener('beforeunload', handler);
    return () => window.removeEventListener('beforeunload', handler);
  }, [busy]);

  const toggleScope = useCallback((node: IScopeNode, checked: boolean) => {
    setSelected((prev) => {
      const next = new Map<string, IScopeNode>();
      prev.forEach((value, key) => next.set(key, value));
      if (checked) {
        next.set(node.key, node);
      } else {
        next.delete(node.key);
      }
      return next;
    });
  }, []);

  const saveReport = useCallback(
    async (r: IRunReport, successText: string): Promise<void> => {
      try {
        await store.save(r);
        setHistoryToken((t) => t + 1);
        setNotice({ type: MessageBarType.success, text: successText });
      } catch (err) {
        setNotice({
          type: MessageBarType.warning,
          text: `The report could not be saved to the site (${errorText(err, 'unknown error')}). Export it to CSV before leaving the page.`
        });
      }
    },
    [store]
  );

  const startScan = useCallback(() => {
    const scopes: IScopeNode[] = [];
    selected.forEach((node) => scopes.push(node));
    if (scopes.length === 0) {
      return;
    }
    rest.reset();
    setCancelRequested(false);
    setNotice(undefined);
    setReport(undefined);
    setProgress({ ...IDLE_PROGRESS, phase: 'scanning' });
    setStage('scanning');

    const run = async (): Promise<void> => {
      const startedAt = new Date().toISOString();
      const scanner = new UniquePermissionScanner(
        rest,
        scopeService,
        context.pageContext.site.serverRelativeUrl,
        options,
        (p) => setProgress(p)
      );
      try {
        const result = await scanner.scan(scopes);
        const r: IRunReport = {
          id: newId(),
          siteUrl: context.pageContext.site.absoluteUrl,
          mode: 'scan',
          runBy: context.pageContext.user.displayName,
          startedAt,
          scanCompletedAt: new Date().toISOString(),
          completedAt: new Date().toISOString(),
          scopes: scopes.map((s) => ({ kind: s.kind, title: s.title, url: s.serverRelativeUrl })),
          options: { ...options },
          stats: result.stats,
          objects: result.objects,
          scanErrors: result.errors
        };
        r.stats = computeStats(r);
        setReport(r);
        setStage('review');
        await saveReport(
          r,
          r.objects.length > 0
            ? 'Scan finished and saved to Reports. Review what was found, exclude anything that should stay unique, then restore inheritance.'
            : 'Scan finished and saved to Reports. Nothing in this scope has unique permissions.'
        );
      } catch (err) {
        setStage('setup');
        setNotice(
          err instanceof CancelledError
            ? { type: MessageBarType.info, text: 'Scan cancelled. Nothing was changed.' }
            : { type: MessageBarType.error, text: errorText(err, 'The scan failed.') }
        );
      }
    };
    run().catch(() => undefined);
  }, [selected, options, rest, scopeService, context, saveReport]);

  const startRestore = useCallback(() => {
    if (!report) {
      return;
    }
    setConfirmOpen(false);
    setConfirmChecked(false);
    rest.reset();
    setCancelRequested(false);
    setNotice(undefined);

    const targets = report.objects.filter(isRestorable);
    const r: IRunReport = { ...report, mode: 'restore', restoreStartedAt: new Date().toISOString() };
    setReport(r);
    setStage('restoring');
    setProgress({ ...progress, phase: 'restoring', restoreDone: 0, restoreTotal: targets.length, restored: 0, failed: 0 });

    const run = async (): Promise<void> => {
      const restorer = new InheritanceRestorer(rest, (p) =>
        setProgress((prev) => ({
          ...prev,
          phase: 'restoring',
          currentItem: p.current,
          restoreDone: p.done,
          restoreTotal: p.total,
          restored: p.restored,
          failed: p.failed
        }))
      );
      let failure: string | undefined;
      try {
        await restorer.restore(targets, r.options.backupPermissions);
      } catch (err) {
        if (err instanceof CancelledError) {
          r.cancelled = true;
        } else {
          failure = errorText(err, 'The restore stopped unexpectedly.');
        }
      }
      for (const o of targets) {
        if (o.status === 'found') {
          o.status = 'skipped';
          o.message = r.cancelled ? 'Not processed: the run was cancelled.' : 'Not processed: the run stopped early.';
        }
      }
      const final: IRunReport = { ...r, objects: [...r.objects], completedAt: new Date().toISOString() };
      final.stats = computeStats(final);
      setReport(final);
      setStage('done');
      const summary =
        `Inheritance restored on ${final.stats.restored.toLocaleString()} object(s)` +
        (final.stats.failed ? `, ${final.stats.failed.toLocaleString()} failed` : '') +
        (final.cancelled ? ' before the run was cancelled' : '') +
        '. The report is saved under Reports.';
      await saveReport(final, summary);
      if (final.stats.failed > 0 && !failure) {
        setNotice({
          type: MessageBarType.warning,
          text: `${summary} Filter the report by "Failed" to see why, then use "Retry failed" once it is fixed.`
        });
      }
      if (failure) {
        setNotice({ type: MessageBarType.error, text: `${failure} ${summary}` });
      }
    };
    run().catch(() => undefined);
  }, [report, rest, progress, saveReport]);

  const cancel = useCallback(() => {
    setCancelRequested(true);
    rest.cancel();
  }, [rest]);

  const setExcluded = useCallback(
    (objects: IUniqueObject[], excluded: boolean) => {
      if (!report) {
        return;
      }
      for (const o of objects) {
        if (excluded && o.status === 'found') {
          o.status = 'excluded';
        } else if (!excluded && o.status === 'excluded') {
          o.status = 'found';
        }
      }
      const next = { ...report, objects: [...report.objects] };
      next.stats = computeStats(next);
      setReport(next);
    },
    [report]
  );

  const startOver = useCallback(() => {
    setStage('setup');
    setReport(undefined);
    setNotice(undefined);
    setProgress(IDLE_PROGRESS);
  }, []);

  const toRestore = report ? report.objects.filter(isRestorable) : [];
  const retryable = stage === 'done' ? toRestore.length : 0;
  const restoreCounts = useMemo(() => {
    const counts: { [kind: string]: number } = {};
    toRestore.forEach((o) => {
      counts[o.kind] = (counts[o.kind] || 0) + 1;
    });
    return counts;
  }, [toRestore]);

  const currentStep = STEP_OF[stage];

  const progressLabel =
    stage === 'scanning'
      ? `Checked ${formatCompact(progress.itemsChecked)} items in ${progress.listsChecked.toLocaleString()} lists across ${progress.websChecked.toLocaleString()} sites · ${progress.uniqueFound.toLocaleString()} with unique permissions`
      : `Restored ${(progress.restored || 0).toLocaleString()} of ${(progress.restoreTotal || 0).toLocaleString()}` +
        (progress.failed ? ` · ${progress.failed.toLocaleString()} failed` : '');
  const progressPercent =
    stage === 'scanning'
      ? progress.listTitle && progress.listItemsTotal
        ? Math.min(1, (progress.listItemsDone || 0) / progress.listItemsTotal)
        : undefined
      : progress.restoreTotal
      ? (progress.restoreDone || 0) / progress.restoreTotal
      : undefined;

  const selectedList: IScopeNode[] = [];
  selected.forEach((n) => selectedList.push(n));

  return (
    <div className={styles.reInherit}>
      <div className={styles.hero}>
        <Logo size={52} className={styles.heroLogo} />
        <div className={styles.heroText}>
          <div className={styles.heroEyebrow}>Permission inheritance restorer</div>
          <h2 className={styles.heroTitle}>{props.description || 'ReInherit'}</h2>
          <div className={styles.heroMeta}>
            Find unique permissions and put sites, libraries, folders and files back on inherited permissions - with a
            report of every change.
          </div>
        </div>
      </div>

      {notice && (
        <MessageBar messageBarType={notice.type} onDismiss={() => setNotice(undefined)} isMultiline>
          {notice.text}
        </MessageBar>
      )}

      <Pivot selectedKey={tab} onLinkClick={(item) => item && setTab(item.props.itemKey || 'cleanup')} className={styles.tabs}>
        <PivotItem headerText="Clean up" itemKey="cleanup" itemIcon="Sync" />
        <PivotItem headerText="Reports" itemKey="reports" itemIcon="ReportDocument" />
      </Pivot>

      {tab === 'reports' && <ReportHistory store={store} refreshToken={historyToken} />}

      {tab === 'cleanup' && !canManage && (
        <MessageBar messageBarType={MessageBarType.info} isMultiline>
          Restoring inheritance needs the Manage Permissions permission on this site (site owners and site collection
          administrators have it). You can still open saved reports under Reports.
        </MessageBar>
      )}

      {tab === 'cleanup' && canManage && (
        <>
          <ol className={styles.steps}>
            {STEPS.map((label, i) => {
              const state = i < currentStep ? 'done' : i === currentStep ? 'active' : 'pending';
              return (
                <li key={label} className={`${styles.step} ${styles[state]}`}>
                  <span className={styles.stepDot}>{state === 'done' ? <Icon iconName="CheckMark" /> : i + 1}</span>
                  <span className={styles.stepLabel}>{label}</span>
                </li>
              );
            })}
          </ol>

          {(stage === 'setup' || stage === 'scanning') && (
            <div className={styles.setupGrid}>
              <div className={styles.card}>
                <div className={styles.cardTitle}>
                  <Icon iconName="Org" /> What to clean up
                </div>
                <div className={styles.cardHint}>
                  Tick sites, libraries, lists or folders. Expand a library to pick individual folders.{' '}
                  <span className={styles.pillUnique}>Unique</span> marks objects that have their own permissions now.
                </div>
                <ScopePicker service={scopeService} selected={selected} onToggle={toggleScope} disabled={busy} />
              </div>
              <div className={styles.card}>
                <div className={styles.cardTitle}>
                  <Icon iconName="Settings" /> Options
                </div>
                <OptionsPanel options={options} onChange={setOptions} disabled={busy} />
                <div className={styles.selectionSummary}>
                  {selectedList.length === 0 ? (
                    <span className={styles.muted}>Nothing selected yet.</span>
                  ) : (
                    <div className={styles.scopeChips}>
                      {selectedList.map((n) => (
                        <span key={n.key} className={styles.scopeChip} title={n.serverRelativeUrl}>
                          <Icon iconName={KIND_ICONS[(n.kind === 'web' ? 'web' : n.kind) as ObjectKind]} /> {n.title}
                        </span>
                      ))}
                    </div>
                  )}
                </div>
                <Stack horizontal tokens={{ childrenGap: 8 }}>
                  {stage === 'setup' ? (
                    <PrimaryButton
                      text="Find unique permissions"
                      iconProps={{ iconName: 'Search' }}
                      disabled={selectedList.length === 0}
                      onClick={startScan}
                    />
                  ) : (
                    <DefaultButton
                      text={cancelRequested ? 'Cancelling…' : 'Cancel scan'}
                      iconProps={{ iconName: 'Cancel' }}
                      disabled={cancelRequested}
                      onClick={cancel}
                    />
                  )}
                </Stack>
                <div className={styles.hint}>Scanning only reads. Nothing changes until you confirm the restore.</div>
              </div>
            </div>
          )}

          {busy && (
            <div className={styles.progressCard}>
              <ProgressIndicator
                label={stage === 'scanning' ? 'Finding unique permissions…' : 'Restoring inheritance…'}
                description={progressLabel}
                percentComplete={progressPercent}
              />
              <Text variant="small" block className={styles.currentItem}>
                {stage === 'scanning' && progress.listTitle
                  ? `Reading ${progress.listTitle} (${Math.round((progressPercent || 0) * 100)}% of item IDs)`
                  : progress.currentItem}
              </Text>
              {stage === 'restoring' && (
                <DefaultButton
                  className={styles.progressCancel}
                  text={cancelRequested ? 'Stopping after the current batch…' : 'Stop'}
                  iconProps={{ iconName: 'CircleStop' }}
                  disabled={cancelRequested}
                  onClick={cancel}
                />
              )}
            </div>
          )}

          {report && stage !== 'scanning' && (
            <ReportView
              report={report}
              selectable={stage === 'review'}
              onSetExcluded={setExcluded}
              actions={
                stage === 'review' ? (
                  <>
                    <DefaultButton text="Back" iconProps={{ iconName: 'Back' }} onClick={startOver} />
                    <PrimaryButton
                      text={`Restore inheritance (${toRestore.length.toLocaleString()})`}
                      iconProps={{ iconName: 'Sync' }}
                      disabled={toRestore.length === 0}
                      onClick={() => setConfirmOpen(true)}
                    />
                  </>
                ) : stage === 'done' ? (
                  <>
                    {retryable > 0 && (
                      <DefaultButton
                        text={`Retry failed / unprocessed (${retryable.toLocaleString()})`}
                        iconProps={{ iconName: 'Refresh' }}
                        onClick={() => setConfirmOpen(true)}
                      />
                    )}
                    <PrimaryButton text="New clean-up" iconProps={{ iconName: 'Add' }} onClick={startOver} />
                  </>
                ) : undefined
              }
            />
          )}
        </>
      )}

      <Dialog
        hidden={!confirmOpen}
        onDismiss={() => setConfirmOpen(false)}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: `Restore inheritance on ${toRestore.length.toLocaleString()} object(s)?`
        }}
        modalProps={{ isBlocking: true }}
        minWidth={420}
      >
        <ul className={styles.confirmList}>
          {Object.keys(restoreCounts).map((kind) => (
            <li key={kind}>
              <Icon iconName={KIND_ICONS[kind as ObjectKind]} /> {restoreCounts[kind].toLocaleString()}{' '}
              {KIND_LABELS[kind as ObjectKind].toLowerCase()}
              {restoreCounts[kind] === 1 ? '' : 's'}
            </li>
          ))}
        </ul>
        <p className={styles.confirmText}>
          Their unique permissions are deleted and they take their permissions from their parent again. People who were
          given access only on these objects - including through sharing links - lose it.
          {report && report.options.backupPermissions
            ? ' The permissions they have now are recorded in the report first.'
            : ' Backup is off, so the report will not record who had access.'}
        </p>
        <Checkbox
          label="I understand this can't be undone automatically"
          checked={confirmChecked}
          onChange={(_, c) => setConfirmChecked(!!c)}
        />
        <DialogFooter>
          <PrimaryButton text="Restore inheritance" disabled={!confirmChecked} onClick={startRestore} />
          <DefaultButton text="Cancel" onClick={() => setConfirmOpen(false)} />
        </DialogFooter>
      </Dialog>
    </div>
  );
};
