import { IAttachment, ISPCollection, SharePointError, SharePointService, SPEntity } from './SharePointService';
import { LISTS, WORK_ITEM_FIELDS } from './ProvisioningService';
import {
  DEFAULT_ELIGIBLE_STATUSES, IArchivedTask, IArchiveJob, IArchiveLogEntry, IArchivePolicy, IArchiveProgress,
  IWorkspaceSummary, MAX_ATTACHMENT_BYTES, MAX_BATCH_SIZE, TERMINAL_STATUSES
} from '../models/Models';

/** Fields copied verbatim between the active and archive lists (and verified on read-back). */
const COPY_FIELDS = ['Title'].concat(WORK_ITEM_FIELDS.map(f => f.name));
const BOOLEAN_FIELDS = new Set(WORK_ITEM_FIELDS.filter(f => f.kind === 'Boolean').map(f => f.name));
const SELECT = ['Id', 'Created', 'Modified', 'Attachments'].concat(COPY_FIELDS).join(',');

/** A lock whose heartbeat is older than this is considered abandoned (browser closed). */
const STALE_LOCK_MINUTES = 30;
const TEXT_LIMIT = 255;

export interface IArchiveRunResult {
  jobId: number;
  /** Root tasks archived. */
  archived: number;
  /** Total records moved (tasks plus their comments, checklist items, dependencies, activity and notifications). */
  records: number;
  failed: number;
  skipped: number;
  processed: number;
  /** Items still above the configured maximum after the run (count/both modes). */
  remainingOverLimit: number;
}

interface IWorkItem extends SPEntity {
  Id: number;
  Attachments?: boolean;
  Created?: string;
  Modified?: string;
}

/** A task that is deliberately kept (not an error), e.g. because an attachment is too large. */
class SkipError extends Error {
  constructor(message: string) {
    super(message);
    Object.setPrototypeOf(this, SkipError.prototype);
  }
}

const q = SharePointService.odataString;
const nowEpoch = (): number => Math.floor(Date.now() / 1000);

/**
 * Archive engine. For each eligible task: copy the task and every related record (with
 * attachments) to the archive list → read back and verify every copied field and the SHA-256 of
 * every attachment → confirm the source did not change meanwhile → write the audit entry → delete
 * the related records and then the task. Any failure leaves the source in place; a re-run reuses
 * existing archive copies (matched on the unique OriginalItemId) instead of duplicating them.
 *
 * Runs in the browser: the page must stay open. A per-workspace lock row (unique WorkspaceId)
 * prevents concurrent runs; its heartbeat is refreshed per page and a stale lock is reclaimed.
 */
export class ArchiveService {
  private cancelRequested: boolean = false;
  private seen: Set<number> = new Set<number>();

  constructor(private readonly sp: SharePointService) {}

  public static validatePolicy(policy: IArchivePolicy): void {
    if (!policy || ['count', 'date', 'both'].indexOf(policy.Mode) < 0) throw new Error('Select a valid archive method.');
    const statuses = (policy.EligibleStatuses || DEFAULT_ELIGIBLE_STATUSES).split(';').map(s => s.trim()).filter(Boolean);
    if (!statuses.length || statuses.some(s => TERMINAL_STATUSES.indexOf(s) < 0)) {
      throw new Error('Only Completed and Cancelled tasks can be archived.');
    }
    if ((policy.Mode === 'count' || policy.Mode === 'both') &&
        (!Number.isInteger(Number(policy.MaxActiveItems)) || Number(policy.MaxActiveItems) < 1)) {
      throw new Error('Maximum active items must be a whole number of 1 or more.');
    }
    if (policy.Mode === 'date' || policy.Mode === 'both') {
      const cutoff = policy.ArchiveBefore ? new Date(policy.ArchiveBefore) : undefined;
      if (!cutoff || isNaN(cutoff.getTime())) throw new Error('Enter a valid archive cutoff date.');
      if (cutoff.getTime() > Date.now()) throw new Error('The archive cutoff date cannot be in the future.');
    }
    const batch = Number(policy.BatchSize);
    if (!Number.isInteger(batch) || batch < 1 || batch > MAX_BATCH_SIZE) throw new Error(`Batch size must be between 1 and ${MAX_BATCH_SIZE}.`);
  }

  /** Request a graceful stop after the current task. */
  public cancel(): void {
    this.cancelRequested = true;
  }

  // ---------------------------------------------------------------------------
  // Read models for the UI
  // ---------------------------------------------------------------------------

  public async getRecentJobs(workspaceId: string, top: number = 20): Promise<IArchiveJob[]> {
    const r = await this.sp.getCollection<SPEntity>(
      `${this.sp.listPath(LISTS.jobs)}/items?$select=Id,Title,WorkspaceId,JobStatus,ArchiveMode,Started,Finished,Modified,Processed,Archived,Records,Skipped,FailedCount,LastError` +
      `&${SharePointService.filter(`WorkspaceId eq '${q(workspaceId)}'`)}&$orderby=Id desc&$top=${top}`
    );
    return r.value.map(j => ({
      Id: j.Id as number, Title: j.Title as string, WorkspaceId: j.WorkspaceId as string,
      Status: j.JobStatus as IArchiveJob['Status'], Mode: j.ArchiveMode as string,
      Started: j.Started as string | undefined, Finished: j.Finished as string | undefined, Modified: j.Modified as string | undefined,
      Processed: j.Processed as number | undefined, Archived: j.Archived as number | undefined, Records: j.Records as number | undefined,
      Skipped: j.Skipped as number | undefined, FailedCount: j.FailedCount as number | undefined, LastError: j.LastError as string | undefined
    }));
  }

  public async getArchivedTasks(ws: IWorkspaceSummary, top: number = 100): Promise<IArchivedTask[]> {
    const r = await this.sp.getCollection<IArchivedTask>(
      `${this.sp.listPath(ws.archiveList)}/items?$select=Id,Title,WorkStatus,AssignedToName,OriginalItemId,ArchivedAt,CompletedDate` +
      `&${SharePointService.filter("RecordType eq 'Task'")}&$orderby=Id desc&$top=${top}`
    );
    return r.value;
  }

  public async getLog(workspaceId: string, top: number = 50): Promise<IArchiveLogEntry[]> {
    const r = await this.sp.getCollection<IArchiveLogEntry>(
      `${this.sp.listPath(LISTS.log)}/items?$select=Id,Title,Outcome,Detail,OriginalItemId,Created` +
      `&${SharePointService.filter(`WorkspaceId eq '${q(workspaceId)}'`)}&$orderby=Id desc&$top=${top}`
    );
    return r.value;
  }

  // ---------------------------------------------------------------------------
  // Archive run
  // ---------------------------------------------------------------------------

  public async run(ws: IWorkspaceSummary, policy: IArchivePolicy, owner: string, onProgress: (p: IArchiveProgress) => void): Promise<IArchiveRunResult> {
    ArchiveService.validatePolicy(policy);
    if (!policy.Enabled) throw new Error('The archive policy is disabled.');
    this.cancelRequested = false;
    this.seen = new Set<number>();

    const lockId = await this.acquireLock(ws.id, owner);
    const result: IArchiveRunResult = { jobId: 0, archived: 0, records: 0, failed: 0, skipped: 0, processed: 0, remainingOverLimit: 0 };
    const report = (message: string): void => onProgress({ ...result, message });
    try {
      const job = await this.sp.createItem(LISTS.jobs, {
        Title: `Archive ${ws.title} ${new Date().toISOString()}`.slice(0, TEXT_LIMIT), WorkspaceId: ws.id, JobStatus: 'Running',
        ArchiveMode: policy.Mode, Processed: 0, Archived: 0, Records: 0, Skipped: 0, FailedCount: 0, Started: new Date().toISOString()
      });
      result.jobId = job.Id;
      await this.sp.updateItem(LISTS.locks, lockId, { JobId: job.Id });

      const pageSize = Math.min(Math.max(Math.floor(Number(policy.BatchSize)) || 100, 1), MAX_BATCH_SIZE);
      const statuses = (policy.EligibleStatuses || DEFAULT_ELIGIBLE_STATUSES).split(';').map(s => s.trim()).filter(Boolean);
      const baseClauses = ["RecordType eq 'Task'", `WorkspaceId eq '${q(ws.id)}'`, `(${statuses.map(s => `WorkStatus eq '${q(s)}'`).join(' or ')})`];
      const max = Math.floor(Number(policy.MaxActiveItems));

      try {
        if (policy.Mode === 'date' || policy.Mode === 'both') {
          const cutoff = new Date(policy.ArchiveBefore as string).toISOString();
          report('Archiving eligible tasks completed before the cutoff date…');
          await this.processQuery(ws, job.Id, lockId, baseClauses.concat(`CompletedDate lt datetime'${cutoff}'`), pageSize, undefined, result, report);
        }
        if (!this.cancelRequested && (policy.Mode === 'count' || policy.Mode === 'both')) {
          // ItemCount includes related records, so the target is measured in records moved.
          const count = await this.sp.getListItemCount(ws.workItemsList);
          const target = Math.max(0, count - max);
          if (target > 0) {
            report(`Active list holds ${count.toLocaleString()} records; archiving the oldest eligible tasks to remove at least ${target.toLocaleString()}…`);
            await this.processQuery(ws, job.Id, lockId, baseClauses, pageSize, result.records + target, result, report);
          }
          const remaining = await this.sp.getListItemCount(ws.workItemsList);
          // eslint-disable-next-line require-atomic-updates -- sequential; nothing else writes `result` concurrently
          result.remainingOverLimit = Math.max(0, remaining - max);
        }

        const status = this.cancelRequested ? 'Abandoned' : (result.failed || result.remainingOverLimit) ? 'CompletedWithErrors' : 'Completed';
        const note = this.cancelRequested ? 'Stopped by user.'
          : result.failed ? `${result.failed} task(s) failed and were kept; see the archive log.`
            : result.remainingOverLimit ? `${result.remainingOverLimit.toLocaleString()} records remain above the maximum (not enough eligible tasks).` : null;
        await this.checkpoint(job.Id, result, { JobStatus: status, Finished: new Date().toISOString(), LastError: note });
        report(this.cancelRequested ? 'Archive run stopped.' : 'Archive run finished.');
        return result;
      } catch (e) {
        try {
          await this.checkpoint(job.Id, result, { JobStatus: 'Failed', LastError: SharePointService.errorMessage(e, 'Unknown error').slice(0, 4000), Finished: new Date().toISOString() });
        } catch {
          // The original error is more useful to the user than a failure to record it.
        }
        throw e;
      }
    } finally {
      await this.releaseLock(lockId);
    }
  }

  private async processQuery(
    ws: IWorkspaceSummary, jobId: number, lockId: number, clauses: string[], pageSize: number,
    recordTarget: number | undefined, result: IArchiveRunResult, report: (message: string) => void
  ): Promise<void> {
    let next: string | undefined =
      `${this.sp.listPath(ws.workItemsList)}/items?$select=${SELECT}&${SharePointService.filter(clauses.join(' and '))}` +
      `&$orderby=CompletedDate asc,Id asc&$top=${pageSize}`;
    const done = (): boolean => this.cancelRequested || (recordTarget !== undefined && result.records >= recordTarget);

    while (next && !done()) {
      const page: ISPCollection<IWorkItem> = await this.sp.getCollection<IWorkItem>(next);
      if (!page.value.length) break;
      for (const item of page.value) {
        if (done()) break;
        if (this.seen.has(item.Id)) continue; // Already handled (and kept) in the date phase.
        this.seen.add(item.Id);
        result.processed++;
        try {
          const moved = await this.archiveTask(ws, jobId, item);
          result.archived++;
          result.records += moved;
        } catch (e) {
          if (e instanceof SkipError) result.skipped++;
          else result.failed++;
          await this.log(ws.id, jobId, item.Id, e instanceof SkipError ? 'Skipped' : 'Failed', SharePointService.errorMessage(e, 'Unknown error'));
        }
        report(`Processed ${result.processed.toLocaleString()} tasks · archived ${result.archived.toLocaleString()} (${result.records.toLocaleString()} records) · ` +
          `skipped ${result.skipped.toLocaleString()} · failed ${result.failed.toLocaleString()}`);
      }
      await this.checkpoint(jobId, result);
      await this.heartbeat(lockId);
      // Key-based ($skiptoken) continuation, so deleting rows on this page does not skip any rows.
      next = page.nextLink;
    }
  }

  /** @returns number of records moved (task + related records). Throws to keep the source. */
  private async archiveTask(ws: IWorkspaceSummary, jobId: number, task: IWorkItem): Promise<number> {
    const children = await this.related(ws.workItemsList, `ParentTaskId eq ${task.Id}`);
    const now = new Date().toISOString();
    const lineage = (src: IWorkItem): SPEntity => ({
      OriginalItemId: src.Id, OriginalCreated: src.Created, OriginalModified: src.Modified, ArchivedAt: now, ArchiveJobId: jobId
    });

    // 0. Check attachment sizes before writing anything, so a skipped task leaves no partial archive copy.
    for (const src of [task].concat(children)) {
      if (src.Attachments === false) continue;
      for (const file of await this.sp.listAttachments(ws.workItemsList, src.Id)) await this.checkSize(file);
    }

    // 1. Copy and verify everything first; nothing is deleted until all copies are verified.
    let attachments = 0;
    for (const src of [task].concat(children)) {
      const existing = await this.sp.getCollection<{ Id: number }>(
        `${this.sp.listPath(ws.archiveList)}/items?$select=Id&${SharePointService.filter(`OriginalItemId eq ${src.Id}`)}&$top=1`
      );
      const targetId = existing.value.length
        ? existing.value[0].Id
        : (await this.sp.createItem(ws.archiveList, { ...ArchiveService.payload(src), ...lineage(src) })).Id;
      if (existing.value.length) await this.sp.updateItem(ws.archiveList, targetId, ArchiveService.payload(src));
      attachments += await this.copyAttachments(ws.workItemsList, src, ws.archiveList, targetId);
      const copy = await this.sp.get<SPEntity>(`${this.sp.listPath(ws.archiveList)}/items(${targetId})?$select=Id,OriginalItemId,${COPY_FIELDS.join(',')}`);
      if (Number(copy.OriginalItemId) !== src.Id) throw new Error(`Archive verification failed for record ${src.Id}: original ID does not match.`);
      ArchiveService.verifyFields(src, copy, `record ${src.Id}`);
    }

    // 2. Make sure nobody changed the task or added related records while we were copying.
    const current = await this.sp.get<IWorkItem>(`${this.sp.listPath(ws.workItemsList)}/items(${task.Id})?$select=Id,Modified`);
    if (String(current.Modified) !== String(task.Modified)) throw new Error('The task was changed while it was being archived; it was kept for the next run.');
    const childrenNow = await this.related(ws.workItemsList, `ParentTaskId eq ${task.Id}`);
    if (childrenNow.length !== children.length) throw new Error('Related records were added while the task was being archived; it was kept for the next run.');

    // 3. Audit, then delete related records and finally the task.
    await this.log(ws.id, jobId, task.Id, 'Archived', `${children.length} related record(s), ${attachments} attachment(s).`, true);
    for (const c of children) await this.sp.deleteItem(ws.workItemsList, c.Id);
    await this.sp.deleteItem(ws.workItemsList, task.Id);
    return 1 + children.length;
  }

  // ---------------------------------------------------------------------------
  // Restore
  // ---------------------------------------------------------------------------

  /** Restores an archived task and its related records. Returns the new active task ID. */
  public async restore(ws: IWorkspaceSummary, archiveId: number): Promise<number> {
    const locks = await this.sp.getCollection<{ Id: number }>(
      `${this.sp.listPath(LISTS.locks)}/items?$select=Id&${SharePointService.filter(`WorkspaceId eq '${q(ws.id)}'`)}&$top=1`
    );
    if (locks.value.length) throw new Error('An archive run is in progress for this workspace. Restore when it has finished.');

    const root = await this.sp.get<IWorkItem>(`${this.sp.listPath(ws.archiveList)}/items(${archiveId})?$select=Id,OriginalItemId,Attachments,${COPY_FIELDS.join(',')}`);
    if (!root.OriginalItemId || root.RecordType !== 'Task') throw new Error('This archive record is not a task and cannot be restored on its own.');
    const stillActive = await this.sp.getCollection<{ Id: number }>(
      `${this.sp.listPath(ws.workItemsList)}/items?$select=Id&${SharePointService.filter(`Id eq ${Number(root.OriginalItemId)}`)}&$top=1`
    );
    if (stillActive.value.length) {
      throw new Error('The original task is still in the active list (an earlier archive run did not finish for it). Run the archive again to complete it instead of restoring.');
    }
    const children = await this.related(ws.archiveList, `ParentTaskId eq ${Number(root.OriginalItemId)}`, ',OriginalItemId');

    const restoreOne = async (src: IWorkItem, overrides: SPEntity): Promise<number> => {
      const existing = await this.sp.getCollection<{ Id: number }>(
        `${this.sp.listPath(ws.workItemsList)}/items?$select=Id&${SharePointService.filter(`RestoredFromArchiveId eq ${src.Id}`)}&$top=1`
      );
      const targetId = existing.value.length
        ? existing.value[0].Id
        : (await this.sp.createItem(ws.workItemsList, { ...ArchiveService.payload(src), ...overrides, RestoredFromArchiveId: src.Id })).Id;
      await this.copyAttachments(ws.archiveList, src, ws.workItemsList, targetId);
      const copy = await this.sp.get<SPEntity>(`${this.sp.listPath(ws.workItemsList)}/items(${targetId})?$select=Id,RestoredFromArchiveId,${COPY_FIELDS.join(',')}`);
      if (Number(copy.RestoredFromArchiveId) !== src.Id) throw new Error(`Restore verification failed for archive record ${src.Id}.`);
      ArchiveService.verifyFields({ ...src, ...overrides }, copy, `archive record ${src.Id}`);
      return targetId;
    };

    const newTaskId = await restoreOne(root, {});
    for (const c of children) await restoreOne(c, { ParentTaskId: newTaskId });

    await this.log(ws.id, 0, Number(root.OriginalItemId), 'Restored', `Restored as task ${newTaskId} with ${children.length} related record(s).`, true);
    for (const c of children) await this.sp.deleteItem(ws.archiveList, c.Id);
    await this.sp.deleteItem(ws.archiveList, root.Id);
    return newTaskId;
  }

  // ---------------------------------------------------------------------------
  // Attachments
  // ---------------------------------------------------------------------------

  /** Copies attachments so the target matches the source exactly, verified by SHA-256. Returns the file count. */
  private async copyAttachments(sourceList: string, src: IWorkItem, targetList: string, targetId: number): Promise<number> {
    const sourceFiles = src.Attachments === false ? [] : await this.sp.listAttachments(sourceList, src.Id);
    const targetFiles = await this.sp.listAttachments(targetList, targetId);
    for (const t of targetFiles) {
      if (!sourceFiles.some(s => s.FileName === t.FileName)) await this.sp.deleteAttachment(targetList, targetId, t.FileName);
    }
    for (const file of sourceFiles) {
      await this.checkSize(file);
      const content = await this.sp.getAttachmentContent(file);
      const sourceHash = await ArchiveService.hash(content);
      const existing = targetFiles.find(t => t.FileName === file.FileName);
      if (existing && await ArchiveService.hash(await this.sp.getAttachmentContent(existing)) === sourceHash) continue;
      if (existing) await this.sp.deleteAttachment(targetList, targetId, file.FileName);
      await this.sp.addAttachment(targetList, targetId, file.FileName, content);

      const copied = (await this.sp.listAttachments(targetList, targetId)).find(t => t.FileName === file.FileName) as IAttachment | undefined;
      if (!copied || await ArchiveService.hash(await this.sp.getAttachmentContent(copied)) !== sourceHash) {
        throw new Error(`Attachment verification failed for “${file.FileName}”.`);
      }
    }
    return sourceFiles.length;
  }

  private async checkSize(file: IAttachment): Promise<void> {
    const size = (await this.sp.get<{ Length?: string | number }>(
      `web/getfilebyserverrelativeurl('${encodeURIComponent(q(file.ServerRelativeUrl))}')?$select=Length`
    )).Length;
    if (Number(size) > MAX_ATTACHMENT_BYTES) {
      throw new SkipError(`Attachment “${file.FileName}” is larger than ${MAX_ATTACHMENT_BYTES / 1048576} MB; the task was kept in the active list.`);
    }
  }

  private static async hash(content: ArrayBuffer): Promise<string> {
    const subtle = typeof window !== 'undefined' && window.crypto && window.crypto.subtle;
    const bytes = new Uint8Array(subtle ? await subtle.digest('SHA-256', content) : content);
    let out = String(content.byteLength) + ':';
    for (let i = 0; i < bytes.length; i++) out += ('0' + bytes[i].toString(16)).slice(-2);
    return out;
  }

  // ---------------------------------------------------------------------------
  // Locking, checkpoints and logging
  // ---------------------------------------------------------------------------

  private async acquireLock(workspaceId: string, owner: string): Promise<number> {
    const create = (): Promise<{ Id: number }> => this.sp.createItem(LISTS.locks, {
      Title: `Archive lock ${workspaceId}`, WorkspaceId: workspaceId, HeartbeatEpoch: nowEpoch(), LockOwner: owner.slice(0, TEXT_LIMIT)
    });
    const existing = await this.sp.getCollection<{ Id: number; HeartbeatEpoch?: number; JobId?: number; LockOwner?: string }>(
      `${this.sp.listPath(LISTS.locks)}/items?$select=Id,HeartbeatEpoch,JobId,LockOwner&${SharePointService.filter(`WorkspaceId eq '${q(workspaceId)}'`)}&$top=1`
    );
    if (existing.value.length) {
      const lock = existing.value[0];
      const idleMinutes = (nowEpoch() - Number(lock.HeartbeatEpoch || 0)) / 60;
      if (idleMinutes < STALE_LOCK_MINUTES) {
        throw new Error(`An archive run started by ${lock.LockOwner || 'another user'} is in progress for this workspace. Try again when it has finished.`);
      }
      if (lock.JobId) {
        await this.sp.updateItem(LISTS.jobs, lock.JobId, {
          JobStatus: 'Abandoned', Finished: new Date().toISOString(),
          LastError: `No progress for ${Math.round(idleMinutes)} minutes; the browser running it was probably closed.`
        }).catch(() => undefined);
      }
      await this.sp.deleteItem(LISTS.locks, lock.Id);
    }
    try {
      return (await create()).Id;
    } catch (e) {
      // The unique WorkspaceId constraint rejected the insert: someone else acquired it first.
      if (e instanceof SharePointError && (e.status === 400 || e.status === 409)) {
        throw new Error('Another archive run for this workspace started at the same moment. Try again shortly.');
      }
      throw e;
    }
  }

  private heartbeat(lockId: number): Promise<void> {
    return this.sp.updateItem(LISTS.locks, lockId, { HeartbeatEpoch: nowEpoch() });
  }

  private async releaseLock(lockId: number): Promise<void> {
    try {
      await this.sp.deleteItem(LISTS.locks, lockId);
    } catch {
      // A stale lock is reclaimed automatically after STALE_LOCK_MINUTES.
    }
  }

  private checkpoint(jobId: number, r: IArchiveRunResult, extra?: SPEntity): Promise<void> {
    return this.sp.updateItem(LISTS.jobs, jobId, {
      Processed: r.processed, Archived: r.archived, Records: r.records, Skipped: r.skipped, FailedCount: r.failed, ...(extra || {})
    });
  }

  private async log(workspaceId: string, jobId: number, originalId: number, outcome: string, detail?: string, mustSucceed: boolean = false): Promise<void> {
    try {
      await this.sp.createItem(LISTS.log, {
        Title: `${outcome} ${workspaceId} item ${originalId}`.slice(0, TEXT_LIMIT), WorkspaceId: workspaceId,
        ArchiveJobId: jobId || undefined, OriginalItemId: originalId, Outcome: outcome, Detail: detail ? detail.slice(0, 4000) : undefined
      });
    } catch (e) {
      // An "Archived"/"Restored" audit entry is a precondition for deleting the source; other entries are best effort.
      if (mustSucceed) throw e;
    }
  }

  private async related(list: string, filter: string, extraSelect: string = ''): Promise<IWorkItem[]> {
    return this.sp.getAll<IWorkItem>(`${this.sp.listPath(list)}/items?$select=${SELECT}${extraSelect}&${SharePointService.filter(filter)}&$orderby=Id asc&$top=500`, 20);
  }

  private static payload(src: SPEntity): SPEntity {
    const out: SPEntity = {};
    COPY_FIELDS.forEach(f => { if (src[f] !== undefined && src[f] !== null) out[f] = src[f]; });
    return out;
  }

  private static verifyFields(expected: SPEntity, actual: SPEntity, label: string): void {
    const norm = (f: string, v: unknown): string =>
      BOOLEAN_FIELDS.has(f) ? (v ? '1' : '0') : v === undefined || v === null ? '' : String(v);
    for (const f of COPY_FIELDS) {
      if (norm(f, expected[f]) !== norm(f, actual[f])) throw new Error(`Verification failed for ${label}: ${f} does not match.`);
    }
  }
}

