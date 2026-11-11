import { ISPCollection, SharePointService, SPEntity } from './SharePointService';
import { LISTS, WORK_ITEM_FIELDS } from './ProvisioningService';
import {
  DEFAULT_ELIGIBLE_STATUSES, IArchiveJob, IArchivePolicy, IArchiveProgress, IWorkspaceSummary, MAX_BATCH_SIZE
} from '../models/Models';

/** Fields copied verbatim from the active item to the archive item. */
const COPY_FIELDS = ['Title'].concat(WORK_ITEM_FIELDS.map(f => f.name));
const SELECT = ['Id', 'Created', 'Modified', 'Attachments'].concat(COPY_FIELDS).join(',');
/** Fields compared on read-back before the source is deleted. */
const VERIFY_FIELDS = ['Title', 'WorkStatus', 'RecordType', 'DueDate', 'CompletedDate', 'WorkDescription', 'AssignedToKey'];

/** A "Running" job that has not checkpointed for this long is considered abandoned (browser closed). */
const STALE_JOB_MINUTES = 30;
const TEXT_LIMIT = 255;

export interface IArchiveRunResult {
  jobId: number;
  archived: number;
  failed: number;
  skipped: number;
  processed: number;
}

interface IWorkItem extends SPEntity {
  Id: number;
  Attachments?: boolean;
  Created?: string;
  Modified?: string;
}

/**
 * Conservative archive engine.
 *
 * Sequence per item: copy (or reuse an existing copy found by OriginalItemId) → read back and
 * compare key fields → audit log → delete source. Items with attachments are skipped and kept.
 * Runs in the browser: the page must stay open (the UI warns before unload). Progress is
 * checkpointed to the job record after every page so an interrupted run is visible and
 * automatically released after STALE_JOB_MINUTES.
 */
export class ArchiveService {
  private cancelRequested: boolean = false;
  /** Items already counted in this run (the count phase can revisit items the date phase skipped). */
  private seen: Set<number> = new Set<number>();

  constructor(private readonly sp: SharePointService) {}

  public static validatePolicy(policy: IArchivePolicy): void {
    if (!policy || ['count', 'date', 'both'].indexOf(policy.Mode) < 0) {
      throw new Error('Select a valid archive method.');
    }
    if ((policy.Mode === 'count' || policy.Mode === 'both') &&
        (!Number.isFinite(Number(policy.MaxActiveItems)) || Number(policy.MaxActiveItems) < 1)) {
      throw new Error('Maximum active items must be a whole number of 1 or more.');
    }
    if (policy.Mode === 'date' || policy.Mode === 'both') {
      const cutoff = policy.ArchiveBefore ? new Date(policy.ArchiveBefore) : undefined;
      if (!cutoff || isNaN(cutoff.getTime())) throw new Error('Enter a valid archive cutoff date.');
      if (cutoff.getTime() > Date.now()) throw new Error('The archive cutoff date cannot be in the future.');
    }
    const batch = Number(policy.BatchSize);
    if (!Number.isFinite(batch) || batch < 1 || batch > MAX_BATCH_SIZE) {
      throw new Error(`Batch size must be between 1 and ${MAX_BATCH_SIZE}.`);
    }
  }

  /** Request a graceful stop after the current item. */
  public cancel(): void {
    this.cancelRequested = true;
  }

  public async getRecentJobs(workspaceId: string, top: number = 20): Promise<IArchiveJob[]> {
    const r = await this.sp.getCollection<SPEntity>(
      `${this.sp.listPath(LISTS.jobs)}/items?$select=Id,Title,WorkspaceId,JobStatus,ArchiveMode,Started,Finished,Modified,Processed,Archived,Skipped,FailedCount,LastError` +
      `&$filter=${encodeURIComponent(`WorkspaceId eq '${SharePointService.odataString(workspaceId)}'`)}&$orderby=Id desc&$top=${top}`
    );
    return r.value.map(j => ({
      Id: j.Id as number,
      Title: j.Title as string,
      WorkspaceId: j.WorkspaceId as string,
      Status: j.JobStatus as IArchiveJob['Status'],
      Mode: j.ArchiveMode as string,
      Started: j.Started as string | undefined,
      Finished: j.Finished as string | undefined,
      Modified: j.Modified as string | undefined,
      Processed: j.Processed as number | undefined,
      Archived: j.Archived as number | undefined,
      Skipped: j.Skipped as number | undefined,
      FailedCount: j.FailedCount as number | undefined,
      LastError: j.LastError as string | undefined
    }));
  }

  public async run(
    workspace: IWorkspaceSummary,
    policy: IArchivePolicy,
    onProgress: (progress: IArchiveProgress) => void
  ): Promise<IArchiveRunResult> {
    ArchiveService.validatePolicy(policy);
    this.cancelRequested = false;
    this.seen = new Set<number>();
    await this.releaseOrRejectRunningJob(workspace.id);

    const job = await this.sp.createItem(LISTS.jobs, {
      Title: `Archive ${workspace.title} ${new Date().toISOString()}`,
      WorkspaceId: workspace.id,
      JobStatus: 'Running',
      ArchiveMode: policy.Mode,
      Processed: 0, Archived: 0, Skipped: 0, FailedCount: 0,
      Started: new Date().toISOString()
    });

    const result: IArchiveRunResult = { jobId: job.Id, archived: 0, failed: 0, skipped: 0, processed: 0 };
    const report = (message: string): void => onProgress({ ...result, message });
    const pageSize = Math.min(Math.max(Math.floor(Number(policy.BatchSize)) || 100, 1), MAX_BATCH_SIZE);
    const statuses = (policy.EligibleStatuses || DEFAULT_ELIGIBLE_STATUSES).split(';').map(s => s.trim()).filter(Boolean);
    const statusFilter = statuses.map(s => `WorkStatus eq '${SharePointService.odataString(s)}'`).join(' or ');
    const baseClauses = ["RecordType eq 'Task'", `(${statusFilter})`];

    try {
      // Phase 1 (date / both): everything eligible completed before the cutoff.
      if (policy.Mode === 'date' || policy.Mode === 'both') {
        const cutoff = new Date(policy.ArchiveBefore as string).toISOString();
        report('Archiving eligible items completed before the cutoff date…');
        await this.processQuery(workspace, job.Id, baseClauses.concat(`CompletedDate lt datetime'${cutoff}'`),
          pageSize, undefined, result, report);
      }

      // Phase 2 (count / both): if still above the maximum, archive the oldest remaining eligible items.
      // The count is re-read after phase 1 so it reflects what is actually left in the list.
      if (!this.cancelRequested && (policy.Mode === 'count' || policy.Mode === 'both')) {
        const currentCount = await this.sp.getListItemCount(workspace.workItemsList);
        const target = Math.max(0, currentCount - Math.floor(Number(policy.MaxActiveItems)));
        if (target > 0) {
          report(`Active list holds ${currentCount.toLocaleString()} items; archiving up to ${target.toLocaleString()} of the oldest eligible items…`);
          await this.processQuery(workspace, job.Id, baseClauses, pageSize, target, result, report);
        }
      }

      const status = this.cancelRequested ? 'Abandoned' : result.failed ? 'CompletedWithErrors' : 'Completed';
      await this.checkpoint(job.Id, result, {
        JobStatus: status,
        Finished: new Date().toISOString(),
        LastError: this.cancelRequested ? 'Stopped by user.' : null
      });
      report(this.cancelRequested ? 'Archive run stopped by user.' : 'Archive run finished.');
      return result;
    } catch (e) {
      try {
        await this.checkpoint(job.Id, result, {
          JobStatus: 'Failed',
          LastError: SharePointService.errorMessage(e, 'Unknown error').slice(0, 4000),
          Finished: new Date().toISOString()
        });
      } catch {
        // The original error is more useful to the user than a failure to record it.
      }
      throw e;
    }
  }

  private async releaseOrRejectRunningJob(workspaceId: string): Promise<void> {
    const running = await this.sp.getCollection<{ Id: number; Modified: string }>(
      `${this.sp.listPath(LISTS.jobs)}/items?$select=Id,Modified&$filter=${encodeURIComponent(
        `WorkspaceId eq '${SharePointService.odataString(workspaceId)}' and JobStatus eq 'Running'`
      )}&$top=10`
    );
    for (const job of running.value) {
      const idleMinutes = (Date.now() - new Date(job.Modified).getTime()) / 60000;
      if (idleMinutes < STALE_JOB_MINUTES) {
        throw new Error(
          'Another archive run for this workspace is in progress (it reported progress in the last ' +
          `${STALE_JOB_MINUTES} minutes). Wait for it to finish, or try again later.`
        );
      }
      await this.sp.updateItem(LISTS.jobs, job.Id, {
        JobStatus: 'Abandoned',
        LastError: `No progress for ${Math.round(idleMinutes)} minutes; the browser running it was probably closed.`,
        Finished: new Date().toISOString()
      });
    }
  }

  private checkpoint(jobId: number, result: IArchiveRunResult, extra?: SPEntity): Promise<void> {
    return this.sp.updateItem(LISTS.jobs, jobId, {
      Processed: result.processed,
      Archived: result.archived,
      Skipped: result.skipped,
      FailedCount: result.failed,
      ...(extra || {})
    });
  }

  private async processQuery(
    workspace: IWorkspaceSummary,
    jobId: number,
    clauses: string[],
    pageSize: number,
    successTarget: number | undefined,
    result: IArchiveRunResult,
    report: (message: string) => void
  ): Promise<void> {
    let next: string | undefined =
      `${this.sp.listPath(workspace.workItemsList)}/items?$select=${SELECT}` +
      `&$filter=${encodeURIComponent(clauses.join(' and '))}&$orderby=CompletedDate asc,Id asc&$top=${pageSize}`;
    let successes = 0;
    const done = (): boolean => this.cancelRequested || (successTarget !== undefined && successes >= successTarget);

    while (next && !done()) {
      const page: ISPCollection<IWorkItem> = await this.sp.getCollection<IWorkItem>(next);
      if (!page.value.length) break;

      for (const item of page.value) {
        if (done()) break;
        const firstVisit = !this.seen.has(item.Id);
        this.seen.add(item.Id);
        if (firstVisit) result.processed++;

        if (item.Attachments === true) {
          // Attachment transfer is not part of this release: never delete a source that has attachments.
          if (firstVisit) {
            result.skipped++;
            await this.log(workspace.id, jobId, item.Id, 'Skipped', 'Item has attachments; attachment transfer is not supported in this release.');
          }
        } else if (await this.archiveOne(workspace, jobId, item)) {
          result.archived++;
          successes++;
        } else {
          result.failed++;
        }
        report(`Processed ${result.processed.toLocaleString()} · archived ${result.archived.toLocaleString()} · ` +
          `skipped ${result.skipped.toLocaleString()} · failed ${result.failed.toLocaleString()}`);
      }

      await this.checkpoint(jobId, result);
      // Key-based ($skiptoken) continuation, so deleting rows on this page does not skip any rows.
      next = page.nextLink;
    }
  }

  /** @returns true when the item was copied, verified, audited and the source deleted. */
  private async archiveOne(workspace: IWorkspaceSummary, jobId: number, item: IWorkItem): Promise<boolean> {
    try {
      const existing = await this.sp.getCollection<{ Id: number }>(
        `${this.sp.listPath(workspace.archiveList)}/items?$select=Id&$filter=${encodeURIComponent(`OriginalItemId eq ${item.Id}`)}&$top=1`
      );
      let archiveId = existing.value.length ? existing.value[0].Id : 0;

      if (!archiveId) {
        const payload: SPEntity = {};
        COPY_FIELDS.forEach(field => {
          if (item[field] !== undefined && item[field] !== null) payload[field] = item[field];
        });
        payload.Title = item.Title || '(untitled)';
        payload.OriginalItemId = item.Id;
        payload.OriginalCreated = item.Created;
        payload.OriginalModified = item.Modified;
        payload.ArchivedAt = new Date().toISOString();
        payload.ArchiveJobId = jobId;
        archiveId = (await this.sp.createItem(workspace.archiveList, payload)).Id;
      }

      const verified = await this.sp.get<SPEntity>(
        `${this.sp.listPath(workspace.archiveList)}/items(${archiveId})?$select=Id,OriginalItemId,${VERIFY_FIELDS.join(',')}`
      );
      ArchiveService.verifyCopy(item, verified);

      await this.log(workspace.id, jobId, item.Id, 'Archived', undefined, archiveId, true);
      // Source deletion is the final step and only follows a verified copy and a successful audit write.
      await this.sp.deleteItem(workspace.workItemsList, item.Id);
      return true;
    } catch (e) {
      await this.log(workspace.id, jobId, item.Id, 'Failed', SharePointService.errorMessage(e, 'Unknown error'));
      return false;
    }
  }

  private async log(
    workspaceId: string, jobId: number, originalId: number, outcome: string,
    detail?: string, archiveItemId?: number, mustSucceed: boolean = false
  ): Promise<void> {
    try {
      await this.sp.createItem(LISTS.log, {
        Title: `${outcome} ${workspaceId} item ${originalId}`.slice(0, TEXT_LIMIT),
        WorkspaceId: workspaceId,
        ArchiveJobId: jobId,
        OriginalItemId: originalId,
        ArchiveItemId: archiveItemId,
        Outcome: outcome,
        Detail: detail ? detail.slice(0, 4000) : undefined
      });
    } catch (e) {
      // An "Archived" audit entry is a precondition for deleting the source; other entries are best effort.
      if (mustSucceed) throw e;
    }
  }

  private static verifyCopy(source: SPEntity, archived: SPEntity): void {
    if (!archived || Number(archived.OriginalItemId) !== Number(source.Id)) {
      throw new Error('Archive verification failed: original item ID does not match.');
    }
    for (const field of VERIFY_FIELDS) {
      const normalise = (v: unknown): string => (v === undefined || v === null ? '' : String(v));
      const left = field === 'Title' ? normalise(source.Title) || '(untitled)' : normalise(source[field]);
      if (left !== normalise(archived[field])) {
        throw new Error(`Archive verification failed: ${field} does not match.`);
      }
    }
  }
}
