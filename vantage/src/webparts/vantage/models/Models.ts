export type ArchiveMode = 'count' | 'date' | 'both';

export type ArchiveJobStatus = 'Running' | 'Completed' | 'CompletedWithErrors' | 'Failed' | 'Abandoned';

export interface IArchivePolicy {
  Id?: number;
  Title?: string;
  WorkspaceId: string;
  Enabled: boolean;
  Mode: ArchiveMode;
  MaxActiveItems?: number;
  /** ISO date-time (UTC). Items completed strictly before this instant are eligible. */
  ArchiveBefore?: string;
  /** Semicolon separated list of terminal statuses. */
  EligibleStatuses: string;
  BatchSize: number;
}

export interface IArchiveJob {
  Id: number;
  Title: string;
  WorkspaceId: string;
  Status: ArchiveJobStatus;
  Mode: string;
  Started?: string;
  Finished?: string;
  Modified?: string;
  Processed?: number;
  Archived?: number;
  Skipped?: number;
  FailedCount?: number;
  LastError?: string;
}

export interface IWorkspaceSummary {
  id: string;
  title: string;
  workItemsList: string;
  archiveList: string;
}

export interface IArchiveProgress {
  processed: number;
  archived: number;
  failed: number;
  skipped: number;
  message: string;
}

export const DEFAULT_ELIGIBLE_STATUSES = 'Completed;Cancelled';
export const MAX_BATCH_SIZE = 200;
