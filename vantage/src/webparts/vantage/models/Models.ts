export type ArchiveMode = 'count' | 'date' | 'both';

export type ArchiveJobStatus = 'Running' | 'Completed' | 'CompletedWithErrors' | 'Failed' | 'Abandoned';

export const WORK_STATUSES = ['Not Started', 'In Progress', 'On Hold', 'In Review', 'Completed', 'Cancelled'] as const;
export type WorkStatus = typeof WORK_STATUSES[number];
export const TERMINAL_STATUSES: string[] = ['Completed', 'Cancelled'];

export const PRIORITIES = ['Low', 'Medium', 'High', 'Critical'] as const;
export type Priority = typeof PRIORITIES[number];

export type RecordType = 'Task' | 'Comment' | 'Checklist' | 'Dependency' | 'Activity' | 'Notification';

/** Effective role of the current user in a workspace. */
export type WorkspaceRole = 'owner' | 'admin' | 'member' | 'reader' | 'none';

export interface IArchivePolicy {
  Id?: number;
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
  Records?: number;
  Skipped?: number;
  FailedCount?: number;
  LastError?: string;
}

export interface IWorkspaceSummary {
  id: string;
  /** Registry list item ID. */
  itemId: number;
  title: string;
  workItemsList: string;
  archiveList: string;
  ownerKey?: string;
  secured: boolean;
}

export interface IBoardSummary {
  id: string;
  itemId: number;
  title: string;
  description?: string;
  workspaceId: string;
}

export interface ITask {
  Id: number;
  Title: string;
  WorkDescription?: string;
  BoardId: string;
  WorkspaceId: string;
  RecordType: RecordType;
  WorkStatus: WorkStatus;
  Priority?: Priority;
  AssignedToKey?: string;
  AssignedToName?: string;
  DueDate?: string;
  CompletedDate?: string;
  Progress?: number;
  Modified?: string;
  Created?: string;
}

export interface IRelatedRecord {
  Id: number;
  Title: string;
  RecordType: RecordType;
  WorkDescription?: string;
  ParentTaskId: number;
  RelatedTaskId?: number;
  ChecklistDone?: boolean;
  ActorKey?: string;
  ActorName?: string;
  Created?: string;
}

export interface INotification {
  Id: number;
  Title: string;
  WorkDescription?: string;
  NotificationRead?: boolean;
  ParentTaskId?: number;
  ActorName?: string;
  Created?: string;
}

export interface IArchivedTask {
  Id: number;
  Title: string;
  WorkStatus?: string;
  AssignedToName?: string;
  OriginalItemId: number;
  ArchivedAt?: string;
  CompletedDate?: string;
}

export interface IArchiveLogEntry {
  Id: number;
  Title: string;
  Outcome: string;
  Detail?: string;
  OriginalItemId?: number;
  Created?: string;
}

export interface IArchiveProgress {
  processed: number;
  archived: number;
  records: number;
  failed: number;
  skipped: number;
  message: string;
}

export interface IDashboard {
  sampled: number;
  sampleLimit: number;
  open: number;
  overdue: number;
  completed: number;
  mine: number;
  dueThisWeek: number;
  byStatus: Record<string, number>;
  workload: Array<{ key: string; name: string; open: number; overdue: number }>;
}

export interface ICurrentUser {
  displayName: string;
  /** Claims login name. */
  loginName: string;
  email: string;
}

export const DEFAULT_ELIGIBLE_STATUSES = 'Completed;Cancelled';
export const MAX_BATCH_SIZE = 200;
export const MAX_ATTACHMENT_BYTES = 25 * 1024 * 1024;
