import { SharePointService, SPEntity } from './SharePointService';
import {
  ICurrentUser, IDashboard, INotification, IRelatedRecord, ITask, IWorkspaceSummary, RecordType, TERMINAL_STATUSES
} from '../models/Models';

const TASK_SELECT = 'Id,Title,WorkDescription,BoardId,WorkspaceId,RecordType,WorkStatus,Priority,AssignedToKey,AssignedToName,DueDate,CompletedDate,Progress,Modified,Created';
const RELATED_SELECT = 'Id,Title,RecordType,WorkDescription,ParentTaskId,RelatedTaskId,ChecklistDone,ActorKey,ActorName,Created';
const DASHBOARD_SAMPLE = 2000;

export interface ITaskInput {
  title: string;
  description: string;
  status: string;
  priority: string;
  assignedToKey?: string;
  assignedToName?: string;
  /** ISO date-time or undefined to clear. */
  dueDate?: string;
}

const q = SharePointService.odataString;

/** Task and collaboration records. All records of a workspace live in its active list. */
export class WorkService {
  constructor(private readonly sp: SharePointService, private readonly user: ICurrentUser) {}

  public async listTasks(ws: IWorkspaceSummary, boardId: string, limit: number = 500): Promise<ITask[]> {
    const clauses = [`RecordType eq 'Task'`, `WorkspaceId eq '${q(ws.id)}'`];
    if (boardId) clauses.push(`BoardId eq '${q(boardId)}'`);
    return this.sp.getAll<ITask>(
      `${this.sp.listPath(ws.workItemsList)}/items?$select=${TASK_SELECT}&${SharePointService.filter(clauses.join(' and '))}&$orderby=Id desc&$top=${Math.min(limit, 500)}`,
      Math.ceil(limit / 500)
    );
  }

  public async createTask(ws: IWorkspaceSummary, boardId: string, input: ITaskInput): Promise<number> {
    const title = WorkService.validateTitle(input.title);
    const terminal = TERMINAL_STATUSES.indexOf(input.status) >= 0;
    const created = await this.sp.createItem(ws.workItemsList, {
      Title: title, WorkspaceId: ws.id, BoardId: boardId, RecordType: 'Task',
      WorkDescription: input.description.trim(), WorkStatus: input.status, Priority: input.priority,
      AssignedToKey: input.assignedToKey || '', AssignedToName: input.assignedToName || '',
      DueDate: input.dueDate || null,
      Progress: input.status === 'Completed' ? 100 : 0,
      CompletedDate: terminal ? new Date().toISOString() : null
    });
    await this.addActivity(ws, boardId, created.Id, 'Task created');
    if (input.assignedToKey) await this.notify(ws, input.assignedToKey, `You were assigned “${title}”`, created.Id);
    return created.Id;
  }

  public async updateTask(ws: IWorkspaceSummary, task: ITask, input: ITaskInput): Promise<void> {
    const title = WorkService.validateTitle(input.title);
    const body: SPEntity = {
      Title: title, WorkDescription: input.description.trim(), WorkStatus: input.status, Priority: input.priority,
      AssignedToKey: input.assignedToKey || '', AssignedToName: input.assignedToName || '', DueDate: input.dueDate || null
    };
    const wasTerminal = TERMINAL_STATUSES.indexOf(task.WorkStatus) >= 0;
    const isTerminal = TERMINAL_STATUSES.indexOf(input.status) >= 0;
    if (isTerminal && !wasTerminal) body.CompletedDate = new Date().toISOString();
    if (!isTerminal) body.CompletedDate = null;
    if (input.status === 'Completed') body.Progress = 100;
    else if (task.WorkStatus === 'Completed') body.Progress = 0;
    await this.sp.updateItem(ws.workItemsList, task.Id, body);

    const changes: string[] = [];
    if (task.WorkStatus !== input.status) changes.push(`status ${task.WorkStatus} → ${input.status}`);
    if ((task.AssignedToKey || '') !== (input.assignedToKey || '')) changes.push(`assigned to ${input.assignedToName || 'nobody'}`);
    if ((task.Priority || '') !== input.priority) changes.push(`priority ${input.priority}`);
    await this.addActivity(ws, task.BoardId, task.Id, changes.length ? `Updated: ${changes.join(', ')}` : 'Task details updated');
    if (input.assignedToKey && input.assignedToKey !== task.AssignedToKey) {
      await this.notify(ws, input.assignedToKey, `You were assigned “${title}”`, task.Id);
    } else if (input.assignedToKey && task.WorkStatus !== input.status) {
      await this.notify(ws, input.assignedToKey, `“${title}” moved to ${input.status}`, task.Id);
    }
  }

  /** Deletes the task, its child records and any dependency rows on other tasks that point at it. */
  public async deleteTask(ws: IWorkspaceSummary, taskId: number): Promise<void> {
    const children = await this.sp.getAll<{ Id: number }>(
      `${this.sp.listPath(ws.workItemsList)}/items?$select=Id&${SharePointService.filter(`ParentTaskId eq ${taskId} or RelatedTaskId eq ${taskId}`)}&$top=500`, 20
    );
    for (const c of children) await this.sp.deleteItem(ws.workItemsList, c.Id);
    await this.sp.deleteItem(ws.workItemsList, taskId);
  }

  public async listRelated(ws: IWorkspaceSummary, taskId: number): Promise<IRelatedRecord[]> {
    return this.sp.getAll<IRelatedRecord>(
      `${this.sp.listPath(ws.workItemsList)}/items?$select=${RELATED_SELECT}&${SharePointService.filter(`ParentTaskId eq ${taskId}`)}&$orderby=Id asc&$top=500`, 4
    );
  }

  public async addComment(ws: IWorkspaceSummary, task: ITask, text: string): Promise<void> {
    if (!text.trim()) throw new Error('Comment cannot be empty.');
    await this.child(ws, task, 'Comment', 'Comment', { WorkDescription: text.trim() });
    if (task.AssignedToKey && task.AssignedToKey.toLowerCase() !== this.user.loginName.toLowerCase()) {
      await this.notify(ws, task.AssignedToKey, `${this.user.displayName} commented on “${task.Title}”`, task.Id);
    }
  }

  public async addChecklistItem(ws: IWorkspaceSummary, task: ITask, text: string): Promise<void> {
    const title = text.trim();
    if (!title) throw new Error('Checklist item cannot be empty.');
    await this.child(ws, task, 'Checklist', title.slice(0, 255), { ChecklistDone: false });
  }

  public async setChecklistDone(ws: IWorkspaceSummary, task: ITask, item: IRelatedRecord, done: boolean): Promise<void> {
    await this.sp.updateItem(ws.workItemsList, item.Id, { ChecklistDone: done });
    await this.addActivity(ws, task.BoardId, task.Id, `${done ? 'Completed' : 'Reopened'} checklist item “${item.Title}”`);
  }

  public async addDependency(ws: IWorkspaceSummary, task: ITask, prerequisite: ITask, existing: IRelatedRecord[]): Promise<void> {
    if (task.Id === prerequisite.Id) throw new Error('A task cannot depend on itself.');
    if (existing.some(d => d.RecordType === 'Dependency' && d.RelatedTaskId === prerequisite.Id)) throw new Error('That dependency already exists.');
    await this.child(ws, task, 'Dependency', `Depends on “${prerequisite.Title}” (#${prerequisite.Id})`.slice(0, 255), { RelatedTaskId: prerequisite.Id });
  }

  public async removeRelated(ws: IWorkspaceSummary, task: ITask, item: IRelatedRecord): Promise<void> {
    await this.sp.deleteItem(ws.workItemsList, item.Id);
    await this.addActivity(ws, task.BoardId, task.Id, item.RecordType === 'Dependency' ? `Removed dependency: ${item.Title.replace(/^Depends on /, 'depends on ')}` : `Removed checklist item “${item.Title}”`);
  }

  public async listNotifications(ws: IWorkspaceSummary): Promise<INotification[]> {
    const r = await this.sp.getCollection<INotification>(
      `${this.sp.listPath(ws.workItemsList)}/items?$select=Id,Title,WorkDescription,NotificationRead,ParentTaskId,ActorName,Created` +
      `&${SharePointService.filter(`NotificationTo eq '${q(SharePointService.claims(this.user.loginName).toLowerCase())}' and RecordType eq 'Notification'`)}&$orderby=Id desc&$top=50`
    );
    return r.value;
  }

  public markNotificationRead(ws: IWorkspaceSummary, id: number): Promise<void> {
    return this.sp.updateItem(ws.workItemsList, id, { NotificationRead: true });
  }

  /** Dashboard figures computed from the most recent tasks (bounded for large lists). */
  public async dashboard(ws: IWorkspaceSummary): Promise<IDashboard> {
    const rows = await this.sp.getAll<ITask>(
      `${this.sp.listPath(ws.workItemsList)}/items?$select=Id,WorkStatus,DueDate,AssignedToKey,AssignedToName` +
      `&${SharePointService.filter(`RecordType eq 'Task' and WorkspaceId eq '${q(ws.id)}'`)}&$orderby=Id desc&$top=500`,
      DASHBOARD_SAMPLE / 500
    );
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const weekEnd = new Date(today.getTime() + 7 * 86400000);
    const me = SharePointService.claims(this.user.loginName).toLowerCase();
    const d: IDashboard = { sampled: rows.length, sampleLimit: DASHBOARD_SAMPLE, open: 0, overdue: 0, completed: 0, mine: 0, dueThisWeek: 0, byStatus: {}, workload: [] };
    const load: Record<string, { key: string; name: string; open: number; overdue: number }> = {};
    for (const t of rows) {
      d.byStatus[t.WorkStatus] = (d.byStatus[t.WorkStatus] || 0) + 1;
      if (TERMINAL_STATUSES.indexOf(t.WorkStatus) >= 0) { d.completed++; continue; }
      d.open++;
      const due = t.DueDate ? new Date(t.DueDate) : undefined;
      const overdue = !!due && due < today;
      if (overdue) d.overdue++;
      if (due && due >= today && due < weekEnd) d.dueThisWeek++;
      if ((t.AssignedToKey || '').toLowerCase() === me) d.mine++;
      const key = t.AssignedToKey || '';
      if (!load[key]) load[key] = { key, name: t.AssignedToName || (key ? key.split('|').pop() as string : 'Unassigned'), open: 0, overdue: 0 };
      load[key].open++;
      if (overdue) load[key].overdue++;
    }
    d.workload = Object.keys(load).map(k => load[k]).sort((a, b) => b.open - a.open);
    return d;
  }

  public async addActivity(ws: IWorkspaceSummary, boardId: string, taskId: number, message: string): Promise<void> {
    await this.sp.createItem(ws.workItemsList, {
      Title: message.slice(0, 255), WorkspaceId: ws.id, BoardId: boardId, RecordType: 'Activity', ParentTaskId: taskId,
      ActivityType: 'Event', ActorKey: this.user.loginName, ActorName: this.user.displayName
    });
  }

  private async child(ws: IWorkspaceSummary, task: ITask, type: RecordType, title: string, extra: SPEntity): Promise<void> {
    await this.sp.createItem(ws.workItemsList, {
      Title: title, WorkspaceId: ws.id, BoardId: task.BoardId, RecordType: type, ParentTaskId: task.Id,
      ActorKey: this.user.loginName, ActorName: this.user.displayName, ...extra
    });
    const what = type === 'Checklist' ? `Added checklist item “${title}”` : type === 'Dependency' ? `Added dependency: ${title.replace(/^Depends on /, 'depends on ')}` : `Added ${type.toLowerCase()} “${title}”`;
    if (type !== 'Comment') await this.addActivity(ws, task.BoardId, task.Id, what.slice(0, 255));
  }

  private async notify(ws: IWorkspaceSummary, toKey: string, message: string, taskId: number): Promise<void> {
    try {
      await this.sp.createItem(ws.workItemsList, {
        Title: message.slice(0, 255), WorkspaceId: ws.id, RecordType: 'Notification', ParentTaskId: taskId,
        NotificationTo: SharePointService.claims(toKey).toLowerCase(), NotificationRead: false,
        ActorKey: this.user.loginName, ActorName: this.user.displayName
      });
    } catch {
      // Notifications are best effort; never fail the user's action because of them.
    }
  }

  private static validateTitle(title: string): string {
    const t = title.trim();
    if (!t) throw new Error('Enter a task title.');
    if (t.length > 255) throw new Error('Task titles must be 255 characters or fewer.');
    return t;
  }
}
