import { IFieldDefinition, SharePointService, SPEntity } from './SharePointService';
import { DEFAULT_ELIGIBLE_STATUSES, IBoardSummary, IWorkspaceSummary } from '../models/Models';

export const LISTS = {
  settings: 'VNT_Settings',
  workspaces: 'VNT_Workspaces',
  boards: 'VNT_Boards',
  policies: 'VNT_ArchivePolicies',
  jobs: 'VNT_ArchiveJobs',
  log: 'VNT_ArchiveLog',
  locks: 'VNT_ArchiveLocks'
};

/** v2 = 0.2.0 schema, v3 = 1.0.0 schema (boards, related records, notifications, locks). */
export const SCHEMA_VERSION = 3;
/** SharePoint's supported limit of lists and libraries per site collection. */
const SITE_LIST_LIMIT = 2000;

const CONTROL_LISTS: Array<{ title: string; description: string; fields: IFieldDefinition[] }> = [
  {
    title: LISTS.settings,
    description: 'Vantage application settings.',
    fields: [{ name: 'SettingValue', displayName: 'Value', kind: 'Note' }]
  },
  {
    title: LISTS.workspaces,
    description: 'Vantage workspace registry.',
    fields: [
      { name: 'WorkspaceKey', displayName: 'Workspace Key', kind: 'Text', indexed: true },
      { name: 'WorkspaceStatus', displayName: 'Status', kind: 'Text', indexed: true },
      { name: 'WorkItemsList', displayName: 'Work Items List', kind: 'Text' },
      { name: 'ArchiveList', displayName: 'Archive List', kind: 'Text' },
      { name: 'OwnerKey', displayName: 'Owner', kind: 'Text' },
      { name: 'PermissionsConfigured', displayName: 'Permissions Configured', kind: 'Boolean' }
    ]
  },
  {
    title: LISTS.boards,
    description: 'Vantage board registry.',
    fields: [
      { name: 'BoardKey', displayName: 'Board Key', kind: 'Text', unique: true },
      { name: 'WorkspaceId', displayName: 'Workspace ID', kind: 'Text', indexed: true },
      { name: 'BoardStatus', displayName: 'Status', kind: 'Text', indexed: true },
      { name: 'BoardDescription', displayName: 'Description', kind: 'Note' }
    ]
  },
  {
    title: LISTS.policies,
    description: 'Vantage workspace archive policies.',
    fields: [
      { name: 'WorkspaceId', displayName: 'Workspace ID', kind: 'Text', indexed: true },
      { name: 'Enabled', displayName: 'Enabled', kind: 'Boolean' },
      { name: 'ArchiveMode', displayName: 'Archive Mode', kind: 'Text' },
      { name: 'MaxActiveItems', displayName: 'Maximum Active Items', kind: 'Number' },
      { name: 'ArchiveBefore', displayName: 'Archive Before', kind: 'DateTime' },
      { name: 'EligibleStatuses', displayName: 'Eligible Statuses', kind: 'Text' },
      { name: 'BatchSize', displayName: 'Batch Size', kind: 'Number' }
    ]
  },
  {
    title: LISTS.jobs,
    description: 'Vantage archive job history and status.',
    fields: [
      { name: 'WorkspaceId', displayName: 'Workspace ID', kind: 'Text', indexed: true },
      { name: 'JobStatus', displayName: 'Job Status', kind: 'Text', indexed: true },
      { name: 'ArchiveMode', displayName: 'Archive Mode', kind: 'Text' },
      { name: 'Started', displayName: 'Started', kind: 'DateTime' },
      { name: 'Finished', displayName: 'Finished', kind: 'DateTime' },
      { name: 'Processed', displayName: 'Processed', kind: 'Number' },
      { name: 'Archived', displayName: 'Archived Tasks', kind: 'Number' },
      { name: 'Records', displayName: 'Archived Records', kind: 'Number' },
      { name: 'Skipped', displayName: 'Skipped', kind: 'Number' },
      { name: 'FailedCount', displayName: 'Failed Count', kind: 'Number' },
      { name: 'LastError', displayName: 'Last Error', kind: 'Note' }
    ]
  },
  {
    title: LISTS.log,
    description: 'Vantage auditable archive and restore log.',
    fields: [
      { name: 'WorkspaceId', displayName: 'Workspace ID', kind: 'Text', indexed: true },
      { name: 'OriginalItemId', displayName: 'Original Item ID', kind: 'Number', indexed: true },
      { name: 'ArchiveItemId', displayName: 'Archive Item ID', kind: 'Number' },
      { name: 'ArchiveJobId', displayName: 'Archive Job ID', kind: 'Number', indexed: true },
      { name: 'Outcome', displayName: 'Outcome', kind: 'Text' },
      { name: 'Detail', displayName: 'Detail', kind: 'Note' }
    ]
  },
  {
    title: LISTS.locks,
    description: 'Vantage per-workspace archive locks.',
    fields: [
      // Unique values make lock acquisition atomic: a second create for the same workspace fails.
      { name: 'WorkspaceId', displayName: 'Workspace ID', kind: 'Text', unique: true },
      { name: 'JobId', displayName: 'Job ID', kind: 'Number' },
      { name: 'HeartbeatEpoch', displayName: 'Heartbeat (epoch seconds)', kind: 'Number' },
      { name: 'LockOwner', displayName: 'Lock Owner', kind: 'Text' }
    ]
  }
];

const customSlots = (prefix: string, display: string, kind: IFieldDefinition['kind'], count: number): IFieldDefinition[] => {
  const out: IFieldDefinition[] = [];
  for (let i = 1; i <= count; i++) out.push({ name: `${prefix}${i}`, displayName: `${display} ${i}`, kind, showInView: false });
  return out;
};

/** Fields shared by a workspace's active list and its archive list (and copied on archive/restore). */
export const WORK_ITEM_FIELDS: IFieldDefinition[] = [
  { name: 'WorkspaceId', displayName: 'Workspace ID', kind: 'Text', indexed: true, showInView: false },
  { name: 'BoardId', displayName: 'Board ID', kind: 'Text', indexed: true, showInView: false },
  { name: 'RecordType', displayName: 'Record Type', kind: 'Text', indexed: true },
  { name: 'WorkStatus', displayName: 'Status', kind: 'Text', indexed: true },
  { name: 'Priority', displayName: 'Priority', kind: 'Text' },
  { name: 'AssignedToKey', displayName: 'Assigned To (login)', kind: 'Text', indexed: true, showInView: false },
  { name: 'AssignedToName', displayName: 'Assigned To', kind: 'Text' },
  { name: 'DueDate', displayName: 'Due Date', kind: 'DateTime', indexed: true },
  { name: 'CompletedDate', displayName: 'Completed Date', kind: 'DateTime', indexed: true },
  { name: 'Progress', displayName: 'Progress', kind: 'Number' },
  { name: 'WorkDescription', displayName: 'Description', kind: 'Note' },
  { name: 'ParentTaskId', displayName: 'Parent Task ID', kind: 'Number', indexed: true, showInView: false },
  { name: 'RelatedTaskId', displayName: 'Related Task ID', kind: 'Number', indexed: true, showInView: false },
  { name: 'ChecklistDone', displayName: 'Checklist Done', kind: 'Boolean', showInView: false },
  { name: 'ActivityType', displayName: 'Activity Type', kind: 'Text', showInView: false },
  { name: 'ActorKey', displayName: 'Actor (login)', kind: 'Text', showInView: false },
  { name: 'ActorName', displayName: 'Actor', kind: 'Text', showInView: false },
  { name: 'NotificationTo', displayName: 'Notification To', kind: 'Text', indexed: true, showInView: false },
  { name: 'NotificationRead', displayName: 'Notification Read', kind: 'Boolean', showInView: false },
  ...customSlots('CustomText', 'Custom Text', 'Text', 8),
  ...customSlots('CustomNumber', 'Custom Number', 'Number', 4),
  ...customSlots('CustomDate', 'Custom Date', 'DateTime', 4),
  ...customSlots('CustomBoolean', 'Custom Boolean', 'Boolean', 2)
];

const ACTIVE_ONLY_FIELDS: IFieldDefinition[] = [
  { name: 'RestoredFromArchiveId', displayName: 'Restored From Archive ID', kind: 'Number', unique: true, showInView: false }
];

const ARCHIVE_ONLY_FIELDS: IFieldDefinition[] = [
  { name: 'OriginalItemId', displayName: 'Original Item ID', kind: 'Number', unique: true },
  { name: 'ArchivedAt', displayName: 'Archived At', kind: 'DateTime', indexed: true },
  { name: 'ArchiveJobId', displayName: 'Archive Job ID', kind: 'Number' },
  { name: 'OriginalCreated', displayName: 'Original Created', kind: 'DateTime', showInView: false },
  { name: 'OriginalModified', displayName: 'Original Modified', kind: 'DateTime', showInView: false }
];

export interface ISetupState {
  /** All lists exist and the schema is current. */
  initialized: boolean;
  /** Lists exist but were created by an older version and must be upgraded. */
  upgradeRequired: boolean;
  missingLists: string[];
  schemaVersion: number;
}

const newKey = (prefix: string): string =>
  `${prefix}-${Date.now().toString(36).toUpperCase()}-${Math.random().toString(36).slice(2, 7).toUpperCase()}`;

export class ProvisioningService {
  constructor(private readonly sp: SharePointService) {}

  public async getSetupState(): Promise<ISetupState> {
    const present = new Set(await this.sp.getListTitles('VNT_'));
    const missingLists = CONTROL_LISTS.map(l => l.title).filter(t => !present.has(t));
    let schemaVersion = 0;
    if (present.has(LISTS.settings)) {
      try {
        schemaVersion = Number(await this.getSetting('SchemaVersion')) || 0;
      } catch {
        schemaVersion = 0;
      }
    }
    const legacyPresent = missingLists.length < CONTROL_LISTS.length;
    return {
      initialized: missingLists.length === 0 && schemaVersion >= SCHEMA_VERSION,
      upgradeRequired: legacyPresent && schemaVersion > 0 && schemaVersion < SCHEMA_VERSION,
      missingLists,
      schemaVersion
    };
  }

  /** Idempotent: creates or upgrades every control list and every registered workspace's lists. */
  public async initialize(onStep?: (message: string) => void): Promise<void> {
    const step = (m: string): void => { if (onStep) onStep(m); };
    for (const list of CONTROL_LISTS) {
      step(`Provisioning ${list.title}…`);
      await this.sp.ensureList(list.title, list.description);
      await this.sp.ensureFields(list.title, list.fields);
    }
    const workspaces = await this.listWorkspaces();
    for (const ws of workspaces) {
      step(`Upgrading workspace “${ws.title}”…`);
      await this.provisionWorkspaceLists(ws.title, ws.workItemsList, ws.archiveList);
    }
    await this.upsertSetting('SchemaVersion', String(SCHEMA_VERSION));
  }

  public async listWorkspaces(): Promise<IWorkspaceSummary[]> {
    const rows = await this.sp.getAll<SPEntity>(
      `${this.sp.listPath(LISTS.workspaces)}/items?$select=Id,Title,WorkspaceKey,WorkItemsList,ArchiveList,OwnerKey,PermissionsConfigured` +
      `&${SharePointService.filter("WorkspaceStatus eq 'Active'")}&$orderby=Title asc&$top=500`, 10
    );
    return rows.map(item => ({
      id: item.WorkspaceKey as string,
      itemId: item.Id as number,
      title: item.Title as string,
      workItemsList: item.WorkItemsList as string,
      archiveList: item.ArchiveList as string,
      ownerKey: item.OwnerKey as string | undefined,
      secured: !!item.PermissionsConfigured
    }));
  }

  public async createWorkspace(name: string, existing: IWorkspaceSummary[], ownerKey: string): Promise<IWorkspaceSummary> {
    const title = name.trim().replace(/\s+/g, ' ');
    if (!title) throw new Error('Enter a workspace name.');
    if (title.length > 80) throw new Error('Workspace names must be 80 characters or fewer.');
    if (existing.some(w => w.title.toLowerCase() === title.toLowerCase())) {
      throw new Error(`A workspace named “${title}” already exists. Select it instead of creating a duplicate.`);
    }
    const listCount = await this.sp.getSiteListCount();
    if (listCount + 2 > SITE_LIST_LIMIT) {
      throw new Error(`This site already has ${listCount} lists and libraries. A workspace needs two more and the supported limit is ${SITE_LIST_LIMIT}. Use another site for new workspaces.`);
    }

    const key = newKey('WS');
    const slug = key.replace(/-/g, '');
    const workItemsList = `VNT_${slug}_Items`;
    const archiveList = `VNT_${slug}_Archive`;
    await this.provisionWorkspaceLists(title, workItemsList, archiveList);

    // The registry row is written last, so a failed provisioning never shows up as a usable workspace.
    const row = await this.sp.createItem(LISTS.workspaces, {
      Title: title, WorkspaceKey: key, WorkItemsList: workItemsList, ArchiveList: archiveList,
      WorkspaceStatus: 'Active', OwnerKey: ownerKey, PermissionsConfigured: false
    });
    await this.sp.createItem(LISTS.policies, {
      Title: `${title} archive policy`, WorkspaceId: key, Enabled: false, ArchiveMode: 'date',
      MaxActiveItems: 100000, EligibleStatuses: DEFAULT_ELIGIBLE_STATUSES, BatchSize: 100
    });
    return { id: key, itemId: row.Id, title, workItemsList, archiveList, ownerKey, secured: false };
  }

  public async markWorkspaceSecured(ws: IWorkspaceSummary): Promise<void> {
    await this.sp.updateItem(LISTS.workspaces, ws.itemId, { PermissionsConfigured: true });
  }

  public async listBoards(workspaceId: string): Promise<IBoardSummary[]> {
    const rows = await this.sp.getAll<SPEntity>(
      `${this.sp.listPath(LISTS.boards)}/items?$select=Id,Title,BoardKey,WorkspaceId,BoardDescription` +
      `&${SharePointService.filter(`WorkspaceId eq '${SharePointService.odataString(workspaceId)}' and BoardStatus eq 'Active'`)}&$orderby=Title asc&$top=500`, 5
    );
    return rows.map(x => ({
      id: x.BoardKey as string, itemId: x.Id as number, title: x.Title as string,
      workspaceId: x.WorkspaceId as string, description: x.BoardDescription as string | undefined
    }));
  }

  public async createBoard(workspaceId: string, name: string, description: string, existing: IBoardSummary[]): Promise<IBoardSummary> {
    const title = name.trim().replace(/\s+/g, ' ');
    if (!title) throw new Error('Enter a board name.');
    if (title.length > 120) throw new Error('Board names must be 120 characters or fewer.');
    if (existing.some(b => b.title.toLowerCase() === title.toLowerCase())) throw new Error(`A board named “${title}” already exists in this workspace.`);
    const key = newKey('BD');
    const row = await this.sp.createItem(LISTS.boards, {
      Title: title, BoardKey: key, WorkspaceId: workspaceId, BoardDescription: description.trim(), BoardStatus: 'Active'
    });
    return { id: key, itemId: row.Id, title, workspaceId, description: description.trim() };
  }

  private async provisionWorkspaceLists(title: string, workItemsList: string, archiveList: string): Promise<void> {
    await this.sp.ensureList(workItemsList, `Active Vantage work items for ${title}.`);
    await this.sp.ensureFields(workItemsList, WORK_ITEM_FIELDS.concat(ACTIVE_ONLY_FIELDS));
    await this.sp.ensureList(archiveList, `Archived Vantage work items for ${title}.`);
    await this.sp.ensureFields(archiveList, WORK_ITEM_FIELDS.concat(ARCHIVE_ONLY_FIELDS));
  }

  private async getSetting(key: string): Promise<string | undefined> {
    const r = await this.sp.getCollection<{ SettingValue?: string }>(
      `${this.sp.listPath(LISTS.settings)}/items?$select=Id,SettingValue&${SharePointService.filter(`Title eq '${SharePointService.odataString(key)}'`)}&$top=1`
    );
    return r.value.length ? r.value[0].SettingValue : undefined;
  }

  private async upsertSetting(key: string, value: string): Promise<void> {
    const existing = await this.sp.getCollection<{ Id: number }>(
      `${this.sp.listPath(LISTS.settings)}/items?$select=Id&${SharePointService.filter(`Title eq '${SharePointService.odataString(key)}'`)}&$top=1`
    );
    if (existing.value.length) await this.sp.updateItem(LISTS.settings, existing.value[0].Id, { SettingValue: value });
    else await this.sp.createItem(LISTS.settings, { Title: key, SettingValue: value });
  }
}
