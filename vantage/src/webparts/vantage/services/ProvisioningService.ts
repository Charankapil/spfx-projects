import { IFieldDefinition, SharePointService } from './SharePointService';
import { DEFAULT_ELIGIBLE_STATUSES, IWorkspaceSummary } from '../models/Models';

export const LISTS = {
  settings: 'VNT_Settings',
  workspaces: 'VNT_Workspaces',
  policies: 'VNT_ArchivePolicies',
  jobs: 'VNT_ArchiveJobs',
  log: 'VNT_ArchiveLog'
};

export const SCHEMA_VERSION = '2';

const CONTROL_LISTS: Array<{ title: string; description: string; fields: IFieldDefinition[] }> = [
  {
    title: LISTS.settings,
    description: 'VANTAGE application settings.',
    fields: [{ name: 'SettingValue', displayName: 'Value', kind: 'Note' }]
  },
  {
    title: LISTS.workspaces,
    description: 'VANTAGE workspace registry.',
    fields: [
      { name: 'WorkspaceKey', displayName: 'Workspace Key', kind: 'Text', indexed: true },
      { name: 'WorkspaceStatus', displayName: 'Status', kind: 'Text', indexed: true },
      { name: 'WorkItemsList', displayName: 'Work Items List', kind: 'Text' },
      { name: 'ArchiveList', displayName: 'Archive List', kind: 'Text' }
    ]
  },
  {
    title: LISTS.policies,
    description: 'VANTAGE workspace archive policies.',
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
    description: 'VANTAGE archive job history and status.',
    fields: [
      { name: 'WorkspaceId', displayName: 'Workspace ID', kind: 'Text', indexed: true },
      { name: 'JobStatus', displayName: 'Job Status', kind: 'Text', indexed: true },
      { name: 'ArchiveMode', displayName: 'Archive Mode', kind: 'Text' },
      { name: 'Started', displayName: 'Started', kind: 'DateTime' },
      { name: 'Finished', displayName: 'Finished', kind: 'DateTime' },
      { name: 'Processed', displayName: 'Processed', kind: 'Number' },
      { name: 'Archived', displayName: 'Archived', kind: 'Number' },
      { name: 'Skipped', displayName: 'Skipped', kind: 'Number' },
      { name: 'FailedCount', displayName: 'Failed Count', kind: 'Number' },
      { name: 'LastError', displayName: 'Last Error', kind: 'Note' }
    ]
  },
  {
    title: LISTS.log,
    description: 'VANTAGE auditable archive operation log.',
    fields: [
      { name: 'WorkspaceId', displayName: 'Workspace ID', kind: 'Text', indexed: true },
      { name: 'OriginalItemId', displayName: 'Original Item ID', kind: 'Number', indexed: true },
      { name: 'ArchiveItemId', displayName: 'Archive Item ID', kind: 'Number' },
      { name: 'ArchiveJobId', displayName: 'Archive Job ID', kind: 'Number', indexed: true },
      { name: 'Outcome', displayName: 'Outcome', kind: 'Text' },
      { name: 'Detail', displayName: 'Detail', kind: 'Note' }
    ]
  }
];

/** Fields shared by a workspace's active list and its archive list. */
export const WORK_ITEM_FIELDS: IFieldDefinition[] = [
  { name: 'WorkspaceId', displayName: 'Workspace ID', kind: 'Text', indexed: true, showInView: false },
  { name: 'BoardId', displayName: 'Board ID', kind: 'Text', indexed: true, showInView: false },
  { name: 'RecordType', displayName: 'Record Type', kind: 'Text', indexed: true },
  { name: 'WorkStatus', displayName: 'Status', kind: 'Text', indexed: true },
  { name: 'Priority', displayName: 'Priority', kind: 'Text' },
  { name: 'AssignedToKey', displayName: 'Assigned To Key', kind: 'Text', indexed: true },
  { name: 'DueDate', displayName: 'Due Date', kind: 'DateTime', indexed: true },
  { name: 'CompletedDate', displayName: 'Completed Date', kind: 'DateTime', indexed: true },
  { name: 'Progress', displayName: 'Progress', kind: 'Number' },
  { name: 'WorkDescription', displayName: 'Description', kind: 'Note' },
  ...[1, 2, 3, 4, 5, 6, 7, 8].map(i =>
    ({ name: `CustomText${i}`, displayName: `Custom Text ${i}`, kind: 'Text', showInView: false } as IFieldDefinition)),
  ...[1, 2, 3, 4].map(i =>
    ({ name: `CustomNumber${i}`, displayName: `Custom Number ${i}`, kind: 'Number', showInView: false } as IFieldDefinition)),
  ...[1, 2, 3, 4].map(i =>
    ({ name: `CustomDate${i}`, displayName: `Custom Date ${i}`, kind: 'DateTime', showInView: false } as IFieldDefinition)),
  ...[1, 2].map(i =>
    ({ name: `CustomBoolean${i}`, displayName: `Custom Boolean ${i}`, kind: 'Boolean', showInView: false } as IFieldDefinition))
];

const ARCHIVE_ONLY_FIELDS: IFieldDefinition[] = [
  { name: 'OriginalItemId', displayName: 'Original Item ID', kind: 'Number', indexed: true },
  { name: 'ArchivedAt', displayName: 'Archived At', kind: 'DateTime', indexed: true },
  { name: 'ArchiveJobId', displayName: 'Archive Job ID', kind: 'Number' },
  { name: 'OriginalCreated', displayName: 'Original Created', kind: 'DateTime', showInView: false },
  { name: 'OriginalModified', displayName: 'Original Modified', kind: 'DateTime', showInView: false }
];

export interface ISetupState {
  initialized: boolean;
  missingLists: string[];
}

export class ProvisioningService {
  constructor(private readonly sp: SharePointService) {}

  public async getSetupState(): Promise<ISetupState> {
    const titles = await this.sp.getListTitles('VNT_');
    const present = new Set(titles);
    const missingLists = CONTROL_LISTS.map(l => l.title).filter(t => !present.has(t));
    return { initialized: missingLists.length === 0, missingLists };
  }

  /** Idempotent: safe to re-run to repair a partially completed setup. */
  public async initialize(onStep?: (message: string) => void): Promise<void> {
    for (const list of CONTROL_LISTS) {
      if (onStep) onStep(`Provisioning ${list.title}…`);
      await this.sp.ensureList(list.title, list.description);
      await this.sp.ensureFields(list.title, list.fields);
    }
    await this.upsertSetting('SchemaVersion', SCHEMA_VERSION);
  }

  public async listWorkspaces(): Promise<IWorkspaceSummary[]> {
    const response = await this.sp.getCollection<{ Title: string; WorkspaceKey: string; WorkItemsList: string; ArchiveList: string }>(
      `${this.sp.listPath(LISTS.workspaces)}/items?$select=Id,Title,WorkspaceKey,WorkItemsList,ArchiveList` +
      `&$filter=${encodeURIComponent("WorkspaceStatus eq 'Active'")}&$orderby=Title asc&$top=5000`
    );
    return response.value.map(item => ({
      id: item.WorkspaceKey,
      title: item.Title,
      workItemsList: item.WorkItemsList,
      archiveList: item.ArchiveList
    }));
  }

  public async createWorkspace(name: string, existing: IWorkspaceSummary[]): Promise<IWorkspaceSummary> {
    const title = name.trim().replace(/\s+/g, ' ');
    if (!title) throw new Error('Enter a workspace name.');
    if (title.length > 120) throw new Error('Workspace names must be 120 characters or fewer.');
    if (existing.some(w => w.title.toLowerCase() === title.toLowerCase())) {
      throw new Error(`A workspace named “${title}” already exists. Select it instead of creating a duplicate.`);
    }

    const key = `WS-${Date.now().toString(36).toUpperCase()}-${Math.random().toString(36).slice(2, 7).toUpperCase()}`;
    const slug = key.replace(/-/g, '');
    const workItemsList = `VNT_${slug}_Items`;
    const archiveList = `VNT_${slug}_Archive`;

    await this.sp.ensureList(workItemsList, `Active VANTAGE work items for ${title}.`);
    await this.sp.ensureFields(workItemsList, WORK_ITEM_FIELDS);
    await this.sp.ensureList(archiveList, `Archived VANTAGE work items for ${title}.`);
    await this.sp.ensureFields(archiveList, WORK_ITEM_FIELDS.concat(ARCHIVE_ONLY_FIELDS));

    // The registry row is written last, so a failed provisioning never shows up as a usable workspace.
    await this.sp.createItem(LISTS.workspaces, {
      Title: title, WorkspaceKey: key, WorkItemsList: workItemsList, ArchiveList: archiveList, WorkspaceStatus: 'Active'
    });
    await this.sp.createItem(LISTS.policies, {
      Title: `${title} archive policy`, WorkspaceId: key, Enabled: false, ArchiveMode: 'date',
      MaxActiveItems: 100000, EligibleStatuses: DEFAULT_ELIGIBLE_STATUSES, BatchSize: 100
    });
    return { id: key, title, workItemsList, archiveList };
  }

  private async upsertSetting(key: string, value: string): Promise<void> {
    const existing = await this.sp.getCollection<{ Id: number }>(
      `${this.sp.listPath(LISTS.settings)}/items?$select=Id&$filter=${encodeURIComponent(`Title eq '${SharePointService.odataString(key)}'`)}&$top=1`
    );
    if (existing.value.length) await this.sp.updateItem(LISTS.settings, existing.value[0].Id, { SettingValue: value });
    else await this.sp.createItem(LISTS.settings, { Title: key, SettingValue: value });
  }
}
