import { SpClient, SpError } from './SpClient';
import { ensureField, ensureList, findList, deleteField, renameField, recycleList, setTitleFieldName, itemEntityType } from './lists';
import { LIST_BOARDS } from './Provisioner';
import { IBoard, IBoardColumn, IBoardConfig, BoardPrivacy, ColumnType, IBoardGroup, ILabel, IBoardRoles } from '../models/types';
import { getTemplate, STATUS_LABELS } from '../models/templates';
import { F, UF, fieldXml } from '../engine/fieldMap';
import { shortId } from '../engine/ids';

/* eslint-disable @typescript-eslint/no-explicit-any */

const REGISTRY_SELECT = 'Id,Title,WB_Key,WB_Description,WB_Folder,WB_Color,WB_Privacy,WB_ItemsListId,WB_UpdatesListId,WB_BoardOwnersId,WB_ProjectHeadId,WB_ProjectLeadId,WB_ProjectSponsorId,WB_Config,WB_Archived';

export function itemsListTitle(key: string): string {
  return 'WB_Board_' + key;
}

export function updatesListTitle(key: string): string {
  return 'WB_Board_' + key + '_Updates';
}

export function parseConfig(raw: string | null | undefined): IBoardConfig {
  let cfg: Partial<IBoardConfig> = {};
  try {
    cfg = raw ? JSON.parse(raw) : {};
  } catch {
    cfg = {};
  }
  const groups = Array.isArray(cfg.groups) && cfg.groups.length > 0 ? cfg.groups : [{ id: 'g_default', title: 'Group', color: '#579bfc' }];
  return {
    v: 1,
    columns: Array.isArray(cfg.columns) ? cfg.columns : [],
    groups,
    kanbanColumnId: cfg.kanbanColumnId,
    timelineColumnId: cfg.timelineColumnId,
    defaultView: cfg.defaultView || 'table',
    linksField: cfg.linksField === true
  };
}

function ids(raw: any): number[] {
  return Array.isArray(raw) ? raw : raw && Array.isArray(raw.results) ? raw.results : [];
}

function readBoard(row: any): IBoard {
  const owners = row.WB_BoardOwnersId;
  return {
    id: row.Id,
    title: row.Title || '',
    key: row.WB_Key || '',
    description: row.WB_Description || '',
    folder: row.WB_Folder || '',
    color: row.WB_Color || '#579bfc',
    privacy: row.WB_Privacy === 'Private' ? 'Private' : 'Main',
    itemsListId: row.WB_ItemsListId || '',
    updatesListId: row.WB_UpdatesListId || '',
    ownerIds: ids(owners),
    roles: { head: ids(row.WB_ProjectHeadId), lead: ids(row.WB_ProjectLeadId), sponsor: ids(row.WB_ProjectSponsorId) },
    archived: row.WB_Archived === true,
    config: parseConfig(row.WB_Config),
    etag: row['odata.etag'] || '*'
  };
}

/** Which SharePoint field(s) a new column of this type should use. Reuses a free fixed field first. */
export function pickFieldsForNewColumn(config: IBoardConfig, type: ColumnType): { field: string; fieldEnd?: string } {
  const used: string[] = [];
  config.columns.forEach(c => {
    used.push(c.field);
    if (c.fieldEnd) {
      used.push(c.fieldEnd);
    }
  });
  const free = (f: string): boolean => used.indexOf(f) < 0;
  if (type === 'people' && free(F.Owner)) {
    return { field: F.Owner };
  }
  if (type === 'status' && free(F.Status)) {
    return { field: F.Status };
  }
  if (type === 'timeline' && free(F.StartDate) && free(F.DueDate)) {
    return { field: F.StartDate, fieldEnd: F.DueDate };
  }
  if (type === 'date' && free(F.DueDate)) {
    return { field: F.DueDate };
  }
  const base = 'WB_c_' + shortId(6);
  return type === 'timeline' ? { field: base, fieldEnd: base + 'e' } : { field: base };
}

function isFixedField(field: string): boolean {
  return [F.Owner, F.Status, F.StartDate, F.DueDate].indexOf(field) >= 0;
}

export interface ICreateBoardInput {
  title: string;
  key: string;
  description: string;
  folder: string;
  color: string;
  privacy: BoardPrivacy;
  templateId: string;
  memberIds: number[];
  roles?: IBoardRoles;
}

export type RoleName = 'Owner' | 'Member' | 'Viewer';

export interface IBoardMember {
  principalId: number;
  title: string;
  isGroup: boolean;
  roles: string[];
}

/** RoleTypeKind: Reader 2, Contributor 3, Administrator 5, Editor 6. */
const ROLE_TYPE: { [role: string]: number } = { Viewer: 2, Member: 3, Owner: 6, Full: 5 };

export class BoardService {
  private readonly sp: SpClient;
  private registryId: string | null = null;
  private readonly roleDefIds: { [type: number]: number } = {};

  constructor(sp: SpClient) {
    this.sp = sp;
  }

  private async registry(): Promise<string> {
    if (!this.registryId) {
      const list = await findList(this.sp, LIST_BOARDS);
      if (!list) {
        throw new Error('Work Boards is not set up on this site.');
      }
      this.registryId = list.Id;
    }
    return this.registryId;
  }

  public async listBoards(): Promise<IBoard[]> {
    const reg = await this.registry();
    const rows = await this.sp.getAll<any>(`web/lists(guid'${reg}')/items?$select=${REGISTRY_SELECT}&$top=500`);
    return rows.map(readBoard).sort((a, b) => a.title.localeCompare(b.title));
  }

  public async getBoard(id: number): Promise<IBoard> {
    const reg = await this.registry();
    const row = await this.sp.get<any>(`web/lists(guid'${reg}')/items(${id})?$select=${REGISTRY_SELECT}`);
    return readBoard(row);
  }

  /** Create lists, fields and the registry row for a new board. Steps are reported through onStep. */
  public async createBoard(input: ICreateBoardInput, meId: number, onStep: (message: string) => void): Promise<IBoard> {
    const key = input.key.toUpperCase();
    const template = getTemplate(input.templateId);

    const existing = await this.listBoards();
    if (existing.some(b => b.key.toUpperCase() === key) || (await findList(this.sp, itemsListTitle(key)))) {
      throw new Error(`The board key ${key} is already in use on this site. Choose another key.`);
    }

    // Columns: template-local ids stay; new columns get their own SharePoint fields.
    const columns: IBoardColumn[] = template.columns.map(c => {
      const col: IBoardColumn = { ...(c as IBoardColumn) };
      if (!c.field) {
        // Template columns that don't name a fixed field get their own new field.
        col.field = 'WB_c_' + shortId(6);
        if (c.type === 'timeline') {
          col.fieldEnd = col.field + 'e';
        }
      }
      if (col.labels) {
        col.labels = col.labels.map(l => ({ ...l }));
      }
      return col;
    });
    const config: IBoardConfig = {
      v: 1,
      columns,
      groups: template.groups.map(g => ({ ...g })),
      kanbanColumnId: template.kanbanColumnId,
      timelineColumnId: template.timelineColumnId,
      defaultView: template.defaultView,
      linksField: true
    };

    onStep('Creating the board list');
    const itemsList = await ensureList(this.sp, itemsListTitle(key), {
      description: `Work Boards: items of "${input.title}".`,
      hidden: true,
      versioning: true,
      attachments: true
    });
    await setTitleFieldName(this.sp, itemsList.Id, 'Item');

    onStep('Adding columns');
    const statusCol = columns.filter(c => c.field === F.Status)[0];
    const fixed: [string, string][] = [
      [F.GroupId, fieldXml('text', F.GroupId, 'Group', { indexed: true })],
      [F.SortOrder, fieldXml('sortnumber', F.SortOrder, 'Sort order', { indexed: true })],
      [F.ParentId, fieldXml('id', F.ParentId, 'Parent item', { indexed: true })],
      [F.Owner, fieldXml('people', F.Owner, 'Owner')],
      [F.Status, fieldXml('status', F.Status, 'Status', { indexed: true, choices: (statusCol ? statusCol.labels || [] : STATUS_LABELS).map(l => l.text) })],
      [F.StartDate, fieldXml('date', F.StartDate, 'Start date', { indexed: true })],
      [F.DueDate, fieldXml('date', F.DueDate, 'Due date', { indexed: true })],
      [F.Links, fieldXml('longtext', F.Links, 'Linked files')]
    ];
    for (const [name, xml] of fixed) {
      await ensureField(this.sp, itemsList.Id, name, xml);
    }
    for (const col of columns) {
      await this.ensureColumnFields(itemsList.Id, col);
    }

    onStep('Creating the updates list');
    const updatesList = await ensureList(this.sp, updatesListTitle(key), {
      description: `Work Boards: updates (comments) on "${input.title}".`,
      hidden: true,
      versioning: false,
      attachments: false
    });
    const updFields: [string, string][] = [
      [UF.ItemId, fieldXml('id', UF.ItemId, 'Item id', { indexed: true })],
      [UF.ParentId, fieldXml('id', UF.ParentId, 'Reply to', { indexed: true })],
      [UF.Body, fieldXml('longtext', UF.Body, 'Body')],
      [UF.Mentions, fieldXml('people', UF.Mentions, 'Mentions')],
      [UF.LikedBy, fieldXml('people', UF.LikedBy, 'Liked by')]
    ];
    for (const [name, xml] of updFields) {
      await ensureField(this.sp, updatesList.Id, name, xml);
    }

    onStep('Registering the board');
    const reg = await this.registry();
    const type = await itemEntityType(this.sp, reg);
    const created = await this.sp.post<any>(`web/lists(guid'${reg}')/items`, {
      __metadata: { type },
      Title: input.title,
      WB_Key: key,
      WB_Description: input.description,
      WB_Folder: input.folder,
      WB_Color: input.color,
      WB_Privacy: 'Main',
      WB_ItemsListId: itemsList.Id,
      WB_UpdatesListId: updatesList.Id,
      WB_BoardOwnersId: { results: [meId] },
      WB_ProjectHeadId: { results: input.roles ? input.roles.head : [] },
      WB_ProjectLeadId: { results: input.roles ? input.roles.lead : [] },
      WB_ProjectSponsorId: { results: input.roles ? input.roles.sponsor : [] },
      WB_Config: JSON.stringify(config),
      WB_Archived: false
    });
    let board = await this.getBoard(created.Id);

    if (input.privacy === 'Private') {
      onStep('Making the board private');
      board = await this.makePrivate(board, meId, input.memberIds);
      await this.grantRoleViewers(board, input.memberIds.concat([meId]));
    }
    return board;
  }

  private async ensureColumnFields(listId: string, col: IBoardColumn): Promise<void> {
    if (isFixedField(col.field)) {
      return;
    }
    const choices = (col.labels || []).map(l => l.text);
    await ensureField(this.sp, listId, col.field, fieldXml(col.type, col.field, col.title, { choices }));
    if (col.type === 'timeline' && col.fieldEnd) {
      await ensureField(this.sp, listId, col.fieldEnd, fieldXml('date', col.fieldEnd, col.title + ' (end)'));
    }
  }

  /**
   * Change the board config with optimistic concurrency. `mutate` is applied to the
   * latest saved config, so two people editing different parts don't overwrite each other.
   */
  public async updateConfig(board: IBoard, mutate: (cfg: IBoardConfig) => IBoardConfig): Promise<IBoard> {
    const reg = await this.registry();
    const type = await itemEntityType(this.sp, reg);
    let current = board;
    for (let attempt = 0; attempt < 3; attempt++) {
      const next = mutate(JSON.parse(JSON.stringify(current.config)));
      try {
        const etag = await this.sp.merge(`web/lists(guid'${reg}')/items(${board.id})`, {
          __metadata: { type },
          WB_Config: JSON.stringify(next)
        }, current.etag);
        return { ...current, config: next, etag: etag || current.etag };
      } catch (e) {
        if (e instanceof SpError && e.isConflict) {
          current = await this.getBoard(board.id);
          continue;
        }
        throw e;
      }
    }
    throw new Error('Someone else is changing this board. Try again in a moment.');
  }

  public async updateProps(board: IBoard, props: Partial<Pick<IBoard, 'title' | 'description' | 'folder' | 'color' | 'archived' | 'ownerIds' | 'roles'>>): Promise<IBoard> {
    const reg = await this.registry();
    const type = await itemEntityType(this.sp, reg);
    const body: any = { __metadata: { type } };
    if (props.title !== undefined) {
      body.Title = props.title;
    }
    if (props.description !== undefined) {
      body.WB_Description = props.description;
    }
    if (props.folder !== undefined) {
      body.WB_Folder = props.folder;
    }
    if (props.color !== undefined) {
      body.WB_Color = props.color;
    }
    if (props.archived !== undefined) {
      body.WB_Archived = props.archived;
    }
    if (props.ownerIds !== undefined) {
      body.WB_BoardOwnersId = { results: props.ownerIds };
    }
    if (props.roles !== undefined) {
      body.WB_ProjectHeadId = { results: props.roles.head };
      body.WB_ProjectLeadId = { results: props.roles.lead };
      body.WB_ProjectSponsorId = { results: props.roles.sponsor };
    }
    await this.sp.merge(`web/lists(guid'${reg}')/items(${board.id})`, body);
    return this.getBoard(board.id);
  }

  public async addColumn(board: IBoard, type: ColumnType, title: string, afterColumnId?: string): Promise<IBoard> {
    const fields = pickFieldsForNewColumn(board.config, type);
    const col: IBoardColumn = { id: shortId(6), title, type, field: fields.field };
    if (fields.fieldEnd) {
      col.fieldEnd = fields.fieldEnd;
    }
    if (type === 'status') {
      col.labels = STATUS_LABELS.map(l => ({ ...l, id: 'l_' + shortId(5) }));
    }
    if (type === 'dropdown') {
      col.labels = [];
    }
    await this.ensureColumnFields(board.itemsListId, col);
    return this.updateConfig(board, cfg => {
      const idx = afterColumnId ? cfg.columns.map(c => c.id).indexOf(afterColumnId) : -1;
      if (idx >= 0) {
        cfg.columns.splice(idx + 1, 0, col);
      } else {
        cfg.columns.push(col);
      }
      if (type === 'status' && !cfg.kanbanColumnId) {
        cfg.kanbanColumnId = col.id;
      }
      if ((type === 'timeline' || type === 'date') && !cfg.timelineColumnId) {
        cfg.timelineColumnId = col.id;
      }
      return cfg;
    });
  }

  public async renameColumn(board: IBoard, columnId: string, title: string): Promise<IBoard> {
    const col = board.config.columns.filter(c => c.id === columnId)[0];
    const updated = await this.updateConfig(board, cfg => {
      cfg.columns.forEach(c => {
        if (c.id === columnId) {
          c.title = title;
        }
      });
      return cfg;
    });
    if (col && !isFixedField(col.field)) {
      // Keep the SharePoint list readable too. Not critical if it fails.
      renameField(this.sp, board.itemsListId, col.field, title).catch(() => undefined);
    }
    return updated;
  }

  public async deleteColumn(board: IBoard, columnId: string): Promise<IBoard> {
    const col = board.config.columns.filter(c => c.id === columnId)[0];
    const updated = await this.updateConfig(board, cfg => {
      cfg.columns = cfg.columns.filter(c => c.id !== columnId);
      if (cfg.kanbanColumnId === columnId) {
        const s = cfg.columns.filter(c => c.type === 'status')[0];
        cfg.kanbanColumnId = s ? s.id : undefined;
      }
      if (cfg.timelineColumnId === columnId) {
        const t = cfg.columns.filter(c => c.type === 'timeline' || c.type === 'date')[0];
        cfg.timelineColumnId = t ? t.id : undefined;
      }
      return cfg;
    });
    if (col && !isFixedField(col.field)) {
      await deleteField(this.sp, board.itemsListId, col.field);
      if (col.fieldEnd) {
        await deleteField(this.sp, board.itemsListId, col.fieldEnd);
      }
    }
    return updated;
  }

  public async saveLabels(board: IBoard, columnId: string, labels: ILabel[]): Promise<IBoard> {
    return this.updateConfig(board, cfg => {
      cfg.columns.forEach(c => {
        if (c.id === columnId) {
          c.labels = labels;
        }
      });
      return cfg;
    });
  }

  public async saveGroups(board: IBoard, groups: IBoardGroup[]): Promise<IBoard> {
    return this.updateConfig(board, cfg => {
      cfg.groups = groups;
      return cfg;
    });
  }

  public async moveColumn(board: IBoard, columnId: string, toIndex: number): Promise<IBoard> {
    return this.updateConfig(board, cfg => {
      const from = cfg.columns.map(c => c.id).indexOf(columnId);
      if (from < 0) {
        return cfg;
      }
      const [col] = cfg.columns.splice(from, 1);
      cfg.columns.splice(Math.max(0, Math.min(toIndex, cfg.columns.length)), 0, col);
      return cfg;
    });
  }

  /** Move both lists and the registry row to the recycle bin. */
  public async deleteBoard(board: IBoard): Promise<void> {
    await recycleList(this.sp, board.itemsListId);
    await recycleList(this.sp, board.updatesListId);
    const reg = await this.registry();
    await this.sp.post(`web/lists(guid'${reg}')/items(${board.id})/recycle`);
  }

  /**
   * Older boards have no WB_Links field. Someone who can manage the board adds it once;
   * the flag in the config then tells everyone to read it.
   */
  public async enableLinks(board: IBoard): Promise<IBoard> {
    await ensureField(this.sp, board.itemsListId, F.Links, fieldXml('longtext', F.Links, 'Linked files'));
    return this.updateConfig(board, cfg => ({ ...cfg, linksField: true }));
  }

  /**
   * On a private board, give people in project roles read access so they can follow it.
   * People who already have access (`skipIds`, or listed members) keep their access.
   */
  public async grantRoleViewers(board: IBoard, skipIds: number[] = []): Promise<void> {
    if (board.privacy !== 'Private') {
      return;
    }
    const members = await this.listMembers(board).catch(() => [] as IBoardMember[]);
    const has = skipIds.concat(members.map(m => m.principalId));
    const roleIds = board.roles.head.concat(board.roles.lead, board.roles.sponsor);
    for (const id of roleIds.filter((x, i) => roleIds.indexOf(x) === i && has.indexOf(x) < 0)) {
      await this.grant(board, id, 'Viewer', false);
    }
  }

  /* ---------- Permissions ---------- */

  private async roleDefId(roleType: number): Promise<number> {
    if (!this.roleDefIds[roleType]) {
      const def = await this.sp.get<{ Id: number }>(`web/roledefinitions/getbytype(${roleType})?$select=Id`);
      this.roleDefIds[roleType] = def.Id;
    }
    return this.roleDefIds[roleType];
  }

  /** The three securable objects of a board. */
  private async scopes(board: IBoard): Promise<{ path: string; kind: 'items' | 'updates' | 'registry' }[]> {
    const reg = await this.registry();
    return [
      { path: `web/lists(guid'${board.itemsListId}')`, kind: 'items' },
      { path: `web/lists(guid'${board.updatesListId}')`, kind: 'updates' },
      { path: `web/lists(guid'${reg}')/items(${board.id})`, kind: 'registry' }
    ];
  }

  /** Role on each scope: members contribute to items and updates but only read the registry row. */
  private roleFor(role: RoleName, kind: 'items' | 'updates' | 'registry'): number {
    if (role === 'Owner') {
      return kind === 'registry' ? ROLE_TYPE.Member : ROLE_TYPE.Owner;
    }
    if (role === 'Member') {
      return kind === 'registry' ? ROLE_TYPE.Viewer : ROLE_TYPE.Member;
    }
    return ROLE_TYPE.Viewer;
  }

  public async makePrivate(board: IBoard, meId: number, memberIds: number[]): Promise<IBoard> {
    const scopes = await this.scopes(board);
    const ownersGroup = await this.sp.get<{ Id?: number }>('web/AssociatedOwnerGroup?$select=Id').catch(() => null);
    const fullId = await this.roleDefId(ROLE_TYPE.Full);
    for (const scope of scopes) {
      await this.sp.post(`${scope.path}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`);
      if (ownersGroup && ownersGroup.Id) {
        await this.sp.post(`${scope.path}/roleassignments/addroleassignment(principalid=${ownersGroup.Id},roledefid=${fullId})`);
      }
    }
    // Add, don't replace: removing our own assignment first could lock us out of the scope.
    await this.grant(board, meId, 'Owner', false);
    for (const id of memberIds) {
      if (id !== meId) {
        await this.grant(board, id, 'Member');
      }
    }
    const reg = await this.registry();
    const type = await itemEntityType(this.sp, reg);
    await this.sp.merge(`web/lists(guid'${reg}')/items(${board.id})`, { __metadata: { type }, WB_Privacy: 'Private' });
    return this.getBoard(board.id);
  }

  public async makeMain(board: IBoard): Promise<IBoard> {
    const scopes = await this.scopes(board);
    for (const scope of scopes) {
      await this.sp.post(`${scope.path}/resetroleinheritance`);
    }
    const reg = await this.registry();
    const type = await itemEntityType(this.sp, reg);
    await this.sp.merge(`web/lists(guid'${reg}')/items(${board.id})`, { __metadata: { type }, WB_Privacy: 'Main' });
    return this.getBoard(board.id);
  }

  /** Give a person a role on the board. `replace` removes their existing role first, so Member to Viewer takes effect. */
  public async grant(board: IBoard, principalId: number, role: RoleName, replace: boolean = true): Promise<void> {
    const scopes = await this.scopes(board);
    for (const scope of scopes) {
      if (replace) {
        await this.sp.remove(`${scope.path}/roleassignments/getbyprincipalid(${principalId})`).catch(() => undefined);
      }
      const defId = await this.roleDefId(this.roleFor(role, scope.kind));
      await this.sp.post(`${scope.path}/roleassignments/addroleassignment(principalid=${principalId},roledefid=${defId})`);
    }
  }

  public async revoke(board: IBoard, principalId: number): Promise<void> {
    const scopes = await this.scopes(board);
    for (const scope of scopes) {
      await this.sp.remove(`${scope.path}/roleassignments/getbyprincipalid(${principalId})`).catch(() => undefined);
    }
  }

  public async listMembers(board: IBoard): Promise<IBoardMember[]> {
    const rows = await this.sp.get<{ value: any[] }>(
      `web/lists(guid'${board.itemsListId}')/roleassignments?$expand=Member,RoleDefinitionBindings&$select=PrincipalId,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name`
    );
    return rows.value.map(r => ({
      principalId: r.PrincipalId,
      title: (r.Member && r.Member.Title) || 'Unknown',
      isGroup: !!(r.Member && r.Member.PrincipalType !== 1),
      roles: (r.RoleDefinitionBindings || []).map((d: any) => d.Name)
    }));
  }

  /** Does the current user's copy of the board list still exist? (It may have been deleted elsewhere.) */
  public async boardExists(board: IBoard): Promise<boolean> {
    try {
      await this.sp.get(`web/lists(guid'${board.itemsListId}')?$select=Id`);
      return true;
    } catch (e) {
      if (e instanceof SpError && e.isNotFound) {
        return false;
      }
      throw e;
    }
  }
}
