import { SpClient, SpError, mapLimit } from './SpClient';
import { itemEntityType } from './lists';
import { IBoard, IBoardColumn, IWorkItem, CellValue, IAttachment, IActivityEntry, IMyWorkRow, IPerson } from '../models/types';
import { F, itemQueryParts, readItem, writeCell } from '../engine/fieldMap';
import { diffVersions, IVersionRow } from '../engine/activity';
import { fromSpDate } from '../engine/dates';

/* eslint-disable @typescript-eslint/no-explicit-any */

export interface IListState {
  lastModified: string;
  lastDeleted: string;
}

export interface INewItem {
  title: string;
  groupId: string;
  sortOrder: number;
  parentId?: number | null;
  values?: { [columnId: string]: CellValue };
}

/** Items of one board list. */
export class ItemService {
  private readonly sp: SpClient;

  constructor(sp: SpClient) {
    this.sp = sp;
  }

  private listPath(board: IBoard): string {
    return `web/lists(guid'${board.itemsListId}')`;
  }

  private query(board: IBoard): string {
    const { select, expand } = itemQueryParts(board.config.columns);
    return `$select=${select.join(',')}&$expand=${expand.join(',')}`;
  }

  public async loadItems(board: IBoard): Promise<IWorkItem[]> {
    const rows = await this.sp.getAll<any>(`${this.listPath(board)}/items?${this.query(board)}&$top=5000`);
    return rows.map(r => readItem(board.config.columns, r));
  }

  public async getItem(board: IBoard, id: number): Promise<IWorkItem> {
    const row = await this.sp.get<any>(`${this.listPath(board)}/items(${id})?${this.query(board)}`);
    return readItem(board.config.columns, row);
  }

  /** Cheap check used by the change poller. */
  public async listState(board: IBoard): Promise<IListState> {
    const info = await this.sp.get<any>(`${this.listPath(board)}?$select=LastItemUserModifiedDate,LastItemDeletedDate`);
    return { lastModified: info.LastItemUserModifiedDate || '', lastDeleted: info.LastItemDeletedDate || '' };
  }

  public async listPermissions(board: IBoard): Promise<{ High: string; Low: string }> {
    return this.sp.get(`${this.listPath(board)}/EffectiveBasePermissions`);
  }

  private cellPayload(columns: IBoardColumn[], values: { [columnId: string]: CellValue }): { [k: string]: any } {
    let body: { [k: string]: any } = {};
    columns.forEach(col => {
      if (Object.prototype.hasOwnProperty.call(values, col.id)) {
        body = { ...body, ...writeCell(col, values[col.id]) };
      }
    });
    return body;
  }

  public async createItem(board: IBoard, item: INewItem): Promise<IWorkItem> {
    const type = await itemEntityType(this.sp, board.itemsListId);
    const body = {
      __metadata: { type },
      Title: item.title,
      [F.GroupId]: item.groupId,
      [F.SortOrder]: item.sortOrder,
      [F.ParentId]: item.parentId || null,
      ...this.cellPayload(board.config.columns, item.values || {})
    };
    const created = await this.sp.post<any>(`${this.listPath(board)}/items`, body);
    return this.getItem(board, created.Id);
  }

  /**
   * Save changed values. If someone else saved the item in the meantime (412), reload it and
   * apply only our changed fields on top, so edits to different fields both survive.
   */
  public async updateItem(
    board: IBoard,
    item: IWorkItem,
    changes: { title?: string; groupId?: string; sortOrder?: number; parentId?: number | null; values?: { [columnId: string]: CellValue } }
  ): Promise<{ etag: string; conflicted: boolean }> {
    const type = await itemEntityType(this.sp, board.itemsListId);
    const body: { [k: string]: any } = { __metadata: { type } };
    if (changes.title !== undefined) {
      body.Title = changes.title;
    }
    if (changes.groupId !== undefined) {
      body[F.GroupId] = changes.groupId;
    }
    if (changes.sortOrder !== undefined) {
      body[F.SortOrder] = changes.sortOrder;
    }
    if (changes.parentId !== undefined) {
      body[F.ParentId] = changes.parentId;
    }
    Object.assign(body, this.cellPayload(board.config.columns, changes.values || {}));
    const path = `${this.listPath(board)}/items(${item.id})`;
    try {
      const etag = await this.sp.merge(path, body, item.etag || '*');
      return { etag, conflicted: false };
    } catch (e) {
      if (e instanceof SpError && e.isConflict) {
        const etag = await this.sp.merge(path, body, '*');
        return { etag, conflicted: true };
      }
      throw e;
    }
  }

  /** Move to the SharePoint recycle bin (restorable). */
  public async recycleItem(board: IBoard, id: number): Promise<void> {
    await this.sp.post(`${this.listPath(board)}/items(${id})/recycle`);
  }

  /* ---------- Files (item attachments) ---------- */

  public async listAttachments(board: IBoard, id: number): Promise<IAttachment[]> {
    const res = await this.sp.get<{ value: { FileName: string; ServerRelativeUrl: string }[] }>(
      `${this.listPath(board)}/items(${id})/AttachmentFiles`
    );
    const origin = this.sp.webUrl.replace(/^(https?:\/\/[^/]+).*$/, '$1');
    return res.value.map(a => ({ fileName: a.FileName, url: origin + a.ServerRelativeUrl }));
  }

  public async addAttachment(board: IBoard, id: number, file: File): Promise<void> {
    const content = await file.arrayBuffer();
    const name = file.name.replace(/'/g, "''");
    await this.sp.postBinary(`${this.listPath(board)}/items(${id})/AttachmentFiles/add(FileName='${encodeURIComponent(name)}')`, content);
  }

  public async deleteAttachment(board: IBoard, id: number, fileName: string): Promise<void> {
    const name = fileName.replace(/'/g, "''");
    const path = `${this.listPath(board)}/items(${id})/AttachmentFiles/getByFileName('${encodeURIComponent(name)}')`;
    try {
      // Recycle so the file can be restored.
      await this.sp.post(`${path}/recycleObject`);
    } catch (e) {
      if (e instanceof SpError && (e.isNotFound || e.status === 400)) {
        await this.sp.remove(path);
      } else {
        throw e;
      }
    }
  }

  /* ---------- Activity (version history) ---------- */

  public async activity(board: IBoard, id: number): Promise<IActivityEntry[]> {
    const res = await this.sp.get<{ value: IVersionRow[] }>(`${this.listPath(board)}/items(${id})/versions`);
    return diffVersions(board.config, res.value);
  }

  /* ---------- My Work ---------- */

  /**
   * Items where the user is in the board's main People column (WB_Owner), across boards.
   * Filters on the server; a list over the 5,000-item threshold falls back to a paged scan.
   */
  public async myWork(boards: IBoard[], me: IPerson): Promise<IMyWorkRow[]> {
    const perBoard = await mapLimit(boards, 4, async board => {
      const ownerCol = board.config.columns.filter(c => c.field === F.Owner)[0];
      if (!ownerCol) {
        return [] as IMyWorkRow[];
      }
      const statusCol = board.config.columns.filter(c => c.field === F.Status)[0];
      const dueCol = board.config.columns.filter(c => c.field === F.DueDate || c.fieldEnd === F.DueDate)[0];
      const select = `Id,Title,${F.GroupId},${F.SortOrder},${F.ParentId},${F.Status},${F.StartDate},${F.DueDate},Created,Modified,Attachments,${F.Owner}/Id,${F.Owner}/Title,${F.Owner}/EMail`;
      const base = `web/lists(guid'${board.itemsListId}')/items?$select=${select}&$expand=${F.Owner}`;
      let rows: any[];
      try {
        rows = await this.sp.getAll<any>(`${base}&$filter=${F.Owner}/Id eq ${me.id}&$top=500`, 2000);
      } catch (e) {
        if (e instanceof SpError && (e.isThreshold || e.status === 500)) {
          rows = (await this.sp.getAll<any>(`${base}&$top=5000`)).filter(r => {
            const owners = r[F.Owner];
            return Array.isArray(owners) && owners.some((o: any) => o.Id === me.id);
          });
        } else if (e instanceof SpError && (e.isNotFound || e.isAccessDenied)) {
          return [] as IMyWorkRow[];
        } else {
          throw e;
        }
      }
      return rows.map(r => {
        const item = readItem([ownerCol], r);
        const statusText: string | null = statusCol ? r[F.Status] || null : null;
        const status = statusCol && statusText ? (statusCol.labels || []).filter(l => l.text === statusText)[0] || { id: '', text: statusText, color: '#c4c4c4' } : null;
        return {
          board,
          item,
          status,
          due: dueCol ? fromSpDate(r[F.DueDate]) : null
        } as IMyWorkRow;
      });
    });
    const all: IMyWorkRow[] = [];
    perBoard.forEach(rows => rows.forEach(r => all.push(r)));
    return all;
  }

  /** Status counts for a board's main Status column, for board cards on the home page. */
  public async statusCounts(board: IBoard): Promise<{ [label: string]: number }> {
    const rows = await this.sp.getAll<any>(`${this.listPath(board)}/items?$select=${F.Status},${F.ParentId}&$top=5000`, 20000);
    const counts: { [label: string]: number } = {};
    rows.forEach(r => {
      if (r[F.ParentId]) {
        return;
      }
      const s = r[F.Status] || '';
      counts[s] = (counts[s] || 0) + 1;
    });
    return counts;
  }
}
