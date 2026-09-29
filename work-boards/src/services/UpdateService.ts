import { SpClient } from './SpClient';
import { itemEntityType } from './lists';
import { IBoard, IUpdate } from '../models/types';
import { UF } from '../engine/fieldMap';
import { mentionIds } from '../engine/mentions';

/* eslint-disable @typescript-eslint/no-explicit-any */

function ids(raw: any): number[] {
  if (Array.isArray(raw)) {
    return raw;
  }
  return raw && Array.isArray(raw.results) ? raw.results : [];
}

function readUpdate(row: any): IUpdate {
  return {
    id: row.Id,
    itemId: row[UF.ItemId],
    parentId: row[UF.ParentId] || null,
    body: row[UF.Body] || '',
    author: row.Author ? { id: row.Author.Id, title: row.Author.Title, email: row.Author.EMail || '' } : null,
    created: row.Created,
    mentionIds: ids(row[UF.Mentions + 'Id']),
    likedByIds: ids(row[UF.LikedBy + 'Id']),
    etag: row['odata.etag'] || '*'
  };
}

const SELECT = `Id,${UF.ItemId},${UF.ParentId},${UF.Body},${UF.Mentions}Id,${UF.LikedBy}Id,Created,Author/Id,Author/Title,Author/EMail`;

/** Updates (comments) live in each board's companion list, so they share the board's permissions. */
export class UpdateService {
  private readonly sp: SpClient;

  constructor(sp: SpClient) {
    this.sp = sp;
  }

  private path(board: IBoard): string {
    return `web/lists(guid'${board.updatesListId}')`;
  }

  public async forItem(board: IBoard, itemId: number): Promise<IUpdate[]> {
    const rows = await this.sp.getAll<any>(
      `${this.path(board)}/items?$select=${SELECT}&$expand=Author&$filter=${UF.ItemId} eq ${itemId}&$top=500`
    );
    return rows.map(readUpdate).sort((a, b) => a.created.localeCompare(b.created));
  }

  /** Number of updates per item, for the speech-bubble badges. */
  public async counts(board: IBoard): Promise<{ [itemId: number]: number }> {
    const rows = await this.sp.getAll<any>(`${this.path(board)}/items?$select=${UF.ItemId}&$top=5000`, 50000);
    const out: { [itemId: number]: number } = {};
    rows.forEach(r => {
      const id = r[UF.ItemId];
      out[id] = (out[id] || 0) + 1;
    });
    return out;
  }

  public async add(board: IBoard, itemId: number, body: string, parentId: number | null): Promise<IUpdate> {
    const type = await itemEntityType(this.sp, board.updatesListId);
    const created = await this.sp.post<any>(`${this.path(board)}/items`, {
      __metadata: { type },
      Title: body.substring(0, 100).replace(/\n/g, ' '),
      [UF.ItemId]: itemId,
      [UF.ParentId]: parentId,
      [UF.Body]: body,
      [UF.Mentions + 'Id']: { results: mentionIds(body) }
    });
    const row = await this.sp.get<any>(`${this.path(board)}/items(${created.Id})?$select=${SELECT}&$expand=Author`);
    return readUpdate(row);
  }

  public async toggleLike(board: IBoard, update: IUpdate, meId: number): Promise<IUpdate> {
    const liked = update.likedByIds.indexOf(meId) >= 0;
    const next = liked ? update.likedByIds.filter(id => id !== meId) : update.likedByIds.concat([meId]);
    const type = await itemEntityType(this.sp, board.updatesListId);
    // '*' so two people liking at once both count: we send the full list we saw plus our change.
    await this.sp.merge(`${this.path(board)}/items(${update.id})`, {
      __metadata: { type },
      [UF.LikedBy + 'Id']: { results: next }
    });
    return { ...update, likedByIds: next };
  }

  public async remove(board: IBoard, update: IUpdate): Promise<void> {
    await this.sp.post(`${this.path(board)}/items(${update.id})/recycle`);
  }
}
