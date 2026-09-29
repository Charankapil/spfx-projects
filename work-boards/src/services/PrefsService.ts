import { SpClient } from './SpClient';
import { findList, itemEntityType } from './lists';
import { LIST_PREFS } from './Provisioner';
import { IUserPrefs } from '../models/types';

const EMPTY: IUserPrefs = { favourites: [], recent: [] };
const MAX_RECENT = 8;

/**
 * Favourites and recent boards, kept in WB_UserPrefs. That list only lets people
 * read and edit their own rows, so the first row returned is always the user's own.
 */
export class PrefsService {
  private readonly sp: SpClient;
  private listId: string | null = null;
  private rowId: number | null = null;
  private prefs: IUserPrefs = EMPTY;

  constructor(sp: SpClient) {
    this.sp = sp;
  }

  public async load(): Promise<IUserPrefs> {
    const list = await findList(this.sp, LIST_PREFS);
    if (!list) {
      return EMPTY;
    }
    this.listId = list.Id;
    const rows = await this.sp.get<{ value: { Id: number; WB_Prefs: string | null }[] }>(
      `web/lists(guid'${list.Id}')/items?$select=Id,WB_Prefs&$top=1&$orderby=Id`
    );
    const row = rows.value[0];
    if (row) {
      this.rowId = row.Id;
      try {
        const parsed = JSON.parse(row.WB_Prefs || '{}');
        this.prefs = {
          favourites: Array.isArray(parsed.favourites) ? parsed.favourites : [],
          recent: Array.isArray(parsed.recent) ? parsed.recent : []
        };
      } catch {
        this.prefs = EMPTY;
      }
    }
    return this.prefs;
  }

  public get current(): IUserPrefs {
    return this.prefs;
  }

  public async toggleFavourite(boardId: number): Promise<IUserPrefs> {
    const favs = this.prefs.favourites.slice();
    const i = favs.indexOf(boardId);
    if (i >= 0) {
      favs.splice(i, 1);
    } else {
      favs.push(boardId);
    }
    return this.save({ ...this.prefs, favourites: favs });
  }

  public async touchRecent(boardId: number): Promise<IUserPrefs> {
    const recent = [boardId].concat(this.prefs.recent.filter(id => id !== boardId)).slice(0, MAX_RECENT);
    if (recent.join(',') === this.prefs.recent.join(',')) {
      return this.prefs;
    }
    return this.save({ ...this.prefs, recent });
  }

  private async save(prefs: IUserPrefs): Promise<IUserPrefs> {
    this.prefs = prefs;
    if (!this.listId) {
      return prefs;
    }
    const type = await itemEntityType(this.sp, this.listId);
    const body = { __metadata: { type }, WB_Prefs: JSON.stringify(prefs) };
    if (this.rowId) {
      await this.sp.merge(`web/lists(guid'${this.listId}')/items(${this.rowId})`, body);
    } else {
      const created = await this.sp.post<{ Id: number }>(`web/lists(guid'${this.listId}')/items`, { ...body, Title: 'prefs' });
      this.rowId = created && created.Id ? created.Id : null;
    }
    return prefs;
  }
}
