import { entityOf, HttpError, odataString, rowsOf, SPClient, trimSlash, encodePathLiteral } from '../core/SPClient';
import {
  IAssignment,
  IBannerLock,
  IPlannedChange,
  ISavedAssignment,
  isSystemBinding,
  LockScope,
  planReadOnly,
  planRestore
} from '../../../shared/readOnlyLocks';

/** Component id of the read-only banner application customizer shipped in this package. */
export const BANNER_COMPONENT_ID = '67006b0f-75c3-43e7-b6c0-d6519821241e';
export const LOCK_FILE = 'admin-center-locks.json';
export const MAX_MESSAGE = 500;

export type ObjectKind = 'web' | 'list' | 'item';

export interface ILockObject {
  kind: ObjectKind;
  title: string;
  listId?: string;
  itemId?: number;
}

export interface ILockedObject extends ILockObject {
  /** Inherited permissions before the lock (unlock re-inherits). */
  inherited: boolean;
  /** Original permission levels (non-system) per principal, when the object had its own. */
  saved: ISavedAssignment[];
}

export interface ILockRecord extends IBannerLock {
  by: string;
  keepOwners: boolean;
  status: 'locking' | 'locked';
  objects: ILockedObject[];
}

interface ILockStore {
  v: 1;
  locks: ILockRecord[];
}

export interface ILockTarget {
  scope: LockScope;
  title: string;
  /** Server-relative path shown in the banner match (web, library root folder or folder). */
  path: string;
  listId?: string;
  /** Folder path (server-relative) when scope is 'folder'. */
  folderPath?: string;
}

export interface ILockOptions {
  keepOwners: boolean;
  /** Site scope: also lock lists and libraries that have their own permissions. */
  includeUniqueLists: boolean;
}

export interface IPreviewObject {
  object: ILockObject;
  inherited: boolean;
  changes: IPlannedChange[];
}

export interface IPreview {
  objects: IPreviewObject[];
  read: { id: number; name: string };
  keepPrincipalIds: number[];
  warnings: string[];
}

function str(v: unknown): string {
  return v === null || v === undefined ? '' : String(v);
}

/**
 * Applies and removes permission-based read-only locks as the signed-in user (needs
 * Manage Permissions on the object, i.e. site owner or site collection admin).
 * Requests: a handful per object plus two per principal that changes, all paced by SPClient.
 */
export class LockService {
  constructor(private client: SPClient) {}

  private api(webUrl: string, path: string): string {
    return `${trimSlash(webUrl)}/_api/${path}`;
  }

  private base(webUrl: string, o: ILockObject): string {
    if (o.kind === 'list') {
      return this.api(webUrl, `web/lists(guid'${odataString(o.listId || '')}')`);
    }
    if (o.kind === 'item') {
      return this.api(webUrl, `web/lists(guid'${odataString(o.listId || '')}')/items(${o.itemId})`);
    }
    return this.api(webUrl, 'web');
  }

  private serverRelative(webUrl: string): string {
    const m = /^https:\/\/[^/]+(\/[^?#]*)?/i.exec(trimSlash(webUrl));
    return m && m[1] ? decodeURIComponent(trimSlash(m[1])) : '';
  }

  // ---- reading permissions ---------------------------------------------------

  public async hasUnique(webUrl: string, o: ILockObject): Promise<boolean> {
    const j = await this.client.get<unknown>(`${this.base(webUrl, o)}?$select=HasUniqueRoleAssignments`, { cacheTtlMs: 0 });
    return !!entityOf<{ HasUniqueRoleAssignments?: boolean }>(j).HasUniqueRoleAssignments;
  }

  public async getAssignments(webUrl: string, o: ILockObject): Promise<IAssignment[]> {
    const j = await this.client.get(
      `${this.base(webUrl, o)}/roleassignments?$select=PrincipalId,Member/Title,Member/LoginName,Member/PrincipalType,RoleDefinitionBindings/Id,RoleDefinitionBindings/Name,RoleDefinitionBindings/RoleTypeKind,RoleDefinitionBindings/Hidden&$expand=Member,RoleDefinitionBindings&$top=2000`,
      { cacheTtlMs: 0 }
    );
    return rowsOf<{ [k: string]: unknown }>(j).map((r) => {
      const m = (r.Member || {}) as { [k: string]: unknown };
      return {
        principalId: Number(r.PrincipalId),
        title: str(m.Title),
        login: str(m.LoginName),
        principalType: Number(m.PrincipalType) || 0,
        bindings: rowsOf<{ [k: string]: unknown }>(r.RoleDefinitionBindings).map((b) => ({ id: Number(b.Id), name: str(b.Name), kind: Number(b.RoleTypeKind) || 0, hidden: !!b.Hidden }))
      };
    });
  }

  public async readRole(webUrl: string): Promise<{ id: number; name: string }> {
    const r = entityOf<{ Id?: number; Name?: string }>(await this.client.get(this.api(webUrl, 'web/roledefinitions/getbytype(2)?$select=Id,Name')));
    if (!r || !r.Id) {
      throw new Error('The Read permission level was not found on this site.');
    }
    return { id: Number(r.Id), name: str(r.Name) || 'Read' };
  }

  public async ownersGroupId(webUrl: string): Promise<number | undefined> {
    try {
      const g = entityOf<{ Id?: number }>(await this.client.get(this.api(webUrl, 'web/AssociatedOwnerGroup?$select=Id')));
      return g && g.Id ? Number(g.Id) : undefined;
    } catch {
      return undefined;
    }
  }

  public async isSiteAdmin(webUrl: string): Promise<boolean> {
    const u = entityOf<{ IsSiteAdmin?: boolean }>(await this.client.get(this.api(webUrl, 'web/currentuser?$select=IsSiteAdmin'), { cacheTtlMs: 0 }));
    return !!(u && u.IsSiteAdmin);
  }

  public async listFolders(webUrl: string, folderPath: string): Promise<Array<{ name: string; path: string }>> {
    const j = await this.client.get(this.api(webUrl, `web/GetFolderByServerRelativeUrl('${encodePathLiteral(folderPath)}')/Folders?$select=Name,ServerRelativeUrl&$top=500`), { cacheTtlMs: 0 });
    return rowsOf<{ Name?: string; ServerRelativeUrl?: string }>(j)
      .filter((f) => f.Name && f.Name !== 'Forms' && str(f.Name).charAt(0) !== '_')
      .map((f) => ({ name: str(f.Name), path: str(f.ServerRelativeUrl) }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }

  private async folderItem(webUrl: string, folderPath: string): Promise<{ listId: string; itemId: number }> {
    const j = entityOf<{ Id?: number; ParentList?: { Id?: string } }>(
      await this.client.get(this.api(webUrl, `web/GetFolderByServerRelativeUrl('${encodePathLiteral(folderPath)}')/ListItemAllFields?$select=Id,ParentList/Id&$expand=ParentList`), { cacheTtlMs: 0 })
    );
    if (!j || !j.Id || !j.ParentList || !j.ParentList.Id) {
      throw new Error('That folder could not be found (only folders inside a library or list can be locked).');
    }
    return { listId: String(j.ParentList.Id), itemId: Number(j.Id) };
  }

  // ---- lock store ------------------------------------------------------------

  public async loadLocks(webUrl: string): Promise<ILockRecord[]> {
    const url = this.api(webUrl, `web/GetFileByServerRelativeUrl('${encodePathLiteral(`${this.serverRelative(webUrl)}/SiteAssets/${LOCK_FILE}`)}')/$value`);
    try {
      const raw = await this.client.get<ILockStore>(url, { cacheTtlMs: 0 });
      return raw && Array.isArray(raw.locks) ? raw.locks : [];
    } catch (e) {
      if (e instanceof HttpError && e.status === 404) {
        return [];
      }
      throw e;
    }
  }

  private async saveLocks(webUrl: string, locks: ILockRecord[]): Promise<void> {
    const folder = `${this.serverRelative(webUrl)}/SiteAssets`;
    const write = (): Promise<unknown> =>
      this.client.post(this.api(webUrl, `web/GetFolderByServerRelativeUrl('${encodePathLiteral(folder)}')/Files/add(url='${LOCK_FILE}',overwrite=true)`), JSON.stringify({ v: 1, locks }), { raw: true }, webUrl);
    try {
      await write();
    } catch (e) {
      if (!(e instanceof HttpError) || e.status !== 404) {
        throw e;
      }
      // No Site Assets library on this web yet: create it, then save.
      await this.client.post(this.api(webUrl, 'web/lists/EnsureSiteAssetsLibrary'), undefined, {}, webUrl);
      await write();
    }
  }

  // ---- banner (application customizer registered on the web) ------------------

  public async syncBanner(webUrl: string, locks: ILockRecord[]): Promise<void> {
    const j = await this.client.get(this.api(webUrl, `web/UserCustomActions?$select=Id&$filter=ClientSideComponentId eq guid'${BANNER_COMPONENT_ID}'`), { cacheTtlMs: 0 });
    const existing = rowsOf<{ Id?: string }>(j).map((a) => str(a.Id));
    const banner: IBannerLock[] = locks.map((l) => ({ id: l.id, scope: l.scope, path: l.path, title: l.title, message: l.message, at: l.at }));
    if (!banner.length) {
      for (const id of existing) {
        await this.client.post(this.api(webUrl, `web/UserCustomActions('${odataString(id)}')`), undefined, { method: 'DELETE' }, webUrl);
      }
      return;
    }
    const props = JSON.stringify({ locks: banner });
    if (existing.length) {
      await this.client.post(this.api(webUrl, `web/UserCustomActions('${odataString(existing[0])}')`), { ClientSideComponentProperties: props }, { method: 'MERGE' }, webUrl);
      for (const extra of existing.slice(1)) {
        await this.client.post(this.api(webUrl, `web/UserCustomActions('${odataString(extra)}')`), undefined, { method: 'DELETE' }, webUrl);
      }
    } else {
      await this.client.post(
        this.api(webUrl, 'web/UserCustomActions'),
        {
          Title: 'SharePoint Admin Center read-only banner',
          Name: 'SPAdminCenterReadOnlyBanner',
          Location: 'ClientSideExtension.ApplicationCustomizer',
          ClientSideComponentId: BANNER_COMPONENT_ID,
          ClientSideComponentProperties: props,
          Sequence: 10
        },
        {},
        webUrl
      );
    }
  }

  // ---- plan / lock / unlock ---------------------------------------------------------

  private async resolveObjects(webUrl: string, target: ILockTarget, options: ILockOptions, uniqueLists: Array<{ id: string; title: string }>): Promise<ILockObject[]> {
    if (target.scope === 'list') {
      return [{ kind: 'list', listId: target.listId, title: target.title }];
    }
    if (target.scope === 'folder') {
      const f = await this.folderItem(webUrl, target.folderPath || target.path);
      return [{ kind: 'item', listId: f.listId, itemId: f.itemId, title: target.title }];
    }
    const objects: ILockObject[] = [{ kind: 'web', title: target.title }];
    if (options.includeUniqueLists) {
      uniqueLists.forEach((l) => objects.push({ kind: 'list', listId: l.id, title: l.title }));
    }
    return objects;
  }

  /** What a lock would change, without changing anything. */
  public async preview(webUrl: string, target: ILockTarget, options: ILockOptions, uniqueLists: Array<{ id: string; title: string }>): Promise<IPreview> {
    const read = await this.readRole(webUrl);
    const owners = options.keepOwners ? await this.ownersGroupId(webUrl) : undefined;
    const keep = owners ? [owners] : [];
    const objects: IPreviewObject[] = [];
    for (const o of await this.resolveObjects(webUrl, target, options, uniqueLists)) {
      const unique = await this.hasUnique(webUrl, o);
      const assignments = await this.getAssignments(webUrl, o);
      objects.push({ object: o, inherited: !unique, changes: planReadOnly(assignments, read.id, read.name, keep) });
    }
    const warnings = [
      target.scope === 'web'
        ? 'Folders and files inside libraries that have their own permissions keep them. Subsites with their own permissions are not changed.'
        : 'Subfolders and files inside it that have their own permissions keep them.',
      'Site collection admins always keep full control. On a Microsoft 365 group-connected site that includes the group owners.'
    ];
    return { objects, read, keepPrincipalIds: keep, warnings };
  }

  private async applyChanges(webUrl: string, o: ILockObject, changes: IPlannedChange[], tick: () => void): Promise<void> {
    const ra = `${this.base(webUrl, o)}/roleassignments`;
    for (const c of changes) {
      for (const add of c.add) {
        await this.client.post(`${ra}/addroleassignment(principalid=${c.principalId},roledefid=${add})`, undefined, {}, webUrl);
      }
      for (const rem of c.remove) {
        await this.client.post(`${ra}/removeroleassignment(principalid=${c.principalId},roledefid=${rem})`, undefined, {}, webUrl);
      }
      tick();
    }
  }

  public async lock(
    webUrl: string,
    target: ILockTarget,
    message: string,
    options: ILockOptions,
    uniqueLists: Array<{ id: string; title: string }>,
    by: string,
    onProgress?: (text: string) => void
  ): Promise<ILockRecord> {
    const say = (t: string): void => {
      if (onProgress) {
        onProgress(t);
      }
    };
    const locks = await this.loadLocks(webUrl);
    const path = trimSlash(target.path);
    if (locks.some((l) => l.path.toLowerCase() === path.toLowerCase())) {
      throw new Error('That is already locked. Unlock it first to change the lock.');
    }
    say('Checking current permissions…');
    const plan = await this.preview(webUrl, target, options, uniqueLists);
    const record: ILockRecord = {
      id: `lock-${Date.now()}`,
      scope: target.scope,
      path,
      title: target.title,
      message: message.trim().substring(0, MAX_MESSAGE),
      at: Math.floor(Date.now() / 1000),
      by,
      keepOwners: options.keepOwners,
      status: 'locking',
      objects: []
    };
    // Save what each object looks like BEFORE touching it, so an interrupted lock can still be undone.
    for (const p of plan.objects) {
      const before = p.inherited ? [] : await this.getAssignments(webUrl, p.object);
      record.objects.push({
        ...p.object,
        inherited: p.inherited,
        saved: before.map((a) => ({ principalId: a.principalId, title: a.title, roleDefIds: a.bindings.filter((b) => !isSystemBinding(b)).map((b) => b.id) }))
      });
    }
    await this.saveLocks(webUrl, locks.concat(record));

    let done = 0;
    for (const p of plan.objects) {
      say(`Locking ${p.object.title}…`);
      if (p.inherited) {
        await this.client.post(`${this.base(webUrl, p.object)}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)`, undefined, {}, webUrl);
      }
      // Re-plan against the object's own (copied) permissions.
      const current = await this.getAssignments(webUrl, p.object);
      const changes = planReadOnly(current, plan.read.id, plan.read.name, plan.keepPrincipalIds);
      await this.applyChanges(webUrl, p.object, changes, () => say(`Locking ${p.object.title}… ${++done} permission change(s)`));
    }

    record.status = 'locked';
    const after = (await this.loadLocks(webUrl)).filter((l) => l.id !== record.id).concat(record);
    await this.saveLocks(webUrl, after);
    say('Showing the banner…');
    await this.syncBanner(webUrl, after);
    return record;
  }

  public async unlock(webUrl: string, lockId: string, onProgress?: (text: string) => void): Promise<void> {
    const locks = await this.loadLocks(webUrl);
    const record = locks.filter((l) => l.id === lockId)[0];
    if (!record) {
      throw new Error('That lock no longer exists.');
    }
    const read = await this.readRole(webUrl);
    for (const o of record.objects.slice().reverse()) {
      if (onProgress) {
        onProgress(`Restoring ${o.title}…`);
      }
      if (o.inherited) {
        await this.client.post(`${this.base(webUrl, o)}/resetroleinheritance`, undefined, {}, webUrl);
        continue;
      }
      const current = await this.getAssignments(webUrl, o);
      await this.applyChanges(webUrl, o, planRestore(o.saved, current, read.id), () => undefined);
    }
    const remaining = (await this.loadLocks(webUrl)).filter((l) => l.id !== lockId);
    await this.saveLocks(webUrl, remaining);
    await this.syncBanner(webUrl, remaining);
  }

  /** Changes only the banner message of an existing lock. */
  public async updateMessage(webUrl: string, lockId: string, message: string): Promise<void> {
    const locks = await this.loadLocks(webUrl);
    locks.forEach((l) => {
      if (l.id === lockId) {
        l.message = message.trim().substring(0, MAX_MESSAGE);
      }
    });
    await this.saveLocks(webUrl, locks);
    await this.syncBanner(webUrl, locks);
  }
}
