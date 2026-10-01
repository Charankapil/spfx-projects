import { IRestoreOptions } from '../models/IRestoreOptions';
import { IRunProgress } from '../models/IRunProgress';
import { EMPTY_STATS, IRunStats } from '../models/IRunReport';
import { IScopeNode } from '../models/IScopeNode';
import { IUniqueObject } from '../models/IUniqueObject';
import { runPool } from './pool';
import { getWebLists, IListInfo, isBadRequest, isContentList, LIST_SELECT, ScopeService } from './ScopeService';
import { HttpError, isCancelled, quoteForUrl, SpRest } from './SpRest';

/**
 * Items are read in ID ranges rather than by following the paging link: ID
 * is always indexed, so "Id ge a and Id le b" stays under the 5,000-item
 * list view threshold in a library of any size, and independent ranges can
 * be fetched in parallel. A range that times out is split in half and
 * retried, at most MAX_SPLITS times.
 */
const ID_RANGE = 5000;
/** A range that times out is split in two, at most this many times over (5,000 -> 2,500 -> 1,250). */
const MAX_SPLITS = 2;

/** After this many failed reads in a row, stop reading the list instead of piling up errors. */
const MAX_CONSECUTIVE_FAILURES = 5;

/**
 * Folders whose direct contents exceed this are not listed through
 * Folder/Files; the walk hands over to the ID-range read instead.
 */
const MAX_FOLDER_ITEMS_FOR_WALK = 5000;

/** The walk's request budget, as a multiple of what reading the whole list by ID ranges would cost. */
const WALK_BUDGET_FACTOR = 3;
const MIN_WALK_BUDGET = 25;

const FOLDER_ITEM_SELECT = 'ListItemAllFields/Id,ListItemAllFields/HasUniqueRoleAssignments';

interface IFolderEntry {
  Name: string;
  ServerRelativeUrl: string;
  ItemCount?: number;
  ListItemAllFields?: { Id?: number; HasUniqueRoleAssignments?: boolean };
}

/** Thrown inside the folder walk to hand over to the ID-range read. */
class WalkAbandoned extends Error {
  constructor(reason: string) {
    super(reason);
    (Object as unknown as { setPrototypeOf: (o: object, p: object) => void }).setPrototypeOf(this, WalkAbandoned.prototype);
  }
}

/** Libraries need no Title (the file name comes from FileRef), which keeps each row light. */
const LIBRARY_SELECT = 'Id,FileRef,FSObjType,HasUniqueRoleAssignments';
const LIST_ITEM_SELECT = 'Id,Title,FileRef,FSObjType,HasUniqueRoleAssignments';

function itemSelect(list: IListInfo): string {
  return list.BaseType === 1 ? LIBRARY_SELECT : LIST_ITEM_SELECT;
}

interface IItemRow {
  Id: number;
  Title?: string;
  FileRef: string;
  FSObjType: number | string;
  HasUniqueRoleAssignments: boolean;
}

interface IWebInfo {
  Title: string;
  ServerRelativeUrl: string;
  HasUniqueRoleAssignments: boolean;
  WebTemplate?: string;
}

/** A part of a list to read: the root folder or a selected folder. */
interface IItemScope {
  path: string;
  depth: number;
  /** Items directly in the folder, when already known (saves the walk a request). */
  itemCount?: number;
}

export interface IScanResult {
  objects: IUniqueObject[];
  stats: IRunStats;
  errors: string[];
}

function message(err: unknown): string {
  return err instanceof Error ? err.message : 'unknown error';
}

/** Folder levels between the list root and the item (a top-level file is 1). */
export function depthBelow(path: string, rootPath: string): number {
  return path.substring(rootPath.length).split('/').filter((s) => s.length > 0).length;
}

/**
 * Only a server-side timeout or list view threshold error (a 5xx) is worth
 * retrying as two smaller ranges. Throttling is already waited out by the
 * REST client and must never be answered with more requests.
 */
function isWorthSplitting(err: unknown): boolean {
  return err instanceof HttpError && ((err.status >= 500 && err.status !== 503) || err.status === 408);
}

/**
 * Finds everything with unique permissions under the selected scopes.
 * Only objects that have unique permissions are kept in memory, so a scan of
 * a library with millions of items costs one small request per 5,000 IDs
 * and memory proportional to what it finds.
 */
export class UniquePermissionScanner {
  private found = new Map<string, IUniqueObject>();
  private visitedWebs = new Set<string>();
  private contentScannedLists = new Set<string>();
  private errors: string[] = [];
  private stats: IRunStats = { ...EMPTY_STATS };
  private progress: IRunProgress;
  private lastReport = 0;

  constructor(
    private rest: SpRest,
    private scopeService: ScopeService,
    private siteServerRelativeUrl: string,
    private options: IRestoreOptions,
    private onProgress: (progress: IRunProgress) => void,
    /** Ranges / folder listings read at once: 1 in gentle mode, 2 in standard. */
    private lanes = 1
  ) {
    this.progress = {
      phase: 'scanning',
      currentItem: '',
      websChecked: 0,
      listsChecked: 0,
      itemsChecked: 0,
      uniqueFound: 0
    };
  }

  public async scan(scopes: IScopeNode[]): Promise<IScanResult> {
    // Broad scopes first, so a library that is also covered by a selected
    // site is read once, as part of the site.
    const rank = (s: IScopeNode): number => (s.kind === 'web' ? 0 : s.kind === 'folder' ? 2 : 1);
    const ordered = [...scopes].sort((a, b) => rank(a) - rank(b));

    // Selected folders of the same list are read in one pass over the list.
    const folderGroups = new Map<string, IScopeNode[]>();
    for (const scope of ordered) {
      this.rest.throwIfCancelled();
      if (scope.kind === 'folder') {
        const groupKey = `${scope.webUrl}|${scope.listId}`;
        folderGroups.set(groupKey, [...(folderGroups.get(groupKey) || []), scope]);
        continue;
      }
      await this.guard(scope.title, scope.serverRelativeUrl, async () => {
        if (scope.kind === 'web') {
          await this.scanWeb(scope.webUrl, true);
        } else {
          const list = await this.getList(scope.webUrl, scope.listId as string);
          await this.scanList(scope.webUrl, list, true);
        }
      });
    }
    const groups: IScopeNode[][] = [];
    folderGroups.forEach((group) => groups.push(group));
    for (const group of groups) {
      await this.guard(group[0].title, group[0].serverRelativeUrl, () => this.scanFolderScopes(group));
    }

    this.report(true);
    const objects: IUniqueObject[] = [];
    this.found.forEach((o) => objects.push(o));
    return { objects, stats: { ...this.stats, uniqueFound: objects.length }, errors: this.errors };
  }

  /** Records a failure against one scope and carries on with the others. */
  private async guard(title: string, url: string, action: () => Promise<void>): Promise<void> {
    try {
      await action();
    } catch (err) {
      if (isCancelled(err)) {
        throw err;
      }
      this.errors.push(`${title} (${url}): ${message(err)}`);
    }
  }

  private report(force = false): void {
    const now = Date.now();
    if (!force && now - this.lastReport < 200) {
      return;
    }
    this.lastReport = now;
    this.progress.websChecked = this.stats.websChecked;
    this.progress.listsChecked = this.stats.listsChecked;
    this.progress.itemsChecked = this.stats.itemsChecked;
    this.progress.uniqueFound = this.found.size;
    this.onProgress({ ...this.progress });
  }

  private add(object: IUniqueObject): void {
    if (!this.found.has(object.key)) {
      this.found.set(object.key, object);
    }
  }

  private async getList(webUrl: string, listId: string): Promise<IListInfo> {
    return this.rest.getJson<IListInfo>(
      `${webUrl}/_api/web/lists(guid'${listId}')?$select=${LIST_SELECT}&$expand=RootFolder`
    );
  }

  private async scanWeb(webUrl: string, isSelected: boolean): Promise<void> {
    const key = webUrl.toLowerCase();
    if (this.visitedWebs.has(key)) {
      return;
    }
    this.visitedWebs.add(key);
    this.stats.websChecked++;
    this.progress.currentItem = webUrl;
    this.progress.listTitle = undefined;
    this.report(true);

    const web = await this.rest.getJson<IWebInfo>(
      `${webUrl}/_api/web?$select=Title,ServerRelativeUrl,HasUniqueRoleAssignments`
    );
    // The root web of a site collection has no parent to inherit from.
    const isRoot = web.ServerRelativeUrl.toLowerCase() === this.siteServerRelativeUrl.toLowerCase();
    const wanted = isSelected ? this.options.includeSelected : this.options.includeSubsites;
    if (web.HasUniqueRoleAssignments && !isRoot && wanted) {
      this.add({
        key: `web|${key}`,
        kind: 'web',
        name: web.Title,
        path: web.ServerRelativeUrl,
        webUrl,
        depth: 0,
        status: 'found'
      });
    }

    if (isSelected && !this.options.recursive) {
      return;
    }

    // One failing list or subsite (a 403 on something with unique
    // permissions) is recorded and the rest of the web is still scanned.
    try {
      const lists = await getWebLists(this.rest, webUrl);
      for (const list of lists.filter(isContentList)) {
        this.rest.throwIfCancelled();
        try {
          await this.scanList(webUrl, list, false);
        } catch (err) {
          if (isCancelled(err)) {
            throw err;
          }
          this.errors.push(`${list.Title} (${list.RootFolder.ServerRelativeUrl}): ${message(err)}`);
        }
      }
    } catch (err) {
      if (isCancelled(err)) {
        throw err;
      }
      this.errors.push(`Lists of ${webUrl}: ${message(err)}`);
    }

    if (!this.options.includeSubsites) {
      return;
    }
    let subwebs: IWebInfo[] = [];
    try {
      const json = await this.rest.getJson<{ value: IWebInfo[] }>(
        `${webUrl}/_api/web/webs?$select=Title,ServerRelativeUrl,WebTemplate`
      );
      subwebs = (json.value || []).filter((w) => w.WebTemplate !== 'APP');
    } catch (err) {
      if (isCancelled(err)) {
        throw err;
      }
      this.errors.push(`Subsites of ${webUrl}: ${message(err)}`);
    }
    for (const sub of subwebs) {
      this.rest.throwIfCancelled();
      const subUrl = this.scopeService.toAbsoluteUrl(sub.ServerRelativeUrl);
      try {
        await this.scanWeb(subUrl, false);
      } catch (err) {
        if (isCancelled(err)) {
          throw err;
        }
        this.errors.push(`${sub.Title} (${subUrl}): ${message(err)}`);
      }
    }
  }

  private async scanList(webUrl: string, list: IListInfo, isSelected: boolean): Promise<void> {
    this.stats.listsChecked++;
    this.progress.currentItem = list.RootFolder.ServerRelativeUrl;
    this.report();

    const wanted = isSelected ? this.options.includeSelected : this.options.includeLists;
    if (list.HasUniqueRoleAssignments && wanted) {
      this.add({
        key: `list|${list.Id}`,
        kind: list.BaseType === 1 ? 'library' : 'list',
        name: list.Title,
        path: list.RootFolder.ServerRelativeUrl,
        webUrl,
        listId: list.Id,
        listTitle: list.Title,
        depth: 0,
        status: 'found'
      });
    }

    if (isSelected && !this.options.recursive) {
      return;
    }
    await this.scanItems(webUrl, list, [{ path: list.RootFolder.ServerRelativeUrl, depth: 0 }]);
  }

  private async scanFolderScopes(folders: IScopeNode[]): Promise<void> {
    const first = folders[0];
    const list = await this.getList(first.webUrl, first.listId as string);
    const rootPath = list.RootFolder.ServerRelativeUrl;
    const scopes: IItemScope[] = folders.map((f) => ({
      path: f.serverRelativeUrl,
      depth: depthBelow(f.serverRelativeUrl, rootPath),
      itemCount: f.itemCount
    }));

    if (this.options.includeSelected) {
      for (const folder of folders) {
        this.rest.throwIfCancelled();
        if (folder.itemId === undefined) {
          continue;
        }
        this.progress.currentItem = folder.serverRelativeUrl;
        this.report();
        const row = await this.rest.getJson<IItemRow>(
          `${folder.webUrl}/_api/web/lists(guid'${list.Id}')/items(${folder.itemId})?$select=${itemSelect(list)}`
        );
        this.stats.itemsChecked++;
        if (row.HasUniqueRoleAssignments) {
          this.add(this.toObject(folder.webUrl, list, row, depthBelow(row.FileRef, rootPath)));
        }
      }
    }

    if (!this.options.recursive) {
      return;
    }
    await this.scanItems(first.webUrl, list, scopes);
  }

  /**
   * Reads every item of the list once and keeps those below one of `scopes`
   * (the list root or selected folders) that have unique permissions and
   * that the options ask for. maxDepth is measured from the scope's depth.
   */
  private async scanItems(webUrl: string, list: IListInfo, scopes: IItemScope[]): Promise<void> {
    const opts = this.options;
    if (!opts.includeFolders && !opts.includeFiles) {
      return;
    }
    // A whole-list scan already covers any folder of that list.
    if (this.contentScannedLists.has(list.Id)) {
      return;
    }
    const rootPath = list.RootFolder.ServerRelativeUrl;
    const isWholeList = scopes.some((sc) => sc.path.toLowerCase() === rootPath.toLowerCase());
    if (isWholeList && !opts.maxDepth) {
      this.contentScannedLists.add(list.Id);
    }
    if (list.ItemCount === 0) {
      return;
    }

    const listApi = `${webUrl}/_api/web/lists(guid'${list.Id}')`;
    const newest = await this.rest.getJson<{ value: { Id: number }[] }>(
      `${listApi}/items?$select=Id&$orderby=Id desc&$top=1`
    );
    const maxId = newest.value && newest.value.length > 0 ? newest.value[0].Id : 0;
    if (!maxId) {
      return;
    }

    // A selected folder, or a depth limit, usually covers a small part of a
    // big library: walking just those folders is far cheaper than reading
    // every item. The walk gives up (and the ranges below take over) if it
    // would cost more than reading the whole list.
    if (list.BaseType === 1 && (!isWholeList || opts.maxDepth > 0)) {
      const rangeCost = Math.ceil(maxId / ID_RANGE);
      if (await this.walkFolders(webUrl, list, scopes, Math.max(MIN_WALK_BUDGET, rangeCost * WALK_BUDGET_FACTOR))) {
        return;
      }
    }

    const ranges: [number, number][] = [];
    for (let start = 1; start <= maxId; start += ID_RANGE) {
      ranges.push([start, Math.min(start + ID_RANGE - 1, maxId)]);
    }

    // Outermost scope first, so an item under two selected folders is measured from the higher one.
    const prefixes = scopes
      .map((sc) => ({ prefix: `${sc.path.toLowerCase()}/`, depth: sc.depth }))
      .sort((a, b) => a.depth - b.depth);
    this.progress.listTitle = list.Title;
    this.progress.listItemsTotal = maxId;
    this.progress.listItemsDone = 0;
    this.report(true);

    let consecutiveFailures = 0;
    await runPool(ranges, this.lanes, async ([from, to]) => {
      const rows = await this.fetchRange(listApi, list, from, to);
      if (rows === undefined) {
        if (++consecutiveFailures >= MAX_CONSECUTIVE_FAILURES) {
          throw new Error(
            `stopped reading after ${MAX_CONSECUTIVE_FAILURES} failed requests in a row; ` +
              'the errors above say why. Run the scan again later, or with Gentle speed.'
          );
        }
        return;
      }
      consecutiveFailures = 0;
      for (const row of rows) {
        let baseDepth = 0;
        if (!isWholeList) {
          const lower = row.FileRef.toLowerCase();
          const match = prefixes.filter((p) => lower.indexOf(p.prefix) === 0)[0];
          if (!match) {
            continue;
          }
          baseDepth = match.depth;
        }
        this.stats.itemsChecked++;
        this.consider(webUrl, list, row, baseDepth);
      }
      this.progress.listItemsDone = (this.progress.listItemsDone || 0) + (to - from + 1);
      this.progress.currentItem = `${list.Title}: items ${from.toLocaleString()}-${to.toLocaleString()}`;
      this.report();
    });
    this.progress.listTitle = undefined;
  }

  /** Keeps a row if it has unique permissions and is a kind and depth the options ask for. */
  private consider(webUrl: string, list: IListInfo, row: IItemRow, baseDepth: number): void {
    const opts = this.options;
    if (!row.HasUniqueRoleAssignments) {
      return;
    }
    const depth = depthBelow(row.FileRef, list.RootFolder.ServerRelativeUrl);
    const relative = depth - baseDepth;
    if (relative <= 0 || (opts.maxDepth > 0 && relative > opts.maxDepth)) {
      return;
    }
    const isFolder = Number(row.FSObjType) === 1;
    if (isFolder ? !opts.includeFolders : !opts.includeFiles) {
      return;
    }
    this.add(this.toObject(webUrl, list, row, depth));
  }

  /**
   * Lists the selected folders' subfolders (and files, if wanted) level by
   * level, only as deep as maxDepth needs. Returns false if it handed over to
   * the ID-range read: over budget, a folder too big to list, or a tenant
   * that won't return HasUniqueRoleAssignments through Folder/ListItemAllFields.
   */
  private async walkFolders(webUrl: string, list: IListInfo, scopes: IItemScope[], budget: number): Promise<boolean> {
    const opts = this.options;
    const folderApi = (path: string): string =>
      `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${quoteForUrl(path)}')`;
    let requests = 0;
    const checkedBefore = this.stats.itemsChecked;
    const get = async <T>(url: string): Promise<T> => {
      if (++requests > budget) {
        throw new WalkAbandoned('the folders hold more than reading the whole list would cost');
      }
      return this.rest.getJson<T>(url);
    };

    this.progress.listTitle = undefined;
    try {
      // Each entry: a folder to list, its depth below its scope, and the scope's depth.
      let level: { path: string; relative: number; base: number; itemCount?: number }[] = scopes.map((sc) => ({
        path: sc.path,
        relative: 0,
        base: sc.depth,
        itemCount: sc.itemCount
      }));
      while (level.length > 0) {
        const next: typeof level = [];
        await runPool(level, this.lanes, async (folder) => {
          // Children of this folder sit at relative + 1.
          if (opts.maxDepth > 0 && folder.relative + 1 > opts.maxDepth) {
            return;
          }
          let itemCount = folder.itemCount;
          if (itemCount === undefined) {
            itemCount = (await get<{ ItemCount?: number }>(`${folderApi(folder.path)}?$select=ItemCount`)).ItemCount || 0;
          }
          if (itemCount > MAX_FOLDER_ITEMS_FOR_WALK) {
            throw new WalkAbandoned(`${folder.path} holds ${itemCount.toLocaleString()} items directly`);
          }
          if (itemCount === 0) {
            return;
          }
          this.progress.currentItem = folder.path;
          this.report();

          const subfolders = await get<{ value: IFolderEntry[] }>(
            `${folderApi(folder.path)}/Folders?$select=Name,ServerRelativeUrl,ItemCount,${FOLDER_ITEM_SELECT}&$expand=ListItemAllFields`
          );
          for (const sub of subfolders.value || []) {
            const item = sub.ListItemAllFields;
            if (!item || typeof item.Id !== 'number') {
              continue; // "Forms" and other system folders have no list item
            }
            this.stats.itemsChecked++;
            const row: IItemRow = { Id: item.Id, FileRef: sub.ServerRelativeUrl, FSObjType: 1, HasUniqueRoleAssignments: !!item.HasUniqueRoleAssignments };
            this.consider(webUrl, list, row, folder.base);
            next.push({ path: sub.ServerRelativeUrl, relative: folder.relative + 1, base: folder.base, itemCount: sub.ItemCount });
          }

          if (opts.includeFiles) {
            const files = await get<{ value: IFolderEntry[] }>(
              `${folderApi(folder.path)}/Files?$select=Name,ServerRelativeUrl,${FOLDER_ITEM_SELECT}&$expand=ListItemAllFields`
            );
            for (const file of files.value || []) {
              const item = file.ListItemAllFields;
              if (!item || typeof item.Id !== 'number') {
                continue;
              }
              this.stats.itemsChecked++;
              const row: IItemRow = { Id: item.Id, FileRef: file.ServerRelativeUrl, FSObjType: 0, HasUniqueRoleAssignments: !!item.HasUniqueRoleAssignments };
              this.consider(webUrl, list, row, folder.base);
            }
          }
          this.report();
        });
        level = next;
      }
      return true;
    } catch (err) {
      if (isCancelled(err)) {
        throw err;
      }
      if (!(err instanceof WalkAbandoned) && !isBadRequest(err) && !isWorthSplitting(err)) {
        throw err;
      }
      // The ID-range read covers the same items again; don't count them twice.
      // Anything the walk already found is kept (objects are keyed by item).
      this.stats.itemsChecked = checkedBefore;
      console.info(`[ReInherit] Folder walk of ${list.Title} handed over to an ID-range read: ${err instanceof Error ? err.message : err}`);
      return false;
    }
  }

  /** The rows with IDs from..to, or undefined if they could not be read (the error is recorded). */
  private async fetchRange(
    listApi: string,
    list: IListInfo,
    from: number,
    to: number,
    splits = 0
  ): Promise<IItemRow[] | undefined> {
    try {
      const json = await this.rest.getJson<{ value: IItemRow[] }>(
        `${listApi}/items?$select=${itemSelect(list)}&$filter=Id ge ${from} and Id le ${to}&$top=${to - from + 1}`
      );
      return json.value || [];
    } catch (err) {
      if (isCancelled(err)) {
        throw err;
      }
      if (splits < MAX_SPLITS && to > from && isWorthSplitting(err)) {
        const mid = Math.floor((from + to) / 2);
        const first = await this.fetchRange(listApi, list, from, mid, splits + 1);
        const second = await this.fetchRange(listApi, list, mid + 1, to, splits + 1);
        return first && second ? first.concat(second) : first || second;
      }
      this.errors.push(`${list.Title}: items with ID ${from}-${to} could not be read (${message(err)})`);
      return undefined;
    }
  }

  private toObject(webUrl: string, list: IListInfo, row: IItemRow, depth: number): IUniqueObject {
    const isFolder = Number(row.FSObjType) === 1;
    const isLibrary = list.BaseType === 1;
    return {
      key: `item|${list.Id}|${row.Id}`,
      kind: isFolder ? 'folder' : isLibrary ? 'file' : 'item',
      name: isLibrary || !row.Title ? row.FileRef.substring(row.FileRef.lastIndexOf('/') + 1) : row.Title,
      path: row.FileRef,
      webUrl,
      listId: list.Id,
      listTitle: list.Title,
      itemId: row.Id,
      depth,
      status: 'found'
    };
  }
}
