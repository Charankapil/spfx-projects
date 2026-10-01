import { IRestoreOptions } from '../models/IRestoreOptions';
import { IRunProgress } from '../models/IRunProgress';
import { EMPTY_STATS, IRunStats } from '../models/IRunReport';
import { IScopeNode } from '../models/IScopeNode';
import { IUniqueObject } from '../models/IUniqueObject';
import { runPool } from './pool';
import { getWebLists, IListInfo, isContentList, LIST_SELECT, ScopeService } from './ScopeService';
import { CancelledError, HttpError, SpRest } from './SpRest';

/**
 * Items are read in ID ranges rather than by following the paging link: ID
 * is always indexed, so "Id ge a and Id le b" stays under the 5,000-item
 * list view threshold in a library of any size, and independent ranges can
 * be fetched in parallel. A range that times out is split in half and
 * retried, down to MIN_ID_RANGE.
 */
const ID_RANGE = 5000;
const MIN_ID_RANGE = 250;
const PARALLEL_RANGES = 2;

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

function isRetryable(err: unknown): boolean {
  // Timeouts and threshold errors surface as 5xx; a TypeError is a dropped connection.
  return !(err instanceof HttpError) || err.status >= 500 || err.status === 408;
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
    private onProgress: (progress: IRunProgress) => void
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
      if (err instanceof CancelledError) {
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
          if (err instanceof CancelledError) {
            throw err;
          }
          this.errors.push(`${list.Title} (${list.RootFolder.ServerRelativeUrl}): ${message(err)}`);
        }
      }
    } catch (err) {
      if (err instanceof CancelledError) {
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
      if (err instanceof CancelledError) {
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
        if (err instanceof CancelledError) {
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
      depth: depthBelow(f.serverRelativeUrl, rootPath)
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

    await runPool(ranges, PARALLEL_RANGES, async ([from, to]) => {
      const rows = await this.fetchRange(listApi, list, from, to);
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
        if (!row.HasUniqueRoleAssignments) {
          continue;
        }
        const depth = depthBelow(row.FileRef, rootPath);
        const relative = depth - baseDepth;
        if (relative <= 0 || (opts.maxDepth > 0 && relative > opts.maxDepth)) {
          continue;
        }
        const isFolder = Number(row.FSObjType) === 1;
        if (isFolder ? !opts.includeFolders : !opts.includeFiles) {
          continue;
        }
        this.add(this.toObject(webUrl, list, row, depth));
      }
      this.progress.listItemsDone = (this.progress.listItemsDone || 0) + (to - from + 1);
      this.progress.currentItem = `${list.Title}: items ${from.toLocaleString()}-${to.toLocaleString()}`;
      this.report();
    });
    this.progress.listTitle = undefined;
  }

  private async fetchRange(listApi: string, list: IListInfo, from: number, to: number): Promise<IItemRow[]> {
    try {
      const json = await this.rest.getJson<{ value: IItemRow[] }>(
        `${listApi}/items?$select=${itemSelect(list)}&$filter=Id ge ${from} and Id le ${to}&$top=${to - from + 1}`
      );
      return json.value || [];
    } catch (err) {
      if (err instanceof CancelledError) {
        throw err;
      }
      if (to - from + 1 > MIN_ID_RANGE && isRetryable(err)) {
        const mid = Math.floor((from + to) / 2);
        const first = await this.fetchRange(listApi, list, from, mid);
        const second = await this.fetchRange(listApi, list, mid + 1, to);
        return first.concat(second);
      }
      this.errors.push(`${list.Title}: items with ID ${from}-${to} could not be read (${message(err)})`);
      return [];
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
