import { WebPartContext } from '@microsoft/sp-webpart-base';

import { IScopeNode } from '../models/IScopeNode';
import { HttpError, quoteForUrl, SpRest } from './SpRest';

export function isBadRequest(err: unknown): boolean {
  return err instanceof HttpError && err.status === 400;
}

interface IWebInfo {
  Title: string;
  ServerRelativeUrl: string;
  HasUniqueRoleAssignments: boolean;
  WebTemplate?: string;
}

export interface IListInfo {
  Id: string;
  Title: string;
  ItemCount: number;
  BaseType: number;
  BaseTemplate: number;
  Hidden: boolean;
  IsCatalog?: boolean;
  IsSystemList?: boolean;
  HasUniqueRoleAssignments: boolean;
  RootFolder: { ServerRelativeUrl: string };
}

interface IFolderInfo {
  Name: string;
  ServerRelativeUrl: string;
  ItemCount: number;
  ListItemAllFields?: { Id?: number; HasUniqueRoleAssignments?: boolean };
}

/** Folders shown per level in the picker; the scan itself has no such limit. */
export const MAX_FOLDERS_SHOWN = 1000;

export const LIST_SELECT =
  'Id,Title,ItemCount,BaseType,BaseTemplate,Hidden,IsCatalog,IsSystemList,HasUniqueRoleAssignments,RootFolder/ServerRelativeUrl';

/** Catalogs and system lists by template, for when IsCatalog / IsSystemList can't be selected. */
const SYSTEM_TEMPLATES = [113, 114, 116, 121, 122, 123, 124, 125, 140, 160, 544];

/**
 * The web's lists. IsSystemList is SharePoint Online only; if the tenant
 * refuses to select it, the lists are read again without it and system
 * lists are recognised by template instead.
 */
export async function getWebLists(rest: SpRest, webUrl: string): Promise<IListInfo[]> {
  try {
    const json = await rest.getJson<{ value: IListInfo[] }>(
      `${webUrl}/_api/web/lists?$select=${LIST_SELECT}&$expand=RootFolder`
    );
    return json.value || [];
  } catch (err) {
    if (!isBadRequest(err)) {
      throw err;
    }
    const json = await rest.getJson<{ value: IListInfo[] }>(
      `${webUrl}/_api/web/lists?$select=Id,Title,ItemCount,BaseType,BaseTemplate,Hidden,HasUniqueRoleAssignments,RootFolder/ServerRelativeUrl&$expand=RootFolder`
    );
    return (json.value || []).map((l) => ({ ...l, IsSystemList: SYSTEM_TEMPLATES.indexOf(l.BaseTemplate) >= 0 }));
  }
}

/**
 * Lists a person would recognise as their content: not hidden, not a
 * gallery/catalog (master pages, web parts, ...) and not one of the
 * system lists SharePoint maintains itself.
 */
export function isContentList(list: IListInfo): boolean {
  return !list.Hidden && !list.IsCatalog && !list.IsSystemList && !isProtectedLibrary(list);
}

/**
 * Libraries SharePoint's own features depend on. They commonly carry unique
 * permissions on purpose (anonymous read of the site's branding files, for
 * one), so they are never offered or reset.
 */
const PROTECTED_LIBRARIES = /\/(style library|formservertemplates|_catalogs\/[^/]+)$/i;

function isProtectedLibrary(list: IListInfo): boolean {
  return PROTECTED_LIBRARIES.test(list.RootFolder.ServerRelativeUrl);
}

export class ScopeService {
  constructor(private context: WebPartContext, private rest: SpRest) {}

  public toAbsoluteUrl(serverRelativeUrl: string): string {
    return `${new URL(this.context.pageContext.site.absoluteUrl).origin}${serverRelativeUrl}`;
  }

  /** The site collection's root web - the top of the picker. */
  public async getRootNode(): Promise<IScopeNode> {
    const siteUrl = this.context.pageContext.site.absoluteUrl;
    const web = await this.rest.getJson<IWebInfo>(
      `${siteUrl}/_api/web?$select=Title,ServerRelativeUrl,HasUniqueRoleAssignments`
    );
    return {
      key: `web|${web.ServerRelativeUrl}`,
      kind: 'web',
      title: web.Title,
      webUrl: siteUrl,
      serverRelativeUrl: web.ServerRelativeUrl,
      hasUniquePermissions: web.HasUniqueRoleAssignments,
      isRootWeb: true
    };
  }

  public async getChildren(node: IScopeNode): Promise<IScopeNode[]> {
    if (node.kind === 'web') {
      return this.getWebChildren(node);
    }
    return this.getFolderChildren(node);
  }

  private async getWebChildren(node: IScopeNode): Promise<IScopeNode[]> {
    const [webs, lists] = await Promise.all([
      this.rest.getJson<{ value: IWebInfo[] }>(
        `${node.webUrl}/_api/web/webs?$select=Title,ServerRelativeUrl,HasUniqueRoleAssignments,WebTemplate`
      ),
      getWebLists(this.rest, node.webUrl)
    ]);

    const webNodes: IScopeNode[] = (webs.value || [])
      .filter((w) => w.WebTemplate !== 'APP')
      .map((w) => ({
        key: `web|${w.ServerRelativeUrl}`,
        kind: 'web',
        title: w.Title,
        webUrl: this.toAbsoluteUrl(w.ServerRelativeUrl),
        serverRelativeUrl: w.ServerRelativeUrl,
        hasUniquePermissions: w.HasUniqueRoleAssignments
      }));

    const listNodes: IScopeNode[] = lists.filter(isContentList).map((l) => ({
      key: `list|${l.Id}`,
      kind: l.BaseType === 1 ? 'library' : 'list',
      title: l.Title,
      webUrl: node.webUrl,
      serverRelativeUrl: l.RootFolder.ServerRelativeUrl,
      listId: l.Id,
      listRootUrl: l.RootFolder.ServerRelativeUrl,
      itemCount: l.ItemCount,
      hasUniquePermissions: l.HasUniqueRoleAssignments
    }));

    const byTitle = (a: IScopeNode, b: IScopeNode): number => a.title.localeCompare(b.title);
    return [...webNodes.sort(byTitle), ...listNodes.sort(byTitle)];
  }

  /**
   * Direct subfolders of a library or folder. Folder.Folders is a folder
   * listing, not a list query, so it is not subject to the 5,000-item list
   * view threshold even in very large libraries.
   */
  private async getFolderChildren(node: IScopeNode): Promise<IScopeNode[]> {
    const base =
      `${node.webUrl}/_api/web/GetFolderByServerRelativeUrl('${quoteForUrl(node.serverRelativeUrl)}')/Folders` +
      `?$expand=ListItemAllFields&$top=${MAX_FOLDERS_SHOWN}&$select=Name,ServerRelativeUrl,ItemCount,ListItemAllFields/Id`;
    let json: { value: IFolderInfo[] };
    try {
      json = await this.rest.getJson<{ value: IFolderInfo[] }>(`${base},ListItemAllFields/HasUniqueRoleAssignments`);
    } catch (err) {
      // Without the flag the folders still load; they just aren't tagged "Unique".
      if (!isBadRequest(err)) {
        throw err;
      }
      json = await this.rest.getJson<{ value: IFolderInfo[] }>(base);
    }
    return (json.value || [])
      .filter((f) => f.ListItemAllFields && typeof f.ListItemAllFields.Id === 'number')
      .map(
        (f): IScopeNode => ({
          key: `folder|${f.ServerRelativeUrl}`,
          kind: 'folder',
          title: f.Name,
          webUrl: node.webUrl,
          serverRelativeUrl: f.ServerRelativeUrl,
          listId: node.listId,
          listRootUrl: node.listRootUrl,
          itemId: f.ListItemAllFields?.Id,
          itemCount: f.ItemCount,
          hasUniquePermissions: !!f.ListItemAllFields?.HasUniqueRoleAssignments
        })
      )
      .sort((a, b) => a.title.localeCompare(b.title));
  }
}
