import { entityOf, nextLinkOf, odataString, rowsOf, SPClient, trimSlash } from '../core/SPClient';
import {
  IChangeRow,
  IGroupInfo,
  IListInfo,
  IRecycleItem,
  IRoleAssignment,
  ISiteInfo,
  ISubWeb,
  IUserInfo,
  IWebInfo,
  ListKind,
  UserKind
} from '../models';
import { toDate } from './SearchApi';

/**
 * SharePoint REST calls for administration. Every method is a small, bounded
 * request against one web: metadata only - never item or file enumeration.
 * Paged reads have a hard page cap, so a huge site cannot turn a click into
 * hundreds of requests.
 */

const USER_PAGE = 1000;
const USER_MAX_PAGES = 5;

function num(v: unknown): number {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function str(v: unknown): string {
  return v === null || v === undefined ? '' : String(v);
}

export function classifyUser(u: { LoginName?: string; PrincipalType?: number; IsShareByEmailGuestUser?: boolean; IsEmailAuthenticationGuestUser?: boolean }): UserKind {
  const login = (u.LoginName || '').toLowerCase();
  if (login.indexOf('spo-grid-all-users') >= 0 || login === 'c:0(.s|true') {
    return 'OrgWide';
  }
  if (login.indexOf('#ext#') >= 0 || login.indexOf('urn:spo:guest') >= 0 || u.IsShareByEmailGuestUser || u.IsEmailAuthenticationGuestUser) {
    return 'Guest';
  }
  if (login.indexOf('sharepoint\\system') >= 0 || login === 'c:0!.s|windows' || login.indexOf('spocrawl') >= 0 || login.indexOf('app@sharepoint') >= 0) {
    return 'System';
  }
  if (u.PrincipalType === 4 || u.PrincipalType === 8 || u.PrincipalType === 2) {
    return 'Group';
  }
  return 'Member';
}

/** Turns an e-mail address into the claims login SharePoint expects. */
export function toLogin(input: string): string {
  const v = input.trim();
  if (v.indexOf('|') >= 0) {
    return v;
  }
  return 'i:0#.f|membership|' + v;
}

/**
 * SharePoint change token for "everything since `ms`": 1;<scope 1 = site>;<site id>;<.NET ticks>;-1.
 * Ticks = (ms since 1970 + 62135596800000) * 10000. The final four digits are appended as text
 * because the full value (~6.4e17) exceeds JavaScript's safe integer range.
 */
export function changeToken(siteId: string, ms: number): string {
  return `1;1;${siteId};${String(Math.floor(ms) + 62135596800000)}0000;-1`;
}

export function permissionLevel(low: number, high: number): string {
  const has = (kind: number): boolean => {
    if (kind <= 32) {
      return Math.floor(low / Math.pow(2, kind - 1)) % 2 === 1;
    }
    return Math.floor(high / Math.pow(2, kind - 33)) % 2 === 1;
  };
  // PermissionKind numbers: ViewListItems 1, AddListItems 2, EditListItems 3,
  // DeleteListItems 4, ManageWeb 19, ManagePermissions 25, FullMask = all bits.
  if (low >= 4294967295 && high >= 2147483647) {
    return 'Full control';
  }
  if (has(19) && has(25)) {
    return 'Design / manage';
  }
  if (has(2) && has(3) && has(4)) {
    return 'Edit / contribute';
  }
  if (has(1)) {
    return 'Read';
  }
  return 'No access';
}

export class AdminApi {
  constructor(private client: SPClient) {}

  private api(webUrl: string, path: string): string {
    return `${trimSlash(webUrl)}/_api/${path}`;
  }

  // ---- site / web ------------------------------------------------------------

  public async getWeb(webUrl: string): Promise<IWebInfo> {
    const j = await this.client.get<{ [k: string]: unknown }>(
      this.api(
        webUrl,
        'web?$select=Id,Title,Description,Url,ServerRelativeUrl,Created,LastItemModifiedDate,WebTemplate,Language,HasUniqueRoleAssignments,SiteLogoUrl'
      )
    );
    const w = entityOf<{ [k: string]: unknown }>(j);
    return {
      id: str(w.Id),
      title: str(w.Title),
      description: str(w.Description),
      url: trimSlash(str(w.Url)),
      serverRelativeUrl: str(w.ServerRelativeUrl),
      created: toDate(str(w.Created)),
      lastModified: toDate(str(w.LastItemModifiedDate)),
      template: str(w.WebTemplate),
      language: num(w.Language),
      hasUniquePermissions: !!w.HasUniqueRoleAssignments,
      siteLogoUrl: str(w.SiteLogoUrl) || undefined
    };
  }

  public async getSite(webUrl: string): Promise<ISiteInfo> {
    const j = await this.client.get<{ [k: string]: unknown }>(
      this.api(webUrl, 'site?$select=Id,Url,ReadOnly,Usage,GroupId,HubSiteId,Classification')
    );
    const s = entityOf<{ [k: string]: unknown }>(j);
    const usage = (s.Usage || {}) as { [k: string]: unknown };
    const info: ISiteInfo = {
      id: str(s.Id),
      url: trimSlash(str(s.Url)),
      readOnly: !!s.ReadOnly,
      storageBytes: num(usage.Storage),
      storageFraction: num(usage.StoragePercentageUsed),
      groupId: str(s.GroupId) || undefined,
      hubSiteId: str(s.HubSiteId) || undefined,
      classification: str(s.Classification) || undefined
    };
    // Primary owner is nice to have; never let it break the page.
    try {
      const o = entityOf<{ Title?: string; Email?: string }>(await this.client.get(this.api(webUrl, 'site/owner?$select=Title,Email')));
      info.ownerTitle = o.Title;
      info.ownerEmail = o.Email;
    } catch {
      /* optional */
    }
    return info;
  }

  /** Storage of one site collection: a single small request (never cached; used for snapshots). */
  public async getSiteUsage(siteUrl: string): Promise<{ bytes: number; fraction: number }> {
    const j = await this.client.get<{ [k: string]: unknown }>(this.api(siteUrl, 'site?$select=Usage'), { cacheTtlMs: 0 });
    const usage = (entityOf<{ [k: string]: unknown }>(j).Usage || {}) as { [k: string]: unknown };
    return { bytes: num(usage.Storage), fraction: num(usage.StoragePercentageUsed) };
  }

  public async getSubWebs(webUrl: string): Promise<ISubWeb[]> {
    const j = await this.client.get(
      this.api(webUrl, 'web/webs?$select=Id,Title,Url,WebTemplate,Created,LastItemModifiedDate,HasUniqueRoleAssignments&$top=500')
    );
    return rowsOf<{ [k: string]: unknown }>(j).map((w) => ({
      id: str(w.Id),
      title: str(w.Title),
      url: trimSlash(str(w.Url)),
      template: str(w.WebTemplate),
      created: toDate(str(w.Created)),
      lastModified: toDate(str(w.LastItemModifiedDate)),
      hasUniquePermissions: !!w.HasUniqueRoleAssignments
    }));
  }

  public async updateWeb(webUrl: string, changes: { Title?: string; Description?: string }): Promise<void> {
    await this.client.post(this.api(webUrl, 'web'), changes, { method: 'MERGE' }, webUrl);
  }

  // ---- lists & libraries -----------------------------------------------------

  public async getLists(webUrl: string): Promise<IListInfo[]> {
    const j = await this.client.get(
      this.api(
        webUrl,
        'web/lists?$select=Id,Title,Description,BaseTemplate,BaseType,ItemCount,Created,LastItemModifiedDate,LastItemUserModifiedDate,EnableVersioning,EnableMinorVersions,EnableModeration,Hidden,NoCrawl,HasUniqueRoleAssignments,IsSystemList,DefaultViewUrl,RootFolder/ServerRelativeUrl&$expand=RootFolder&$top=1000'
      )
    );
    const origin = (/^(https:\/\/[^/]+)/i.exec(webUrl) || [''])[1];
    return rowsOf<{ [k: string]: unknown }>(j).map((l) => {
      const baseType = num(l.BaseType);
      const kind: ListKind = baseType === 1 ? 'Library' : baseType === 0 ? 'List' : 'Other';
      const root = (l.RootFolder || {}) as { ServerRelativeUrl?: string };
      return {
        id: str(l.Id),
        title: str(l.Title),
        description: str(l.Description),
        kind,
        baseTemplate: num(l.BaseTemplate),
        itemCount: num(l.ItemCount),
        created: toDate(str(l.Created)),
        lastModified: toDate(str(l.LastItemUserModifiedDate) || str(l.LastItemModifiedDate)),
        versioning: !!l.EnableVersioning,
        minorVersions: !!l.EnableMinorVersions,
        moderation: !!l.EnableModeration,
        hidden: !!l.Hidden,
        noCrawl: !!l.NoCrawl,
        uniquePermissions: !!l.HasUniqueRoleAssignments,
        isSystem: !!l.IsSystemList,
        url: origin + str(root.ServerRelativeUrl),
        rootFolder: str(root.ServerRelativeUrl)
      };
    });
  }

  public async updateList(webUrl: string, listId: string, changes: { EnableVersioning?: boolean; NoCrawl?: boolean; Hidden?: boolean }): Promise<void> {
    await this.client.post(this.api(webUrl, `web/lists/getbyid('${odataString(listId)}')`), changes, { method: 'MERGE' }, webUrl);
  }

  // ---- people ---------------------------------------------------------------

  public async getUsers(webUrl: string): Promise<{ users: IUserInfo[]; truncated: boolean }> {
    let url: string | undefined = this.api(
      webUrl,
      `web/siteusers?$select=Id,Title,Email,LoginName,PrincipalType,IsSiteAdmin,IsShareByEmailGuestUser,IsEmailAuthenticationGuestUser&$top=${USER_PAGE}`
    );
    const users: IUserInfo[] = [];
    let pages = 0;
    let truncated = false;
    while (url) {
      const j: unknown = await this.client.get(url);
      rowsOf<{ [k: string]: unknown }>(j).forEach((u) => {
        const info = {
          LoginName: str(u.LoginName),
          PrincipalType: num(u.PrincipalType),
          IsShareByEmailGuestUser: !!u.IsShareByEmailGuestUser,
          IsEmailAuthenticationGuestUser: !!u.IsEmailAuthenticationGuestUser
        };
        users.push({
          id: num(u.Id),
          title: str(u.Title),
          email: str(u.Email),
          loginName: info.LoginName,
          principalType: info.PrincipalType,
          isSiteAdmin: !!u.IsSiteAdmin,
          kind: classifyUser(info)
        });
      });
      pages++;
      url = nextLinkOf(j);
      if (url && pages >= USER_MAX_PAGES) {
        truncated = true;
        break;
      }
    }
    return { users, truncated };
  }

  public async getGroups(webUrl: string): Promise<IGroupInfo[]> {
    const j = await this.client.get(
      this.api(webUrl, 'web/sitegroups?$select=Id,Title,Description,OwnerTitle,AllowMembersEditMembership&$top=500')
    );
    const groups: IGroupInfo[] = rowsOf<{ [k: string]: unknown }>(j).map((g) => ({
      id: num(g.Id),
      title: str(g.Title),
      description: str(g.Description),
      ownerTitle: str(g.OwnerTitle),
      allowMembersEditMembership: !!g.AllowMembersEditMembership
    }));
    // Which groups are the Owners / Members / Visitors of the site (optional).
    try {
      const a = entityOf<{ [k: string]: { Id?: number } | undefined }>(
        await this.client.get(
          this.api(
            webUrl,
            'web?$select=AssociatedOwnerGroup/Id,AssociatedMemberGroup/Id,AssociatedVisitorGroup/Id&$expand=AssociatedOwnerGroup,AssociatedMemberGroup,AssociatedVisitorGroup'
          )
        )
      );
      const mark = (key: string, role: 'Owners' | 'Members' | 'Visitors'): void => {
        const id = a[key] && (a[key] as { Id?: number }).Id;
        groups.forEach((g) => {
          if (id && g.id === id) {
            g.role = role;
          }
        });
      };
      mark('AssociatedOwnerGroup', 'Owners');
      mark('AssociatedMemberGroup', 'Members');
      mark('AssociatedVisitorGroup', 'Visitors');
    } catch {
      /* optional */
    }
    return groups;
  }

  public async getGroupMembers(webUrl: string, groupId: number): Promise<IUserInfo[]> {
    const j = await this.client.get(
      this.api(
        webUrl,
        `web/sitegroups/getbyid(${groupId})/users?$select=Id,Title,Email,LoginName,PrincipalType,IsSiteAdmin,IsShareByEmailGuestUser,IsEmailAuthenticationGuestUser&$top=1000`
      ),
      { cacheTtlMs: 60000 }
    );
    return rowsOf<{ [k: string]: unknown }>(j).map((u) => ({
      id: num(u.Id),
      title: str(u.Title),
      email: str(u.Email),
      loginName: str(u.LoginName),
      principalType: num(u.PrincipalType),
      isSiteAdmin: !!u.IsSiteAdmin,
      kind: classifyUser({
        LoginName: str(u.LoginName),
        PrincipalType: num(u.PrincipalType),
        IsShareByEmailGuestUser: !!u.IsShareByEmailGuestUser,
        IsEmailAuthenticationGuestUser: !!u.IsEmailAuthenticationGuestUser
      })
    }));
  }

  /** Resolves (and if needed creates the site-level entry for) a person. */
  public async ensureUser(webUrl: string, emailOrLogin: string): Promise<{ id: number; loginName: string; title: string }> {
    const r = await this.client.post<{ Id?: number; LoginName?: string; Title?: string }>(
      this.api(webUrl, 'web/ensureuser'),
      { logonName: toLogin(emailOrLogin) },
      {},
      webUrl
    );
    const u = entityOf<{ Id?: number; LoginName?: string; Title?: string }>(r);
    if (!u || !u.LoginName) {
      throw new Error(`Could not find "${emailOrLogin}".`);
    }
    return { id: num(u.Id), loginName: u.LoginName, title: str(u.Title) };
  }

  public async addUserToGroup(webUrl: string, groupId: number, loginName: string): Promise<void> {
    await this.client.post(this.api(webUrl, `web/sitegroups/getbyid(${groupId})/users`), { LoginName: loginName }, {}, webUrl);
  }

  public async removeUserFromGroup(webUrl: string, groupId: number, userId: number): Promise<void> {
    await this.client.post(this.api(webUrl, `web/sitegroups/getbyid(${groupId})/users/removebyid(${userId})`), undefined, {}, webUrl);
  }

  public async removeUserFromSite(webUrl: string, userId: number): Promise<void> {
    await this.client.post(this.api(webUrl, `web/siteusers/removebyid(${userId})`), undefined, {}, webUrl);
  }

  public async setSiteAdmin(webUrl: string, userId: number, isAdmin: boolean): Promise<void> {
    await this.client.post(this.api(webUrl, `web/siteusers/getbyid(${userId})`), { IsSiteAdmin: isAdmin }, { method: 'MERGE' }, webUrl);
  }

  public async getRoleAssignments(webUrl: string): Promise<IRoleAssignment[]> {
    const j = await this.client.get(
      this.api(
        webUrl,
        'web/roleassignments?$select=PrincipalId,Member/Title,Member/LoginName,Member/PrincipalType,RoleDefinitionBindings/Name&$expand=Member,RoleDefinitionBindings&$top=500'
      )
    );
    return rowsOf<{ [k: string]: unknown }>(j).map((r) => {
      const m = (r.Member || {}) as { [k: string]: unknown };
      const roles = rowsOf<{ Name?: string }>(r.RoleDefinitionBindings).map((x) => str(x.Name));
      return {
        principalId: num(r.PrincipalId),
        principalTitle: str(m.Title),
        principalLogin: str(m.LoginName),
        principalType: num(m.PrincipalType),
        roles
      };
    });
  }

  public async getUserGroups(webUrl: string, userId: number): Promise<Array<{ id: number; title: string }>> {
    const j = await this.client.get(this.api(webUrl, `web/siteusers/getbyid(${userId})/groups?$select=Id,Title`), { cacheTtlMs: 60000 });
    return rowsOf<{ Id?: number; Title?: string }>(j).map((g) => ({ id: num(g.Id), title: str(g.Title) }));
  }

  public async getEffectiveLevel(webUrl: string, loginName: string): Promise<string> {
    const j = await this.client.get<{ [k: string]: unknown }>(
      this.api(webUrl, `web/getusereffectivepermissions(@v)?@v='${encodeURIComponent(odataString(loginName))}'`),
      { cacheTtlMs: 60000 }
    );
    const e = entityOf<{ [k: string]: unknown }>(j);
    const bp = ((e.GetUserEffectivePermissions || e.value || e) as { [k: string]: unknown });
    return permissionLevel(num(bp.Low), num(bp.High));
  }

  // ---- recycle bin ----------------------------------------------------------

  public async getRecycleBin(webUrl: string, top: number): Promise<IRecycleItem[]> {
    const select = '$select=Id,Title,LeafName,DirName,ItemType,ItemState,Size,DeletedDate,DeletedByName,AuthorName';
    const base = this.api(webUrl, `site/recyclebin?${select}&$top=${top}`);
    let j: unknown;
    try {
      j = await this.client.get(base + '&$orderby=DeletedDate desc'.replace(/ /g, '%20'));
    } catch {
      // Not every tenant accepts $orderby here: fall back to the default order.
      j = await this.client.get(base);
    }
    return rowsOf<{ [k: string]: unknown }>(j)
      .map((r) => ({
        id: str(r.Id),
        title: str(r.Title),
        leafName: str(r.LeafName),
        dirName: str(r.DirName),
        itemType: num(r.ItemType),
        stage: (num(r.ItemState) === 2 ? 2 : 1) as 1 | 2,
        size: num(r.Size),
        deletedDate: toDate(str(r.DeletedDate)),
        deletedBy: str(r.DeletedByName),
        author: str(r.AuthorName)
      }))
      .sort((a, b) => (b.deletedDate ? b.deletedDate.getTime() : 0) - (a.deletedDate ? a.deletedDate.getTime() : 0));
  }

  public async restoreRecycleItem(webUrl: string, id: string): Promise<void> {
    await this.client.post(this.api(webUrl, `site/recyclebin('${odataString(id)}')/restore`), undefined, {}, webUrl);
  }

  public async deleteRecycleItem(webUrl: string, id: string): Promise<void> {
    await this.client.post(this.api(webUrl, `site/recyclebin('${odataString(id)}')/deleteObject`), undefined, {}, webUrl);
  }

  // ---- change log -----------------------------------------------------------

  /** Site-collection change log for the last `days` days (retained ~60 days). */
  public async getChanges(webUrl: string, siteId: string, days: number): Promise<IChangeRow[]> {
    const token = changeToken(siteId, Date.now() - days * 86400000);
    const body = {
      query: {
        __metadata: { type: 'SP.ChangeQuery' },
        Add: true,
        Update: true,
        DeleteObject: true,
        Restore: true,
        Rename: true,
        GroupMembershipAdd: true,
        GroupMembershipDelete: true,
        RoleAssignmentAdd: true,
        RoleAssignmentDelete: true,
        RoleDefinitionAdd: true,
        RoleDefinitionDelete: true,
        RoleDefinitionUpdate: true,
        SecurityPolicy: true,
        Site: true,
        Web: true,
        List: true,
        Group: true,
        User: true,
        Feature: true,
        Item: false,
        File: false,
        Folder: false,
        Field: false,
        View: false,
        ContentType: false,
        Alert: false,
        Navigation: false,
        SystemUpdate: false,
        FetchLimit: 500,
        ChangeTokenStart: { __metadata: { type: 'SP.ChangeToken' }, StringValue: token }
      }
    };
    const j = await this.client.post<unknown>(this.api(webUrl, 'site/getchanges'), body, { odata3: true }, webUrl);
    return rowsOf<{ [k: string]: unknown }>(j).map(describeChange).reverse();
  }
}

const CHANGE_TYPES: { [n: number]: string } = {
  1: 'Added',
  2: 'Updated',
  3: 'Deleted',
  4: 'Renamed',
  5: 'Moved away',
  6: 'Moved in',
  7: 'Restored',
  8: 'Role definition added',
  9: 'Role definition deleted',
  10: 'Role definition updated',
  11: 'Permission assigned',
  12: 'Permission removed',
  13: 'Member added',
  14: 'Member removed',
  15: 'System update',
  16: 'Navigation changed',
  17: 'Unique permissions added',
  18: 'Unique permissions removed'
};

export function describeChange(c: { [k: string]: unknown }): IChangeRow {
  const typeName = ((c.__metadata as { type?: string } | undefined) || {}).type || '';
  const t = num(c.ChangeType);
  const objectKind = typeName.replace(/^SP\.Change/, '') || 'Object';
  let objectId = '';
  if (objectKind === 'User') {
    objectId = str(c.UserId);
  } else if (objectKind === 'Group') {
    objectId = str(c.GroupId);
  } else if (objectKind === 'List') {
    objectId = str(c.ListId);
  } else if (objectKind === 'Web') {
    objectId = str(c.WebId);
  }
  const security = t >= 8 && t <= 14 ? true : t === 17 || t === 18;
  const structure = t === 1 || t === 3 || t === 4 || t === 7;
  return {
    time: toDate(str(c.Time)),
    changeType: CHANGE_TYPES[t] || `Change ${t}`,
    objectKind,
    objectId,
    severity: security ? 'security' : structure ? 'structure' : 'info'
  };
}
