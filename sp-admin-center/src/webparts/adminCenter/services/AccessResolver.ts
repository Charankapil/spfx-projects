import { entityOf, HttpError, odataString, rowsOf, SPClient, trimSlash } from '../core/SPClient';

/**
 * Turns any SharePoint address a person might paste - a file, folder, library, list,
 * list item, page, site, an "Access denied" page, an Office "Doc.aspx" link, a
 * library view with ?id=, or a sharing link - into the exact securable object
 * (web, list or item) and its permission chain up to the site collection, so access
 * can be granted on the object itself or on the parent it inherits from.
 *
 * Costs a handful of paced requests per lookup (roughly: 1 to find the web, 1-3 to
 * identify the object, 1 per parent folder and subsite level). Nothing is enumerated.
 */

export type SecurableKind = 'site' | 'web' | 'library' | 'list' | 'folder' | 'file' | 'item';

export interface ISecurable {
  kind: SecurableKind;
  title: string;
  /** Server-relative path, for display. */
  path: string;
  /** Web (subsite or root) the object lives in. */
  webUrl: string;
  /** REST base for this object, e.g. {web}/_api/web/lists(guid'..')/items(5). */
  api: string;
  listId?: string;
  itemId?: number;
  /** undefined when it could not be read. */
  hasUnique?: boolean;
  canManage?: boolean;
  accessible: boolean;
  error?: string;
}

export interface IParsedUrl {
  /** Absolute URL (no query) of the object the address points at. */
  url: string;
  /** Server-relative, decoded path. */
  path: string;
  notes: string[];
  fileId?: string;
  listId?: string;
  itemId?: number;
  /** The address is a view or form page of a list: the list itself is meant. */
  listPage?: boolean;
  /** _layouts or a site-level page: the web is meant. */
  webPage?: boolean;
  /** An encoded sharing link (/:w:/s/...), resolved through the shares API. */
  sharingLink?: string;
}

export interface IResolution {
  input: string;
  parsed: IParsedUrl;
  /** Site collection root first, target last. */
  chain: ISecurable[];
  target: ISecurable;
  /** Where the target's permissions come from (itself when it has unique permissions). */
  source?: ISecurable;
  notes: string[];
}

const GUID = /\{?([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\}?/i;

function guidOf(v: string | null | undefined): string | undefined {
  const m = GUID.exec(v || '');
  return m ? m[1].toLowerCase() : undefined;
}

/** "w2c1d...32 hex" (the ?d= parameter of Office links) to a GUID. */
function guidFromHex(v: string | null): string | undefined {
  const m = /^[a-z]?([0-9a-f]{32})$/i.exec(v || '');
  if (!m) {
    return undefined;
  }
  const h = m[1].toLowerCase();
  return `${h.substr(0, 8)}-${h.substr(8, 4)}-${h.substr(12, 4)}-${h.substr(16, 4)}-${h.substr(20)}`;
}

function safeDecode(s: string): string {
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
}

/** Turns "Shared%20Documents/a b" style paths into a clean server-relative path. */
function cleanPath(p: string): string {
  const d = safeDecode(p).replace(/\/{2,}/g, '/');
  return d.length > 1 ? d.replace(/\/+$/, '') : d;
}

/**
 * Pure: works out what an address points at without calling SharePoint.
 * Unwraps Access Denied / Request Access pages (Source=...), Office viewer links,
 * library view links (?id= / RootFolder=), item forms (?ID=) and /:x:/r/ links.
 */
export function parseAccessUrl(input: string, origin: string): IParsedUrl {
  const notes: string[] = [];
  let raw = (input || '').trim().replace(/^[<"']+|[>"']+$/g, '');
  if (raw.charAt(0) === '/') {
    raw = origin + raw;
  }
  if (!/^https:\/\//i.test(raw)) {
    throw new Error('Paste a full https:// address from this tenant.');
  }
  let listId: string | undefined;
  let itemId: number | undefined;
  for (let guard = 0; guard < 5; guard++) {
    let u: URL;
    try {
      u = new URL(raw);
    } catch {
      throw new Error('That does not look like a valid address.');
    }
    const lower = u.pathname.toLowerCase();
    const param = (name: string): string | null => {
      let found: string | null = null;
      u.searchParams.forEach((v, k) => {
        if (found === null && k.toLowerCase() === name.toLowerCase()) {
          found = v;
        }
      });
      return found;
    };
    const wrapped = /\/(accessdenied|requestaccess|accessrequests|error)\.aspx$/i.test(lower) || /\/_layouts\/15\/(start|authenticate)\.aspx$/i.test(lower);
    const source = param('Source') || param('ReturnUrl') || param('destination');
    if (wrapped && source) {
      notes.push(/accessdenied/i.test(lower) ? 'Unwrapped the original address from an "Access denied" page.' : 'Unwrapped the original address from a SharePoint system page.');
      listId = guidOf(param('name')) || listId;
      const li = parseInt(param('listItemId') || '', 10);
      if (!isNaN(li)) {
        itemId = li;
      }
      raw = /^https:\/\//i.test(source) ? source : u.origin + (source.charAt(0) === '/' ? '' : '/') + source;
      continue;
    }
    if (wrapped && listId) {
      // Access denied with only list / item ids: the list's web is the page's web.
      return { url: trimSlash(u.origin + u.pathname.replace(/\/_layouts\/.*$/i, '')), path: cleanPath(u.pathname.replace(/\/_layouts\/.*$/i, '')) || '/', notes, listId, itemId };
    }
    const result = parsePlain(u, param, notes);
    if (listId && !result.listId) {
      result.listId = listId;
    }
    if (itemId !== undefined && result.itemId === undefined) {
      result.itemId = itemId;
    }
    return result;
  }
  throw new Error('That address wraps too many redirects to follow.');
}

function parsePlain(u: URL, param: (n: string) => string | null, notes: string[]): IParsedUrl {
  let path = u.pathname;
  // Short links: /:w:/r/sites/x/... (direct) and /:w:/s/x/<token> (sharing link).
  const short = /^\/:([a-z]):\/([a-z])\/(.*)$/i.exec(path);
  if (short) {
    if (short[2].toLowerCase() !== 'r') {
      notes.push('This is a sharing link; resolving the file or folder it opens.');
      return { url: u.origin + u.pathname, path: cleanPath(u.pathname), notes, sharingLink: u.href };
    }
    path = '/' + short[3];
    notes.push('Opened from an Office short link.');
  }
  const out = (p: string, extra: Partial<IParsedUrl> = {}): IParsedUrl => {
    const clean = cleanPath(p) || '/';
    return { url: trimSlash(u.origin + encodeURI(clean)), path: clean, notes, ...extra };
  };

  const d = guidFromHex(param('d'));
  const lowerPath = path.toLowerCase();
  const layouts = lowerPath.indexOf('/_layouts/');
  if (layouts >= 0) {
    const webPath = path.substring(0, layouts) || '/';
    const sourcedoc = guidOf(param('sourcedoc'));
    if (sourcedoc) {
      notes.push(`Office link to "${param('file') || 'a file'}".`);
      return out(webPath, { fileId: sourcedoc });
    }
    const listParam = guidOf(param('List'));
    if (listParam) {
      const id = parseInt(param('ID') || param('Id') || '', 10);
      return out(webPath, { listId: listParam, itemId: isNaN(id) ? undefined : id, listPage: isNaN(id) });
    }
    return out(webPath, { webPage: true });
  }
  if (d) {
    return out(path, { fileId: d });
  }
  // Library / list views and forms: .../Forms/AllItems.aspx and .../Lists/Name/AllItems.aspx.
  const folderParam = param('id') || param('RootFolder');
  const libView = /^(.*)\/forms\/[^/]+\.aspx$/i.exec(path);
  const listView = /^(.*\/lists\/[^/]+)\/[^/]+\.aspx$/i.exec(path);
  const listPath = libView ? libView[1] : listView ? listView[1] : undefined;
  if (listPath) {
    if (folderParam && folderParam.charAt(0) === '/') {
      notes.push('Opened from a library view; using the folder it shows.');
      return out(folderParam);
    }
    const id = parseInt(param('ID') || '', 10);
    if (!isNaN(id) && /(disp|edit)form\.aspx$/i.test(path)) {
      return out(listPath, { itemId: id, listPage: true });
    }
    return out(listPath, { listPage: true });
  }
  if (folderParam && folderParam.charAt(0) === '/' && /\.aspx$/i.test(path)) {
    notes.push('Using the folder from the address.');
    return out(folderParam);
  }
  return out(path);
}

// ---- permissions helpers -------------------------------------------------------------

/** PermissionKind.ManagePermissions (25) in an EffectiveBasePermissions {Low, High} pair. */
export function canManagePermissions(bp: { Low?: unknown; High?: unknown } | undefined): boolean {
  if (!bp) {
    return false;
  }
  const low = Number(bp.Low) || 0;
  return Math.floor(low / Math.pow(2, 24)) % 2 === 1;
}

export interface IRoleDef {
  id: number;
  name: string;
  description: string;
  order: number;
}

export interface IAssignment {
  principalId: number;
  title: string;
  login: string;
  /** 1 user, 4 security group, 8 SharePoint group. */
  principalType: number;
  roles: string[];
}

function str(v: unknown): string {
  return v === undefined || v === null ? '' : String(v);
}

export function isAccessError(e: unknown): boolean {
  return e instanceof HttpError && (e.status === 401 || e.status === 403);
}

function isMissing(e: unknown): boolean {
  return e instanceof HttpError && (e.status === 404 || e.status === 400 || e.status === 500);
}

const ITEM_SELECT = '$select=Id,Title,FileLeafRef,FileRef,FileSystemObjectType,HasUniqueRoleAssignments,EffectiveBasePermissions,ParentList/Id,ParentList/Title,ParentList/BaseType&$expand=ParentList';

interface IRawItem {
  Id?: number;
  Title?: string;
  FileLeafRef?: string;
  FileRef?: string;
  FileSystemObjectType?: number;
  HasUniqueRoleAssignments?: boolean;
  EffectiveBasePermissions?: { Low?: string; High?: string };
  ParentList?: { Id?: string; Title?: string; BaseType?: number };
}

interface IRawList {
  Id?: string;
  Title?: string;
  BaseType?: number;
  HasUniqueRoleAssignments?: boolean;
  EffectiveBasePermissions?: { Low?: string; High?: string };
  RootFolder?: { ServerRelativeUrl?: string };
}

export class AccessResolver {
  private origin: string;

  constructor(private client: SPClient, private homeWebUrl: string) {
    this.origin = (/^(https:\/\/[^/]+)/i.exec(homeWebUrl) || [''])[1];
  }

  private pathParam(path: string): string {
    return `@p='${encodeURIComponent(odataString(path))}'`;
  }

  public async resolve(input: string): Promise<IResolution> {
    let parsed = parseAccessUrl(input, this.origin);
    const notes = parsed.notes.slice();
    this.client.assertSameOrigin(parsed.url);

    let shareIds: { webUrl?: string; listId?: string; listItemId?: string } | undefined;
    if (parsed.sharingLink) {
      shareIds = await this.resolveSharingLink(parsed.sharingLink);
      parsed = { ...parsed, url: shareIds.webUrl || parsed.url, listId: shareIds.listId, itemId: Number(shareIds.listItemId) || undefined };
    }

    const webUrl = shareIds && shareIds.webUrl ? trimSlash(shareIds.webUrl) : await this.findWeb(parsed.url);
    const webPath = this.serverRelative(webUrl);
    const webs = await this.webChain(webUrl);
    const chain: ISecurable[] = webs.slice();
    const web = webs[webs.length - 1];

    if (web.accessible) {
      const rest = parsed.path.length > webPath.length ? parsed.path.substring(webPath.length === 1 ? 0 : webPath.length) : '';
      const inner = await this.resolveInWeb(webUrl, parsed, rest, notes);
      inner.forEach((o) => chain.push(o));
    } else if (parsed.path !== webPath && trimSlash(parsed.path) !== trimSlash(webPath)) {
      notes.push('You cannot open this site, so the exact item inside it could not be identified. Add yourself as a site collection administrator (SharePoint admin center > Active sites > the site > Permissions), then try again.');
    }

    const target = chain[chain.length - 1];
    let source: ISecurable | undefined;
    for (let i = chain.length - 1; i >= 0; i--) {
      if (chain[i].hasUnique) {
        source = chain[i];
        break;
      }
      if (chain[i].hasUnique === undefined) {
        break;
      }
    }
    return { input, parsed, chain, target, source, notes };
  }

  private serverRelative(absUrl: string): string {
    const m = /^https:\/\/[^/]+(\/[^?#]*)?/i.exec(absUrl);
    return cleanPath(m && m[1] ? m[1] : '/') || '/';
  }

  /** Sharing links: SharePoint's shares endpoint returns the target's web, list and item ids. */
  private async resolveSharingLink(link: string): Promise<{ webUrl?: string; listId?: string; listItemId?: string }> {
    let b64: string;
    try {
      b64 = btoa(link).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
    } catch {
      throw new Error('Could not read this sharing link. Open it in the browser and paste the address it lands on instead.');
    }
    try {
      const j = await this.client.get<{ sharepointIds?: { webUrl?: string; listId?: string; listItemId?: string } }>(
        `${this.origin}/_api/v2.0/shares/u!${b64}/driveItem?$select=id,name,webUrl,sharepointIds`,
        { cacheTtlMs: 0 }
      );
      if (j && j.sharepointIds && j.sharepointIds.webUrl) {
        return j.sharepointIds;
      }
    } catch (e) {
      if (isAccessError(e)) {
        throw new Error('This sharing link does not give you access, so it cannot be resolved. Open the link in the browser and paste the address it lands on (or the file\u2019s path) instead.');
      }
    }
    throw new Error('Could not resolve this sharing link. Open it in the browser and paste the address it lands on instead.');
  }

  /** The web (site or subsite) an address belongs to. */
  private async findWeb(url: string): Promise<string> {
    try {
      const r = await this.client.post<{ value?: string; d?: { GetWebUrlFromPageUrl?: string } }>(
        `${trimSlash(this.homeWebUrl)}/_api/SP.Web.GetWebUrlFromPageUrl`,
        { pageFullUrl: url },
        { keepCache: true },
        this.homeWebUrl
      );
      const v = r && (r.value || (r.d && r.d.GetWebUrlFromPageUrl));
      if (typeof v === 'string' && /^https:\/\//i.test(v)) {
        return trimSlash(v);
      }
    } catch {
      /* fall back to probing */
    }
    // Probe from the shortest path up: the deepest /_api/web that answers with its own URL is the web.
    // The first address that refuses access is a site you cannot open.
    const m = /^(https:\/\/[^/]+)(\/.*)?$/i.exec(url);
    const originPart = m ? m[1] : this.origin;
    const segs = (m && m[2] ? m[2] : '').split('/').filter((x) => !!x);
    let found: string | undefined;
    for (let n = 0; n <= Math.min(segs.length, 8); n++) {
      if (n && /\.[a-z0-9]{2,5}$/i.test(segs[n - 1])) {
        break;
      }
      const candidate = originPart + (n ? '/' + segs.slice(0, n).join('/') : '');
      try {
        const j = entityOf<{ Url?: string }>(await this.client.get(`${candidate}/_api/web?$select=Url`, { cacheTtlMs: 60000 }));
        // Compare decoded: the candidate comes from an encoded address, Url is returned decoded.
        if (j && j.Url && safeDecode(trimSlash(j.Url)).toLowerCase() === safeDecode(candidate).toLowerCase()) {
          found = trimSlash(j.Url);
        }
      } catch (e) {
        // A refusal at the tenant root (n = 0) says nothing about the site that was pasted: keep probing.
        if (isAccessError(e) && n > 0) {
          return candidate;
        }
      }
    }
    if (found) {
      return found;
    }
    throw new Error('No SharePoint site was found at that address.');
  }

  /** Site collection root down to the given web. */
  private async webChain(webUrl: string): Promise<ISecurable[]> {
    const out: ISecurable[] = [];
    let current: string | undefined = webUrl;
    for (let guard = 0; current && guard < 12; guard++) {
      const at: string = current;
      const api = `${trimSlash(at)}/_api/web`;
      try {
        const w = entityOf<{ Title?: string; Url?: string; ServerRelativeUrl?: string; HasUniqueRoleAssignments?: boolean; EffectiveBasePermissions?: { Low?: string }; ParentWeb?: { ServerRelativeUrl?: string } }>(
          await this.client.get(`${api}?$select=Title,Url,ServerRelativeUrl,HasUniqueRoleAssignments,EffectiveBasePermissions,ParentWeb/ServerRelativeUrl&$expand=ParentWeb`, { cacheTtlMs: 0 })
        );
        const parentPath = w.ParentWeb && w.ParentWeb.ServerRelativeUrl;
        out.unshift({
          kind: parentPath ? 'web' : 'site',
          title: str(w.Title) || current,
          path: str(w.ServerRelativeUrl) || this.serverRelative(current),
          webUrl: trimSlash(current),
          api,
          hasUnique: parentPath ? !!w.HasUniqueRoleAssignments : true,
          canManage: canManagePermissions(w.EffectiveBasePermissions),
          accessible: true
        });
        current = parentPath ? this.origin + (parentPath === '/' ? '' : parentPath) : undefined;
      } catch (e) {
        out.unshift({ kind: 'web', title: this.serverRelative(at), path: this.serverRelative(at), webUrl: trimSlash(at), api, accessible: false, error: (e as Error).message });
        break;
      }
    }
    if (out.length && !out[0].accessible) {
      // We could not read the top: it may still be the root of its site collection.
      out[0].kind = 'site';
    }
    return out;
  }

  private listSecurable(webUrl: string, l: IRawList): ISecurable {
    const id = str(l.Id).toLowerCase();
    return {
      kind: l.BaseType === 1 ? 'library' : 'list',
      title: str(l.Title),
      path: cleanPath(str(l.RootFolder && l.RootFolder.ServerRelativeUrl)),
      webUrl,
      api: `${webUrl}/_api/web/lists(guid'${id}')`,
      listId: id,
      hasUnique: !!l.HasUniqueRoleAssignments,
      canManage: canManagePermissions(l.EffectiveBasePermissions),
      accessible: true
    };
  }

  private async getListById(webUrl: string, listId: string): Promise<ISecurable> {
    const l = entityOf<IRawList>(
      await this.client.get(`${webUrl}/_api/web/lists(guid'${listId}')?$select=Id,Title,BaseType,HasUniqueRoleAssignments,EffectiveBasePermissions,RootFolder/ServerRelativeUrl&$expand=RootFolder`, { cacheTtlMs: 0 })
    );
    return this.listSecurable(webUrl, l);
  }

  private async getListByPath(webUrl: string, path: string): Promise<ISecurable | undefined> {
    try {
      const l = entityOf<IRawList>(
        await this.client.get(`${webUrl}/_api/web/GetList(@p)?${this.pathParam(path)}&$select=Id,Title,BaseType,HasUniqueRoleAssignments,EffectiveBasePermissions,RootFolder/ServerRelativeUrl&$expand=RootFolder`, { cacheTtlMs: 0 })
      );
      if (!l || !l.Id) {
        return undefined;
      }
      const sec = this.listSecurable(webUrl, l);
      return sec.path.toLowerCase() === cleanPath(path).toLowerCase() ? sec : undefined;
    } catch (e) {
      if (isAccessError(e)) {
        throw e;
      }
      return undefined;
    }
  }

  private itemSecurable(webUrl: string, listId: string, it: IRawItem, isLibrary: boolean): ISecurable {
    const kind: SecurableKind = it.FileSystemObjectType === 1 ? 'folder' : isLibrary ? 'file' : 'item';
    return {
      kind,
      title: str(kind === 'item' ? it.Title || it.FileLeafRef : it.FileLeafRef || it.Title) || `Item ${it.Id}`,
      path: cleanPath(str(it.FileRef)),
      webUrl,
      api: `${webUrl}/_api/web/lists(guid'${listId}')/items(${Number(it.Id)})`,
      listId,
      itemId: Number(it.Id),
      hasUnique: !!it.HasUniqueRoleAssignments,
      canManage: canManagePermissions(it.EffectiveBasePermissions),
      accessible: true
    };
  }

  /** Reads an item by its file or folder path. undefined = no such file / folder. */
  private async itemByPath(webUrl: string, path: string, as: 'File' | 'Folder'): Promise<IRawItem | undefined> {
    try {
      const it = entityOf<IRawItem>(
        await this.client.get(`${webUrl}/_api/web/Get${as}ByServerRelativePath(decodedurl=@p)/ListItemAllFields?${this.pathParam(path)}&${ITEM_SELECT}`, { cacheTtlMs: 0 })
      );
      return it && it.Id !== undefined && it.Id !== null ? it : undefined;
    } catch (e) {
      if (isAccessError(e)) {
        throw e;
      }
      if (isMissing(e)) {
        return undefined;
      }
      throw e;
    }
  }

  /** List, parent folders and the object itself, inside one web. */
  private async resolveInWeb(webUrl: string, parsed: IParsedUrl, rest: string, notes: string[]): Promise<ISecurable[]> {
    const wrap = async (fn: () => Promise<ISecurable[]>, what: string): Promise<ISecurable[]> => {
      try {
        return await fn();
      } catch (e) {
        if (isAccessError(e)) {
          notes.push(`You do not have access to ${what}. Shown up to the deepest level you can open; a site collection administrator can grant it.`);
          return [];
        }
        throw e;
      }
    };

    // 1. Office link with a file id.
    if (parsed.fileId) {
      return wrap(async () => {
        const it = entityOf<IRawItem>(await this.client.get(`${webUrl}/_api/web/GetFileById('${parsed.fileId}')/ListItemAllFields?${ITEM_SELECT}`, { cacheTtlMs: 0 }));
        return this.withParents(webUrl, it);
      }, 'that file');
    }
    // 2. List (+ item) by id (Access denied pages, _layouts links, sharing links).
    if (parsed.listId && (parsed.itemId !== undefined || parsed.webPage || !rest || rest === '/')) {
      return wrap(async () => {
        const list = await this.getListById(webUrl, parsed.listId as string);
        if (parsed.itemId === undefined) {
          return [list];
        }
        const it = entityOf<IRawItem>(await this.client.get(`${list.api}/items(${parsed.itemId})?${ITEM_SELECT}`, { cacheTtlMs: 0 }));
        return this.withParents(webUrl, it);
      }, 'that list or item');
    }
    if (parsed.webPage || !rest || rest === '/') {
      return [];
    }
    const full = parsed.path;
    // 3. A list / library address or a view / form of one.
    if (parsed.listPage) {
      return wrap(async () => {
        const list = await this.getListByPath(webUrl, full);
        if (!list) {
          notes.push('No list or library was found at that address; showing the site.');
          return [];
        }
        if (parsed.itemId === undefined) {
          return [list];
        }
        const it = entityOf<IRawItem>(await this.client.get(`${list.api}/items(${parsed.itemId})?${ITEM_SELECT}`, { cacheTtlMs: 0 }));
        return this.withParents(webUrl, it);
      }, 'that list');
    }
    // 4. A path: file, folder or list root. If nothing is there (renamed, deleted, typo),
    //    walk up to the nearest folder or library that exists.
    return wrap(async () => {
      let p = full;
      const webPath = this.serverRelative(webUrl);
      for (let guard = 0; guard < 12 && p.length > webPath.length; guard++) {
        const file = await this.itemByPath(webUrl, p, 'File');
        if (file) {
          return this.withParents(webUrl, file);
        }
        const folder = await this.itemByPath(webUrl, p, 'Folder');
        if (folder) {
          return this.withParents(webUrl, folder);
        }
        const list = await this.getListByPath(webUrl, p);
        if (list) {
          return [list];
        }
        const up = p.substring(0, p.lastIndexOf('/'));
        if (guard === 0) {
          notes.push(`Nothing exists at exactly ${full} (moved, renamed or deleted?). Showing the nearest parent that exists.`);
        }
        p = up;
      }
      return [];
    }, 'that item');
  }

  /** The list and every parent folder of an item, then the item. */
  private async withParents(webUrl: string, it: IRawItem): Promise<ISecurable[]> {
    const listId = str(it.ParentList && it.ParentList.Id).toLowerCase();
    if (!listId) {
      throw new HttpError(404, 'Item not found.', webUrl);
    }
    const list = await this.getListById(webUrl, listId);
    const isLibrary = list.kind === 'library';
    const item = this.itemSecurable(webUrl, listId, it, isLibrary);
    const folders: ISecurable[] = [];
    let parent = item.path.substring(0, item.path.lastIndexOf('/'));
    for (let guard = 0; guard < 15 && parent.length > list.path.length && parent.toLowerCase().indexOf(list.path.toLowerCase() + '/') === 0; guard++) {
      try {
        const f = await this.itemByPath(webUrl, parent, 'Folder');
        if (f) {
          folders.unshift(this.itemSecurable(webUrl, listId, f, isLibrary));
        }
      } catch (e) {
        if (!isAccessError(e)) {
          throw e;
        }
        folders.unshift({ kind: 'folder', title: parent.split('/').pop() || parent, path: parent, webUrl, api: '', listId, accessible: false, error: (e as Error).message });
      }
      parent = parent.substring(0, parent.lastIndexOf('/'));
    }
    return [list as ISecurable].concat(folders, [item]);
  }

  // ---- reading and granting ------------------------------------------------------------

  public async roleDefinitions(webUrl: string): Promise<IRoleDef[]> {
    const j = await this.client.get(`${trimSlash(webUrl)}/_api/web/roledefinitions?$select=Id,Name,Description,Hidden,RoleTypeKind,Order`);
    return rowsOf<{ Id?: number; Name?: string; Description?: string; Hidden?: boolean; RoleTypeKind?: number; Order?: number }>(j)
      .filter((r) => !r.Hidden && r.RoleTypeKind !== 1)
      .map((r) => ({ id: Number(r.Id), name: str(r.Name), description: str(r.Description), order: Number(r.Order) || 0 }))
      .sort((a, b) => a.order - b.order);
  }

  public async assignments(o: ISecurable): Promise<IAssignment[]> {
    const j = await this.client.get(`${o.api}/roleassignments?$select=PrincipalId,Member/Title,Member/LoginName,Member/PrincipalType,RoleDefinitionBindings/Name,RoleDefinitionBindings/RoleTypeKind&$expand=Member,RoleDefinitionBindings&$top=1000`, { cacheTtlMs: 0 });
    return rowsOf<{ [k: string]: unknown }>(j).map((r) => {
      const m = (r.Member || {}) as { [k: string]: unknown };
      return {
        principalId: Number(r.PrincipalId),
        title: str(m.Title),
        login: str(m.LoginName),
        principalType: Number(m.PrincipalType) || 0,
        roles: rowsOf<{ Name?: string; RoleTypeKind?: number }>(r.RoleDefinitionBindings)
          .filter((x) => x.RoleTypeKind !== 1)
          .map((x) => str(x.Name))
      };
    });
  }

  /** Gives a principal a permission level on the object; stops inheriting first when asked (copying current permissions). */
  public async grant(o: ISecurable, principalId: number, roleDefId: number, breakInheritance: boolean): Promise<void> {
    if (!o.accessible || !o.api) {
      throw new Error('You cannot open this object, so permissions cannot be changed on it.');
    }
    if (breakInheritance && !o.hasUnique) {
      await this.client.post(`${o.api}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)`, undefined, {}, o.webUrl);
      o.hasUnique = true;
    }
    if (!o.hasUnique) {
      throw new Error(`${o.title} inherits its permissions. Grant on the parent it inherits from, or choose to stop inheriting.`);
    }
    await this.client.post(`${o.api}/roleassignments/addroleassignment(principalid=${principalId},roledefid=${roleDefId})`, undefined, {}, o.webUrl);
  }
}
