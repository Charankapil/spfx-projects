import { CancelledError, HttpError, odataString, rowsOf, SPClient, trimSlash } from '../core/SPClient';
import { groupClaims, ISubject, isM365Login, isPersonLogin, NO_ACCESS, permissionLevel } from './permissions';

/**
 * "Where can these people and groups get in?" - one site at a time.
 *
 * A person: SharePoint's own effective-permission check, which counts direct grants,
 * SharePoint groups, Entra ID / Microsoft 365 groups, Everyone claims and site admin rights.
 * One request for the site and one more per library that has its own permissions.
 *
 * A group (or any other claim): one request looks the claim up in the site's user list, and
 * most sites end there. Where it is known, the site's permissions are read and every grant
 * that reaches the group - directly or through a SharePoint group it is in - is one row.
 *
 * The site's libraries are listed once and shared by every subject.
 *
 * The Power Automate flow (flow/definition.ts) performs the same requests and writes the same
 * rows in the same order; tests/flow.test.js runs both against one fake tenant and compares
 * the files byte for byte. Change one, change the other.
 */

export interface ISite {
  url: string;
  title: string;
}

export interface ISiteResult {
  /** Report rows, in REPORT_HEADERS order. */
  rows: string[][];
  /** What could not be read, in UNREAD_HEADERS order. */
  unread: string[][];
}

/** A web to scan: a site collection's top-level site or one of its subsites. */
export interface IWeb extends ISite {
  object: 'Site' | 'Subsite';
  /** Subsites: Yes when it does not inherit permissions from its parent; '' for a top-level site. */
  unique: string;
  /** The site collection's URL. */
  collection: string;
}

interface IRawWeb {
  Title?: string;
  Url?: string;
  HasUniqueRoleAssignments?: boolean;
}

interface IRawList {
  Id?: string;
  Title?: string;
  BaseType?: number;
  HasUniqueRoleAssignments?: boolean;
  IsSystemList?: boolean;
  RootFolder?: { ServerRelativeUrl?: string };
}

interface IRawAssignment {
  PrincipalId?: number;
  Member?: { Title?: string; PrincipalType?: number };
  RoleDefinitionBindings?: unknown;
}

interface IRawSiteUser {
  Id?: number;
  LoginName?: string;
  IsSiteAdmin?: boolean;
  Groups?: unknown;
}

interface ILibrary {
  id: string;
  title: string;
  url: string;
}

export const ALL_SUBJECTS = '(all)';
const SP_GROUP = 8;

const WEBS_QUERY = '$select=Title,Url,HasUniqueRoleAssignments';
const LISTS_QUERY = '$select=Id,Title,BaseType,HasUniqueRoleAssignments,IsSystemList,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=Hidden eq false&$top=5000';
const ASSIGNMENTS_QUERY = '$select=PrincipalId,Member/Title,Member/PrincipalType,RoleDefinitionBindings/Name,RoleDefinitionBindings/RoleTypeKind&$expand=Member,RoleDefinitionBindings&$top=5000';
const SITE_USER_QUERY = '$select=Id,LoginName,IsSiteAdmin,Groups/Id,Groups/Title&$expand=Groups';

/** Permission level names on an assignment, Limited Access left out. */
export function roleNames(a: IRawAssignment): string[] {
  return rowsOf<{ Name?: string; RoleTypeKind?: number }>(a.RoleDefinitionBindings)
    .filter((r) => r.RoleTypeKind !== 1 && !/^limited access/i.test(String(r.Name || '')))
    .map((r) => String(r.Name || ''));
}

/** The assignments that reach a principal: made to it, or to a SharePoint group it is in. */
export function grantsFor(assignments: IRawAssignment[], principalId: number, spGroupIds: number[]): Array<{ level: string; how: string }> {
  const out: Array<{ level: string; how: string }> = [];
  assignments.forEach((a) => {
    const pid = Number(a.PrincipalId);
    if (pid !== principalId && spGroupIds.indexOf(pid) < 0) {
      return;
    }
    const names = roleNames(a);
    if (!names.length) {
      return;
    }
    const m = a.Member || {};
    out.push({ level: names.join(', '), how: Number(m.PrincipalType) === SP_GROUP ? `SharePoint group: ${m.Title || ''}` : 'Direct' });
  });
  return out;
}

/** The mask of a getusereffectivepermissions response, whichever OData flavour it came in. */
export function levelOf(json: unknown): string {
  const o = (json || {}) as { [k: string]: unknown };
  const d = (o.d ? o.d : o) as { [k: string]: unknown };
  const bp = ((d && (d.GetUserEffectivePermissions || d.value)) || d) as { Low?: unknown; High?: unknown };
  return permissionLevel(Number(bp.Low) || 0, Number(bp.High) || 0);
}

function statusOf(e: unknown): string {
  return e instanceof HttpError ? String(e.status) : '';
}

function messageOf(e: unknown): string {
  return ((e as Error).message || String(e)).replace(/[\r\n]+/g, ' ').substring(0, 300);
}

export class AccessScanner {
  private origin: string;

  constructor(private client: SPClient, homeWebUrl: string) {
    this.origin = (/^(https:\/\/[^/]+)/i.exec(homeWebUrl) || [''])[1];
  }

  private get(url: string): Promise<unknown> {
    return this.client.get(url, { cacheTtlMs: 0 });
  }

  /**
   * Every web of a site collection, the top-level site first and then its subsites level by
   * level (breadth first): one request per web lists its direct subsites.
   */
  public async webs(site: ISite, unread: string[][]): Promise<IWeb[]> {
    const root = trimSlash(site.url);
    const webs: IWeb[] = [{ url: root, title: site.title, object: 'Site', unique: '', collection: root }];
    for (let i = 0; i < webs.length; i++) {
      const w = webs[i];
      try {
        rowsOf<IRawWeb>(await this.get(`${w.url}/_api/web/webs?${WEBS_QUERY}`)).forEach((x) => {
          const url = trimSlash(String(x.Url || ''));
          if (url && !webs.some((o) => o.url.toLowerCase() === url.toLowerCase())) {
            webs.push({ url, title: String(x.Title || '') || url, object: 'Subsite', unique: x.HasUniqueRoleAssignments === true ? 'Yes' : 'No', collection: root });
          }
        });
      } catch (e) {
        if (e instanceof CancelledError) {
          throw e;
        }
        unread.push([ALL_SUBJECTS, w.url, w.title, 'Subsites', statusOf(e), messageOf(e)]);
      }
    }
    return webs;
  }

  /** Everything the subjects can reach in one site collection, subsites included. Only a cancelled scan throws. */
  public async scanSite(site: ISite, subjects: ISubject[], includeLibraries: boolean): Promise<ISiteResult> {
    const out: ISiteResult = { rows: [], unread: [] };
    for (const w of await this.webs(site, out.unread)) {
      const r = await this.scanWeb(w, subjects, includeLibraries);
      r.rows.forEach((x) => out.rows.push(x));
      r.unread.forEach((x) => out.unread.push(x));
    }
    return out;
  }

  /** Everything the subjects can reach on one web (a site or a subsite) and its libraries. */
  public async scanWeb(site: IWeb, subjects: ISubject[], includeLibraries: boolean): Promise<ISiteResult> {
    const web = trimSlash(site.url);
    const out: ISiteResult = { rows: [], unread: [] };
    const unread = (who: string, step: string, e: unknown): void => {
      if (e instanceof CancelledError) {
        throw e;
      }
      out.unread.push([who, web, site.title, step, statusOf(e), messageOf(e)]);
    };

    const inherited: ILibrary[] = [];
    const own: ILibrary[] = [];
    if (includeLibraries) {
      try {
        rowsOf<IRawList>(await this.get(`${web}/_api/web/lists?${LISTS_QUERY}`))
          .filter((l) => l.BaseType === 1 && l.IsSystemList !== true)
          .forEach((l) => {
            const path = (l.RootFolder && l.RootFolder.ServerRelativeUrl) || '';
            const lib = { id: String(l.Id), title: String(l.Title || ''), url: path ? this.origin + path : '' };
            (l.HasUniqueRoleAssignments === true ? own : inherited).push(lib);
          });
      } catch (e) {
        unread(ALL_SUBJECTS, 'Libraries', e);
      }
    }

    // A row about the web itself says Site or Subsite, and for a subsite whether it has its own permissions.
    const row = (who: string, object: 'Site' | 'Library', name: string, url: string, level: string, how: string, unique: string): void => {
      const self = object === 'Site';
      out.rows.push([who, web, site.title, self ? site.object : object, name, url, level, how, self ? site.unique : unique, site.collection]);
    };

    // People first, then groups - the order the flow works in.
    for (const s of subjects.filter((x) => isPersonLogin(x.login))) {
      await this.person(web, site, s, inherited, own, row, unread);
    }
    for (const s of subjects.filter((x) => !isPersonLogin(x.login))) {
      await this.group(web, site, s, inherited, own, row, unread);
    }
    return out;
  }

  private async person(
    web: string,
    site: ISite,
    s: ISubject,
    inherited: ILibrary[],
    own: ILibrary[],
    row: (who: string, object: 'Site' | 'Library', name: string, url: string, level: string, how: string, unique: string) => void,
    unread: (who: string, step: string, e: unknown) => void
  ): Promise<void> {
    const at = `@v='${encodeURIComponent(odataString(s.login.trim()))}'`;
    let siteLevel: string;
    try {
      siteLevel = levelOf(await this.get(`${web}/_api/web/getusereffectivepermissions(@v)?${at}`));
    } catch (e) {
      unread(s.label, 'Site', e);
      return;
    }
    if (siteLevel !== NO_ACCESS) {
      row(s.label, 'Site', site.title, web, siteLevel, 'Effective permission', '');
      inherited.forEach((l) => row(s.label, 'Library', l.title, l.url, siteLevel, 'Inherits from site', 'No'));
    }
    for (const l of own) {
      try {
        const level = levelOf(await this.get(`${web}/_api/web/lists(guid'${l.id}')/getusereffectivepermissions(@v)?${at}`));
        if (level !== NO_ACCESS) {
          row(s.label, 'Library', l.title, l.url, level, 'Effective permission (own permissions)', 'Yes');
        }
      } catch (e) {
        unread(s.label, `Library: ${l.title}`, e);
      }
    }
  }

  private async group(
    web: string,
    site: ISite,
    s: ISubject,
    inherited: ILibrary[],
    own: ILibrary[],
    row: (who: string, object: 'Site' | 'Library', name: string, url: string, level: string, how: string, unique: string) => void,
    unread: (who: string, step: string, e: unknown) => void
  ): Promise<void> {
    const login = s.login.trim();
    const filter = groupClaims(login)
      .map((c) => `LoginName eq '${odataString(c)}'`)
      .join(' or ');
    let found: IRawSiteUser[];
    try {
      // Is the claim known to this site at all? Most sites end here: one request.
      found = rowsOf<IRawSiteUser>(await this.get(`${web}/_api/web/siteusers?$filter=${encodeURIComponent(filter)}&${SITE_USER_QUERY}`));
    } catch (e) {
      unread(s.label, 'Site', e);
      return;
    }
    const m365 = isM365Login(login);
    for (const f of found) {
      const who = m365 && /_o$/i.test(String(f.LoginName || '')) ? `${s.label} (owners)` : s.label;
      const id = Number(f.Id);
      const spGroupIds = rowsOf<{ Id?: number }>(f.Groups).map((g) => Number(g.Id));
      if (f.IsSiteAdmin === true) {
        row(who, 'Site', site.title, web, 'Full control', 'Site collection administrator', '');
      }
      let siteGrants: Array<{ level: string; how: string }>;
      try {
        siteGrants = grantsFor(rowsOf<IRawAssignment>(await this.get(`${web}/_api/web/roleassignments?${ASSIGNMENTS_QUERY}`)), id, spGroupIds);
      } catch (e) {
        unread(who, 'Site permissions', e);
        continue;
      }
      siteGrants.forEach((g) => {
        row(who, 'Site', site.title, web, g.level, g.how, '');
        inherited.forEach((l) => row(who, 'Library', l.title, l.url, g.level, `Inherits from site - ${g.how}`, 'No'));
      });
      for (const l of own) {
        try {
          grantsFor(rowsOf<IRawAssignment>(await this.get(`${web}/_api/web/lists(guid'${l.id}')/roleassignments?${ASSIGNMENTS_QUERY}`)), id, spGroupIds).forEach((g) =>
            row(who, 'Library', l.title, l.url, g.level, g.how, 'Yes')
          );
        } catch (e) {
          unread(who, `Library: ${l.title}`, e);
        }
      }
    }
  }
}
