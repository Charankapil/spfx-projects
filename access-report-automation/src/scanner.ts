import { HttpError } from './http.js';
import {
  adminExplanation,
  Bits,
  bitsOf,
  EEEU_PREFIX,
  explainAssignments,
  Explanation,
  Link,
  NO_ACCESS,
  NO_BITS,
  permissionLevel,
  Principal,
  RawAssignment,
  SpGroupLink
} from './permissions.js';
import { alias, entityOf, odataString, originOf, SharePoint, trimSlash } from './sharepoint.js';

/**
 * "Where can these people and groups get in?" for one site collection.
 *
 * Everything that does not depend on the principal is read once per site and shared by all
 * of them (the site's group entries, each object's role assignments, the list of libraries
 * and subsites). Most sites cost two requests, because most sites have nothing to do with
 * the people in the report.
 *
 * A person's level comes from SharePoint's own effective-permission check; the role
 * assignments explain how they get it. When the two disagree the row says so.
 */

export type ObjectKind = 'Site' | 'Subsite' | 'Library' | 'List';

export interface AccessRow {
  /** Principal.key */
  p: string;
  siteUrl: string;
  siteTitle: string;
  object: ObjectKind;
  title: string;
  url: string;
  /** Full control, Design, Edit, Contribute, Read, View only. */
  access: string;
  /** Permission level names on the object, as the site calls them. */
  levels: string;
  how: string;
  /** The object has its own permissions (does not inherit). */
  unique: boolean;
  note: string;
}

export interface ScanOptions {
  includeLibraries: boolean;
  includeLists: boolean;
  includeSubsites: boolean;
  /** Ask SharePoint for the effective permission of every person on every site, even where nothing links them to it. */
  thorough: boolean;
  /** Group memberships were read from the directory (so "no link to this site" can be trusted). */
  groupsExpanded: boolean;
  maxSubsites: number;
}

interface RawSiteUser {
  Id?: number;
  LoginName?: string;
  Title?: string;
  IsSiteAdmin?: boolean;
  PrincipalType?: number;
}

interface RawList {
  Id?: string;
  Title?: string;
  BaseType?: number;
  HasUniqueRoleAssignments?: boolean;
  IsSystemList?: boolean;
  RootFolder?: { ServerRelativeUrl?: string };
}

interface RawWeb {
  Title?: string;
  Url?: string;
  HasUniqueRoleAssignments?: boolean;
}

/** One principal's standing on one object. */
interface Standing {
  access: string;
  levels: string;
  how: string;
  note: string;
}

interface Active {
  principal: Principal;
  links: Link[];
  spGroups: SpGroupLink[];
  admin?: Explanation;
}

const USER_BATCH = 10;
const UNCONFIRMED = 'Granted through a group, but not confirmed by SharePoint\'s effective-permission check (the person may not have signed in since joining the group).';
const UNEXPLAINED = 'Effective permission (the route was not identified)';

const SITE_USER_SELECT = '$select=Id,LoginName,Title,IsSiteAdmin,PrincipalType';
const LIST_SELECT = '$select=Id,Title,BaseType,HasUniqueRoleAssignments,IsSystemList,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=Hidden eq false&$top=5000';

export class SiteScanner {
  /** Per principal: claim (lower case) -> the group it stands for ('' = the principal itself). */
  private claims = new Map<string, Map<string, { via: string; implicit: boolean }>>();
  /** Things the reader of the report should know, one per principal at most. */
  public readonly warnings = new Map<string, string>();

  constructor(private sp: SharePoint, private principals: Principal[], private opts: ScanOptions) {
    principals.forEach((p) => {
      const m = new Map<string, { via: string; implicit: boolean }>();
      // The first identity for a claim wins: the principal's own claims come first.
      p.identities.forEach((i) => {
        const k = i.login.toLowerCase();
        if (!m.has(k)) {
          m.set(k, { via: i.via, implicit: !!i.implicit });
        }
      });
      this.claims.set(p.key, m);
    });
  }

  /**
   * Everything the principals can reach in one site collection.
   * @param skipWebs  lower-case addresses scanned on their own (not walked again as subsites).
   */
  public async scanSite(site: { url: string; title: string }, skipWebs?: Set<string>): Promise<AccessRow[]> {
    const root = trimSlash(site.url);
    const origin = originOf(root);
    const rows: AccessRow[] = [];

    // 1. Which of our principals does this site collection know, directly or through a group?
    const groupEntries = await this.groupEntries(root);
    const userEntries = await this.userEntries(root);
    const spGroupCache = new Map<number, Promise<Array<{ id: number; title: string }>>>();
    const spGroupsOf = (id: number): Promise<Array<{ id: number; title: string }>> => {
      let p = spGroupCache.get(id);
      if (!p) {
        p = this.sp
          .getAll<{ Id?: number; Title?: string }>(`${root}/_api/web/siteusers/getbyid(${id})/groups?$select=Id,Title`)
          .then((gs) => gs.map((g) => ({ id: Number(g.Id), title: String(g.Title) })));
        spGroupCache.set(id, p);
      }
      return p;
    };

    const actives: Active[] = [];
    const rootEffective = new Map<string, Bits>();
    for (const principal of this.principals) {
      const links: Link[] = [];
      if (principal.login) {
        const own = userEntries.get(principal.login.toLowerCase());
        if (own) {
          links.push({ id: Number(own.Id), via: '', isSiteAdmin: !!own.IsSiteAdmin });
        }
      }
      const claims = this.claims.get(principal.key) as Map<string, { via: string; implicit: boolean }>;
      for (const entry of groupEntries) {
        const login = String(entry.LoginName || '').toLowerCase();
        const hit = claims.get(login) ?? (login.indexOf(EEEU_PREFIX) === 0 ? claims.get(EEEU_PREFIX) : undefined);
        if (hit && !links.some((l) => l.id === Number(entry.Id))) {
          links.push({ id: Number(entry.Id), via: hit.via, implicit: hit.implicit, isSiteAdmin: !!entry.IsSiteAdmin });
        }
      }
      let active = links.length > 0;
      if (principal.kind === 'user' && !active) {
        if (!this.opts.groupsExpanded) {
          // Without directory group data only SharePoint can tell: look everywhere.
          active = true;
        } else if (this.opts.thorough) {
          const bits = await this.effective(`${root}/_api/web`, principal);
          rootEffective.set(principal.key, bits);
          active = permissionLevel(bits) !== NO_ACCESS;
        }
      }
      if (!active) {
        continue;
      }
      const spGroups: SpGroupLink[] = [];
      for (const l of links) {
        (await spGroupsOf(l.id)).forEach((g) => {
          if (!spGroups.some((x) => x.id === g.id && x.via === l.via)) {
            spGroups.push({ id: g.id, title: g.title, via: l.via, implicit: l.implicit });
          }
        });
      }
      actives.push({ principal, links, spGroups, admin: adminExplanation(links) });
    }
    if (!actives.length) {
      return rows;
    }

    // 2. Walk the site, its libraries and its subsites.
    const standingOn = async (api: string, assignments: RawAssignment[], a: Active, cachedBits?: Bits): Promise<Standing | undefined> => {
      if (a.admin) {
        return { access: 'Full control', levels: a.admin.levels.join(', '), how: a.admin.how.join('; '), note: '' };
      }
      const explained = explainAssignments(assignments, a.links, a.spGroups);
      if (a.principal.kind !== 'user') {
        if (!explained) {
          return undefined;
        }
        const access = permissionLevel(explained.bits);
        // A custom level that grants nothing readable is still worth showing by name.
        return { access: access === NO_ACCESS ? 'Custom' : access, levels: explained.levels.join(', '), how: explained.how.join('; '), note: '' };
      }
      const bits = cachedBits ?? (await this.effective(api, a.principal));
      const effective = permissionLevel(bits);
      if (effective !== NO_ACCESS) {
        return { access: effective, levels: explained ? explained.levels.join(', ') : '', how: explained ? explained.how.join('; ') : UNEXPLAINED, note: '' };
      }
      if (explained && permissionLevel(explained.bits) !== NO_ACCESS) {
        // Only directory-group routes can be invisible to the effective check; a direct grant, or one
        // through an "everyone" claim, that SharePoint denies is denied.
        const viaGroupOnly = explained.how.every((h) => /^Group: /.test(h));
        if (viaGroupOnly) {
          return { access: permissionLevel(explained.bits), levels: explained.levels.join(', '), how: explained.how.join('; '), note: UNCONFIRMED };
        }
      }
      return undefined;
    };

    const assignmentsOf = (api: string): Promise<RawAssignment[]> => this.sp.getAll<RawAssignment>(`${api}/roleassignments?$expand=RoleDefinitionBindings&$top=5000`);

    let websSeen = 0;
    const walkWeb = async (webUrl: string, title: string, isRoot: boolean, unique: boolean, inherited: Map<string, Standing>, depth: number): Promise<void> => {
      const api = `${webUrl}/_api/web`;
      const standings = new Map<string, Standing>();
      if (unique) {
        const assignments = await assignmentsOf(api);
        for (const a of actives) {
          const s = await standingOn(api, assignments, a, isRoot ? rootEffective.get(a.principal.key) : undefined);
          if (s) {
            standings.set(a.principal.key, s);
          }
        }
      } else {
        inherited.forEach((s, k) => standings.set(k, { ...s, how: 'Inherits from parent site', note: s.note }));
      }
      for (const a of actives) {
        const s = standings.get(a.principal.key);
        if (s) {
          rows.push({ p: a.principal.key, siteUrl: root, siteTitle: site.title, object: isRoot ? 'Site' : 'Subsite', title, url: webUrl, access: s.access, levels: s.levels, how: s.how, unique: isRoot || unique, note: s.note });
        }
      }

      if (this.opts.includeLibraries || this.opts.includeLists) {
        const lists = (await this.sp.getAll<RawList>(`${api}/lists?${LIST_SELECT}`)).filter((l) => !l.IsSystemList && ((l.BaseType === 1 && this.opts.includeLibraries) || (l.BaseType !== 1 && this.opts.includeLists)));
        for (const l of lists) {
          const object: ObjectKind = l.BaseType === 1 ? 'Library' : 'List';
          const path = l.RootFolder && l.RootFolder.ServerRelativeUrl ? l.RootFolder.ServerRelativeUrl : '';
          const url = path ? origin + encodePath(path) : webUrl;
          const push = (key: string, s: Standing, own: boolean): void => {
            rows.push({ p: key, siteUrl: root, siteTitle: site.title, object, title: String(l.Title || ''), url, access: s.access, levels: s.levels, how: s.how, unique: own, note: s.note });
          };
          if (!l.HasUniqueRoleAssignments) {
            standings.forEach((s, key) => push(key, { ...s, how: isRoot ? 'Inherits from site' : 'Inherits from subsite' }, false));
            continue;
          }
          const listApi = `${api}/lists(guid'${String(l.Id)}')`;
          let assignments: RawAssignment[];
          try {
            assignments = await assignmentsOf(listApi);
          } catch (e) {
            if (e instanceof HttpError && e.status === 404) {
              // Deleted while the site was being read.
              continue;
            }
            throw e;
          }
          for (const a of actives) {
            const s = await standingOn(listApi, assignments, a);
            if (s) {
              push(a.principal.key, s, true);
            }
          }
        }
      }

      if (this.opts.includeSubsites && depth < 15) {
        const subs = await this.sp.getAll<RawWeb>(`${api}/webs?$select=Title,Url&$top=5000`);
        for (const sub of subs) {
          const subUrl = trimSlash(String(sub.Url || ''));
          if (!subUrl || originOf(subUrl) !== origin || (skipWebs && skipWebs.has(subUrl.toLowerCase()))) {
            continue;
          }
          if (++websSeen > this.opts.maxSubsites) {
            throw new Error(`More than ${this.opts.maxSubsites} subsites; raise scan.maxSubsites to read this site`);
          }
          // HasUniqueRoleAssignments is only returned when asked of the web itself.
          const w = entityOf<RawWeb>(await this.sp.get(`${subUrl}/_api/web?$select=Title,Url,HasUniqueRoleAssignments`));
          await walkWeb(subUrl, String(w.Title || sub.Title || subUrl), false, !!w.HasUniqueRoleAssignments, standings, depth + 1);
        }
      }
    };

    await walkWeb(root, site.title, true, true, new Map(), 0);
    return rows;
  }

  /** The site's entries for security groups, Microsoft 365 groups and the Everyone claims. */
  private async groupEntries(root: string): Promise<RawSiteUser[]> {
    let entries: RawSiteUser[];
    try {
      entries = await this.sp.getAll<RawSiteUser>(`${root}/_api/web/siteusers?$filter=${encodeURIComponent('PrincipalType eq 4')}&${SITE_USER_SELECT}`);
    } catch (e) {
      if (!(e instanceof HttpError && e.status === 400)) {
        throw e;
      }
      // The filter was refused: read the whole list instead.
      entries = await this.sp.getAll<RawSiteUser>(`${root}/_api/web/siteusers?${SITE_USER_SELECT}`);
    }
    return entries.filter((e) => e.PrincipalType === undefined || e.PrincipalType === 4);
  }

  /** The site's entries for the people in the report, by login (lower case). */
  private async userEntries(root: string): Promise<Map<string, RawSiteUser>> {
    const logins = this.principals.filter((p) => !!p.login).map((p) => String(p.login));
    const out = new Map<string, RawSiteUser>();
    for (let i = 0; i < logins.length; i += USER_BATCH) {
      const filter = logins
        .slice(i, i + USER_BATCH)
        .map((l) => `LoginName eq '${odataString(l)}'`)
        .join(' or ');
      const found = await this.sp.getAll<RawSiteUser>(`${root}/_api/web/siteusers?$filter=${encodeURIComponent(filter)}&${SITE_USER_SELECT}`);
      found.forEach((f) => out.set(String(f.LoginName || '').toLowerCase(), f));
    }
    return out;
  }

  private async effective(api: string, principal: Principal): Promise<Bits> {
    let j: Record<string, unknown>;
    try {
      j = await this.sp.get<Record<string, unknown>>(`${api}/getusereffectivepermissions(@v)?${alias('v', String(principal.login))}`);
    } catch (e) {
      // The site answered everything else, so this is about the person: SharePoint does not know
      // the account (yet). That is "no access as far as SharePoint can tell", not a broken site.
      if (e instanceof HttpError && (e.status === 400 || e.status === 404)) {
        if (!this.warnings.has(principal.key)) {
          this.warnings.set(principal.key, `SharePoint could not evaluate ${principal.label} on at least one site (${e.message.substring(0, 160)}). Access through groups is still reported from the permissions themselves.`);
        }
        return NO_BITS;
      }
      throw e;
    }
    const d = (j && j.d ? j.d : j) as Record<string, unknown> | undefined;
    const bp = (d && (d.GetUserEffectivePermissions || d.value)) || d;
    return bitsOf(bp as { Low?: unknown; High?: unknown });
  }
}

/** A server-relative path as an address: encodeURI leaves # and ? alone, which are legal in library names. */
export function encodePath(path: string): string {
  return encodeURI(path).replace(/#/g, '%23').replace(/\?/g, '%3F');
}

/** Human text for a site that could not be read. */
export function scanError(e: unknown): string {
  if (e instanceof HttpError) {
    if (e.status === 401 || e.status === 403) {
      return 'Access refused (the site is locked, or the app has no permission here)';
    }
    if (e.status === 404) {
      return 'Site not found (deleted or renamed since the site list was read)';
    }
    if (e.status === 423) {
      return 'The site is locked or archived';
    }
    if (e.status >= 300 && e.status < 400) {
      return 'The site has moved (redirect)';
    }
    return `${e.status || 'Network'}: ${e.message.substring(0, 160)}`;
  }
  return ((e as Error).message || String(e)).substring(0, 200);
}

/** Is this failure about the one site (true) or a sign the whole run is in trouble (false)? */
export function isSiteProblem(e: unknown): boolean {
  if (e instanceof HttpError) {
    return [400, 403, 404, 410, 423].indexOf(e.status) >= 0 || (e.status >= 300 && e.status < 400);
  }
  return /subsites/.test((e as Error).message || '');
}
