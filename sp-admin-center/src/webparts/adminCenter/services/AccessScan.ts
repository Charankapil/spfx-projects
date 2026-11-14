import { HttpError, odataString, rowsOf, SPClient, trimSlash } from '../core/SPClient';
import { permissionLevel, toLogin } from './AdminApi';

/**
 * "Where does this person / group have access?" across many sites, one site at a time.
 *
 * Person: SharePoint's own effective-permission check (it counts direct grants, SharePoint
 * groups, AD / Microsoft 365 groups, Everyone claims and site admin rights) - one request for
 * the site, one listing of its libraries, and one more per library with its own permissions.
 *
 * Group: one request per site looks the group up in the site's user list; only sites where it
 * appears cost more (its SharePoint groups, the site's permissions, libraries with their own).
 *
 * Progress is saved in the browser so a scan stopped by a closed tab, a lost connection or
 * repeated errors carries on from where it was.
 */

export type PrincipalKind = 'user' | 'security' | 'm365';

export interface IScanPrincipal {
  kind: PrincipalKind;
  /** Shown in the UI and the CSV. */
  label: string;
  /** e-mail / UPN (user) or object id (group). */
  key: string;
}

export type ObjectKind = 'Site' | 'Library';

export interface IAccessRow {
  siteUrl: string;
  siteTitle: string;
  object: ObjectKind;
  title: string;
  url: string;
  level: string;
  how: string;
  /** The library has its own permissions (different from the site's). */
  unique: boolean;
}

/** Claims a group is known by inside SharePoint (members, and for Microsoft 365 groups also owners). */
export function groupClaims(p: IScanPrincipal): string[] {
  const id = p.key.trim().toLowerCase();
  if (p.kind === 'security') {
    return [`c:0t.c|tenant|${id}`];
  }
  return [`c:0o.c|federateddirectoryclaimprovider|${id}`, `c:0o.c|federateddirectoryclaimprovider|${id}_o`];
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
  RoleDefinitionBindings?: unknown;
}

const NO_ACCESS = 'No access';

/** Permission level names on an assignment, Limited Access left out. */
function roleNames(a: IRawAssignment): string[] {
  return rowsOf<{ Name?: string; RoleTypeKind?: number }>(a.RoleDefinitionBindings)
    .filter((r) => r.RoleTypeKind !== 1 && !/^limited access/i.test(String(r.Name || '')))
    .map((r) => String(r.Name));
}

/**
 * What a group has on an object, from its role assignments: direct grants to the group's
 * principals, and grants to SharePoint groups the group is a member of.
 */
export function matchAssignments(assignments: IRawAssignment[], principalIds: number[], spGroups: Array<{ id: number; title: string }>): { level: string; how: string } | undefined {
  const levels: string[] = [];
  const how: string[] = [];
  assignments.forEach((a) => {
    const pid = Number(a.PrincipalId);
    const names = roleNames(a);
    if (!names.length) {
      return;
    }
    if (principalIds.indexOf(pid) >= 0) {
      levels.push(...names);
      how.push('Direct');
      return;
    }
    const g = spGroups.filter((x) => x.id === pid)[0];
    if (g) {
      levels.push(...names);
      how.push(`SharePoint group: ${g.title}`);
    }
  });
  if (!levels.length) {
    return undefined;
  }
  const unique = (xs: string[]): string[] => xs.filter((x, i) => xs.indexOf(x) === i);
  return { level: unique(levels).join(', '), how: unique(how).join('; ') };
}

const LIST_SELECT = '$select=Id,Title,BaseType,HasUniqueRoleAssignments,IsSystemList,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=Hidden eq false&$top=5000';

export class AccessScanner {
  private origin: string;

  constructor(private client: SPClient, homeWebUrl: string) {
    this.origin = (/^(https:\/\/[^/]+)/i.exec(homeWebUrl) || [''])[1];
  }

  private libraries(lists: IRawList[]): IRawList[] {
    return lists.filter((l) => l.BaseType === 1 && !l.IsSystemList);
  }

  private libUrl(l: IRawList): string {
    const p = l.RootFolder && l.RootFolder.ServerRelativeUrl ? l.RootFolder.ServerRelativeUrl : '';
    return p ? this.origin + p : '';
  }

  /** Everything one principal can reach on one site (and its libraries, if asked). */
  public async scanSite(site: { url: string; title: string }, p: IScanPrincipal, includeLibraries: boolean): Promise<IAccessRow[]> {
    return p.kind === 'user' ? this.scanUser(site, p, includeLibraries) : this.scanGroup(site, p, includeLibraries);
  }

  private async scanUser(site: { url: string; title: string }, p: IScanPrincipal, includeLibraries: boolean): Promise<IAccessRow[]> {
    const web = trimSlash(site.url);
    const login = toLogin(p.key);
    const at = `@v='${encodeURIComponent(odataString(login))}'`;
    const level = (bp: { Low?: unknown; High?: unknown } | undefined): string => (bp ? permissionLevel(Number(bp.Low) || 0, Number(bp.High) || 0) : NO_ACCESS);
    const eff = (j: unknown): { Low?: unknown; High?: unknown } => {
      const o = j as { [k: string]: unknown };
      const d = (o && o.d ? o.d : o) as { [k: string]: unknown };
      return ((d && (d.GetUserEffectivePermissions || d.value)) || d) as { Low?: unknown; High?: unknown };
    };
    const rows: IAccessRow[] = [];
    const siteLevel = level(eff(await this.client.get(`${web}/_api/web/getusereffectivepermissions(${at})`, { cacheTtlMs: 0 })));
    if (siteLevel !== NO_ACCESS) {
      rows.push({ siteUrl: web, siteTitle: site.title, object: 'Site', title: site.title, url: web, level: siteLevel, how: 'Effective permission', unique: false });
    }
    if (!includeLibraries) {
      return rows;
    }
    const libs = this.libraries(rowsOf<IRawList>(await this.client.get(`${web}/_api/web/lists?${LIST_SELECT}`, { cacheTtlMs: 0 })));
    for (const l of libs) {
      if (!l.HasUniqueRoleAssignments) {
        if (siteLevel !== NO_ACCESS) {
          rows.push({ siteUrl: web, siteTitle: site.title, object: 'Library', title: String(l.Title), url: this.libUrl(l), level: siteLevel, how: 'Inherits from site', unique: false });
        }
        continue;
      }
      const lv = level(eff(await this.client.get(`${web}/_api/web/lists(guid'${l.Id}')/getusereffectivepermissions(${at})`, { cacheTtlMs: 0 })));
      if (lv !== NO_ACCESS) {
        rows.push({ siteUrl: web, siteTitle: site.title, object: 'Library', title: String(l.Title), url: this.libUrl(l), level: lv, how: 'Effective permission (own permissions)', unique: true });
      }
    }
    return rows;
  }

  private async scanGroup(site: { url: string; title: string }, p: IScanPrincipal, includeLibraries: boolean): Promise<IAccessRow[]> {
    const web = trimSlash(site.url);
    const claims = groupClaims(p);
    const filter = claims.map((c) => `LoginName eq '${odataString(c)}'`).join(' or ');
    // 1. Is the group known to this site at all? Most sites end here: one request.
    const found = rowsOf<{ Id?: number; LoginName?: string }>(
      await this.client.get(`${web}/_api/web/siteusers?$filter=${encodeURIComponent(filter)}&$select=Id,LoginName`, { cacheTtlMs: 0 })
    );
    if (!found.length) {
      return [];
    }
    const ids = found.map((f) => Number(f.Id));
    // 2. SharePoint groups the group belongs to (one request per principal found, usually one).
    const spGroups: Array<{ id: number; title: string }> = [];
    for (const id of ids) {
      rowsOf<{ Id?: number; Title?: string }>(await this.client.get(`${web}/_api/web/siteusers/getbyid(${id})/groups?$select=Id,Title`, { cacheTtlMs: 0 })).forEach((g) => {
        if (!spGroups.some((x) => x.id === Number(g.Id))) {
          spGroups.push({ id: Number(g.Id), title: String(g.Title) });
        }
      });
    }
    const ra = (path: string): Promise<IRawAssignment[]> =>
      this.client
        .get(`${web}/_api/${path}/roleassignments?$select=PrincipalId,RoleDefinitionBindings/Name,RoleDefinitionBindings/RoleTypeKind&$expand=RoleDefinitionBindings&$top=5000`, { cacheTtlMs: 0 })
        .then((j) => rowsOf<IRawAssignment>(j));
    const rows: IAccessRow[] = [];
    const siteMatch = matchAssignments(await ra('web'), ids, spGroups);
    if (siteMatch) {
      rows.push({ siteUrl: web, siteTitle: site.title, object: 'Site', title: site.title, url: web, level: siteMatch.level, how: siteMatch.how, unique: false });
    }
    if (!includeLibraries) {
      return rows;
    }
    const libs = this.libraries(rowsOf<IRawList>(await this.client.get(`${web}/_api/web/lists?${LIST_SELECT}`, { cacheTtlMs: 0 })));
    for (const l of libs) {
      if (!l.HasUniqueRoleAssignments) {
        if (siteMatch) {
          rows.push({ siteUrl: web, siteTitle: site.title, object: 'Library', title: String(l.Title), url: this.libUrl(l), level: siteMatch.level, how: 'Inherits from site', unique: false });
        }
        continue;
      }
      const m = matchAssignments(await ra(`web/lists(guid'${l.Id}')`), ids, spGroups);
      if (m) {
        rows.push({ siteUrl: web, siteTitle: site.title, object: 'Library', title: String(l.Title), url: this.libUrl(l), level: m.level, how: m.how, unique: true });
      }
    }
    return rows;
  }
}

/** Human text for a failed site. */
export function scanError(e: unknown): string {
  if (e instanceof HttpError) {
    if (e.status === 401 || e.status === 403) {
      return 'You cannot open this site (not a site collection administrator there)';
    }
    if (e.status === 404) {
      return 'Site not found';
    }
    return `${e.status}: ${e.message.substring(0, 120)}`;
  }
  return ((e as Error).message || String(e)).substring(0, 120);
}

// ---- resumable job ----------------------------------------------------------------------

export type ScanSpeed = 'gentle' | 'normal';

export interface IScanJob {
  v: 1;
  id: string;
  principal: IScanPrincipal;
  includeLibraries: boolean;
  speed: ScanSpeed;
  /** Index of the next site to start. */
  next: number;
  /** Sites started but not finished when the job was saved: done again on resume. */
  pending: number[];
  rows: IAccessRow[];
  failed: Array<{ i: number; error: string }>;
  started: number;
  saved: number;
  finished: boolean;
  total: number;
}

const JOB_KEY = 'spAdminCenter.accessScan.job.v1';
const SITES_KEY = 'spAdminCenter.accessScan.sites.v1';

/** The saved job and its site list (kept apart: the list is written once, the job often). */
export function loadJob(): { job: IScanJob; sites: Array<{ url: string; title: string }> } | undefined {
  try {
    const job = JSON.parse(window.localStorage.getItem(JOB_KEY) || 'null') as IScanJob | null;
    const sites = JSON.parse(window.localStorage.getItem(SITES_KEY) || 'null') as Array<[string, string]> | null;
    if (!job || job.v !== 1 || !sites || sites.length !== job.total) {
      return undefined;
    }
    return { job, sites: sites.map(([url, title]) => ({ url, title })) };
  } catch {
    return undefined;
  }
}

/** false when the browser refused (private mode, storage full): the scan continues, but cannot resume after a reload. */
export function saveJob(job: IScanJob, sites?: Array<{ url: string; title: string }>): boolean {
  try {
    if (sites) {
      window.localStorage.setItem(SITES_KEY, JSON.stringify(sites.map((s) => [s.url, s.title])));
    }
    job.saved = Date.now();
    window.localStorage.setItem(JOB_KEY, JSON.stringify(job));
    return true;
  } catch {
    return false;
  }
}

export function clearJob(): void {
  try {
    window.localStorage.removeItem(JOB_KEY);
    window.localStorage.removeItem(SITES_KEY);
  } catch {
    /* nothing saved */
  }
}

/** The queue of site indexes still to do: interrupted ones first, then the rest. */
export function remaining(job: IScanJob): number[] {
  const out = job.pending.slice();
  for (let i = job.next; i < job.total; i++) {
    out.push(i);
  }
  return out;
}

export const ACCESS_HEADERS = ['Who', 'Site URL', 'Site title', 'Object', 'Name', 'URL', 'Access', 'How', 'Own permissions'];

export function accessRows(principalLabel: string, rows: IAccessRow[]): unknown[][] {
  return rows.map((r) => [principalLabel, r.siteUrl, r.siteTitle, r.object, r.title, r.url, r.level, r.how, r.object === 'Library' ? (r.unique ? 'Yes' : 'No') : '']);
}
