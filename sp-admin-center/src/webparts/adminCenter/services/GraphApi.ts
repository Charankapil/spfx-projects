import { GraphClient } from '../core/GraphClient';
import { parseCsv, stripBom } from './csv';
import { IMapping } from './StorageImport';

/**
 * The Microsoft 365 features that need Microsoft Graph. Each one needs a single
 * delegated permission, so an approver can grant only some of them and the rest
 * of the app keeps working (features say exactly which permission is missing).
 */

export interface IGraphPermission {
  scope: string;
  why: string;
  features: string[];
  /** Directory role the signed-in admin also needs (delegated calls are limited by the user's own rights). */
  role: string;
}

export const GRAPH_PERMISSIONS: IGraphPermission[] = [
  {
    scope: 'Reports.Read.All',
    why: 'Usage reports: every site collection with storage used and allocated, last activity, files and page views.',
    features: ['Sync the tenant dashboard live from Microsoft 365 (no CSV or flow needed)', 'Tenant storage trend (180 days)', 'Total vs active sites trend'],
    role: 'SharePoint Administrator (or Reports Reader / Global Reader).'
  },
  {
    scope: 'Group.ReadWrite.All',
    why: 'Microsoft 365 groups behind team sites: owners, members and ownerless groups.',
    features: ['See and change owners and members of the site’s Microsoft 365 group', 'Find ownerless and single-owner groups tenant-wide', 'Add an owner to an ownerless group'],
    role: 'SharePoint Administrator, Groups Administrator, or owner of the group.'
  },
  {
    scope: 'User.Read.All',
    why: 'Look up people in the directory and list guest accounts.',
    features: ['Guest users review (tenant-wide)', 'Find people when adding group owners / members'],
    role: 'Any member user (directory read).'
  },
  {
    scope: 'AuditLog.Read.All',
    why: 'Last sign-in dates, so stale guests can be found. Optional.',
    features: ['Last sign-in date for guests (stale guest detection)'],
    role: 'Reports Reader, Security Reader or Global Reader, plus Microsoft Entra ID P1.'
  }
];

export function hasScope(granted: string[], scope: string): boolean {
  const s = scope.toLowerCase();
  const all = granted.map((g) => g.toLowerCase());
  if (all.indexOf(s) >= 0) {
    return true;
  }
  // Broader scopes that include the requested one.
  const broader: { [k: string]: string[] } = {
    'user.read.all': ['user.readwrite.all', 'directory.read.all', 'directory.readwrite.all'],
    'group.readwrite.all': ['directory.readwrite.all'],
    'auditlog.read.all': []
  };
  return (broader[s] || []).some((b) => all.indexOf(b) >= 0);
}

// ---- usage reports ------------------------------------------------------------

const norm = (s: string): string => s.toLowerCase().replace(/[^a-z0-9]/g, '');

/** Report rows from either the JSON format (camelCase keys) or the CSV format (titled headers), keyed by normalized name. */
export function reportRecords(body: string): Array<{ [key: string]: string }> {
  const text = stripBom(body || '').trim();
  if (!text) {
    return [];
  }
  if (text.charAt(0) === '{' || text.charAt(0) === '[') {
    const parsed = JSON.parse(text) as { value?: Array<{ [k: string]: unknown }> } | Array<{ [k: string]: unknown }>;
    const items = Array.isArray(parsed) ? parsed : parsed.value || [];
    return items.map((item) => {
      const rec: { [key: string]: string } = {};
      Object.keys(item).forEach((k) => {
        const v = item[k];
        rec[norm(k)] = v === null || v === undefined ? '' : Array.isArray(v) ? v.join(',') : String(v);
      });
      return rec;
    });
  }
  const rows = parseCsv(text);
  const headers = (rows[0] || []).map(norm);
  return rows.slice(1).map((r) => {
    const rec: { [key: string]: string } = {};
    headers.forEach((h, i) => {
      rec[h] = r[i] === undefined ? '' : r[i];
    });
    return rec;
  });
}

/** First non-empty value among several normalized names (JSON and CSV spell some fields differently). */
export function field(rec: { [key: string]: string }, names: string[]): string {
  for (const n of names) {
    const v = rec[n];
    if (v !== undefined && v !== '') {
      return v;
    }
  }
  return '';
}

export const USAGE_HEADERS = ['Site URL', 'Storage Used (Byte)', 'Storage Allocated (Byte)', 'Is Deleted', 'Last Activity Date', 'Root Web Template', 'Owner Display Name', 'File Count', 'Page View Count', 'Site Id'];

/** Column mapping for the table built by usageTable: fixed, so it never disturbs the CSV import mapping. */
export const USAGE_MAPPING: IMapping = {
  url: 'Site URL',
  storage: 'Storage Used (Byte)',
  storageUnit: 'bytes',
  quota: 'Storage Allocated (Byte)',
  quotaUnit: 'bytes',
  deleted: 'Is Deleted',
  lastActivity: 'Last Activity Date',
  template: 'Root Web Template',
  includeDeleted: true,
  excludeOneDrive: true
};

export interface IUsageTable {
  /** CSV-like table (header row first) ready for the snapshot importer. */
  rows: string[][];
  sites: number;
  /** Rows whose site URL is hidden because the tenant conceals user, group and site names in reports. */
  concealed: number;
  refreshDate?: string;
}

export function usageTable(records: Array<{ [key: string]: string }>): IUsageTable {
  const rows: string[][] = [USAGE_HEADERS.slice()];
  let concealed = 0;
  let refreshDate: string | undefined;
  records.forEach((r) => {
    const url = field(r, ['siteurl']);
    refreshDate = refreshDate || field(r, ['reportrefreshdate']) || undefined;
    if (!/^https:\/\//i.test(url)) {
      concealed++;
      return;
    }
    rows.push([
      url,
      field(r, ['storageusedinbytes', 'storageusedbyte']),
      field(r, ['storageallocatedinbytes', 'storageallocatedbyte']),
      field(r, ['isdeleted']).toLowerCase(),
      field(r, ['lastactivitydate']),
      field(r, ['rootwebtemplate']),
      field(r, ['ownerdisplayname']),
      field(r, ['filecount']),
      field(r, ['pageviewcount']),
      field(r, ['siteid'])
    ]);
  });
  return { rows, sites: rows.length - 1, concealed, refreshDate };
}

export interface ITrendPoint {
  date: string;
  value: number;
  /** Second series where the report has one (active sites). */
  value2?: number;
}

export function trendPoints(records: Array<{ [key: string]: string }>, valueNames: string[], value2Names?: string[]): ITrendPoint[] {
  const byDate: { [d: string]: ITrendPoint } = {};
  records.forEach((r) => {
    const type = field(r, ['sitetype']);
    if (type && type.toLowerCase() !== 'all') {
      return;
    }
    const date = field(r, ['reportdate']);
    const value = Number(field(r, valueNames));
    if (!date || !isFinite(value)) {
      return;
    }
    const p: ITrendPoint = { date: date.substring(0, 10), value };
    if (value2Names) {
      const v2 = Number(field(r, value2Names));
      if (isFinite(v2)) {
        p.value2 = v2;
      }
    }
    byDate[p.date] = p;
  });
  return Object.keys(byDate).sort().map((d) => byDate[d]);
}

/** Snapshot time for a report: noon UTC on its refresh date (reports describe the previous day), else now. */
export function reportTime(refreshDate: string | undefined, nowSec: number): number {
  const t = refreshDate ? Date.parse(refreshDate.substring(0, 10) + 'T12:00:00Z') : NaN;
  return isFinite(t) && t / 1000 <= nowSec ? Math.floor(t / 1000) : nowSec;
}

// ---- groups and people ----------------------------------------------------------

export interface IDirectoryObject {
  id: string;
  displayName: string;
  mail?: string;
  userPrincipalName?: string;
  userType?: string;
  type: 'user' | 'group' | 'servicePrincipal' | 'other';
}

export interface IM365Group {
  id: string;
  displayName: string;
  mail?: string;
  visibility?: string;
  createdDateTime?: string;
  description?: string;
  isTeam: boolean;
  ownerCount?: number;
  owners?: IDirectoryObject[];
}

export interface IGuest {
  id: string;
  displayName: string;
  mail: string;
  userPrincipalName: string;
  created?: Date;
  state: string;
  enabled: boolean;
  lastSignIn?: Date;
}

interface IRawDirObject {
  '@odata.type'?: string;
  id: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  userType?: string;
}

interface IRawGroup {
  id: string;
  displayName?: string;
  mail?: string;
  visibility?: string;
  createdDateTime?: string;
  description?: string;
  resourceProvisioningOptions?: string[];
  owners?: IRawDirObject[];
}

interface IRawUser {
  id: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  createdDateTime?: string;
  externalUserState?: string;
  accountEnabled?: boolean;
  signInActivity?: { lastSignInDateTime?: string; lastNonInteractiveSignInDateTime?: string; lastSuccessfulSignInDateTime?: string };
}

function dirObject(o: IRawDirObject): IDirectoryObject {
  const t = (o['@odata.type'] || '').toLowerCase();
  return {
    id: o.id,
    displayName: o.displayName || o.userPrincipalName || o.mail || o.id,
    mail: o.mail,
    userPrincipalName: o.userPrincipalName,
    userType: o.userType,
    type: t.indexOf('user') >= 0 ? 'user' : t.indexOf('group') >= 0 ? 'group' : t.indexOf('serviceprincipal') >= 0 ? 'servicePrincipal' : 'other'
  };
}

export function toGroup(g: IRawGroup): IM365Group {
  return {
    id: g.id,
    displayName: g.displayName || g.mail || g.id,
    mail: g.mail,
    visibility: g.visibility,
    createdDateTime: g.createdDateTime,
    description: g.description,
    isTeam: (g.resourceProvisioningOptions || []).indexOf('Team') >= 0,
    owners: g.owners ? g.owners.map(dirObject) : undefined,
    ownerCount: g.owners ? g.owners.filter((o) => (o['@odata.type'] || '#microsoft.graph.user').toLowerCase().indexOf('user') >= 0).length : undefined
  };
}

function latestDate(values: Array<string | undefined>): Date | undefined {
  let best: Date | undefined;
  values.forEach((v) => {
    const d = v ? new Date(v) : undefined;
    if (d && !isNaN(d.getTime()) && d.getFullYear() > 2000 && (!best || d > best)) {
      best = d;
    }
  });
  return best;
}

export function toGuest(u: IRawUser): IGuest {
  const created = u.createdDateTime ? new Date(u.createdDateTime) : undefined;
  return {
    id: u.id,
    displayName: u.displayName || u.mail || u.userPrincipalName || u.id,
    mail: u.mail || '',
    userPrincipalName: u.userPrincipalName || '',
    created: created && !isNaN(created.getTime()) ? created : undefined,
    state: u.externalUserState || '',
    enabled: u.accountEnabled !== false,
    lastSignIn: u.signInActivity ? latestDate([u.signInActivity.lastSignInDateTime, u.signInActivity.lastNonInteractiveSignInDateTime, u.signInActivity.lastSuccessfulSignInDateTime]) : undefined
  };
}

/** GUIDs only: never let a value from a list or CSV change the shape of a Graph path. */
export function assertGuid(id: string): string {
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id || '')) {
    throw new Error('Not a valid id: ' + id);
  }
  return id;
}

const MEMBER_SELECT = '$select=id,displayName,mail,userPrincipalName,userType';

export class GraphApi {
  constructor(public client: GraphClient) {}

  public grantedScopes(): Promise<string[]> {
    return this.client.grantedScopes();
  }

  // ---- reports

  /**
   * Fetches a usage report. The JSON format (beta endpoint) is tried first because it
   * stays on graph.microsoft.com; the documented v1.0 CSV download is the fallback.
   */
  public async report(name: string, period: 'D7' | 'D30' | 'D90' | 'D180', onProgress?: (n: number) => void): Promise<Array<{ [key: string]: string }>> {
    const fn = `${name}(period='${period}')`;
    try {
      const all = await this.client.getAll<{ [k: string]: unknown }>(`/beta/reports/${fn}?$format=application/json`, 200000, onProgress);
      return reportRecords(JSON.stringify(all.items));
    } catch (e) {
      const status = (e as { status?: number }).status;
      if (status === 401 || status === 403 || status === 429 || status === 503) {
        throw e;
      }
      const csv = await this.client.getText(`/v1.0/reports/${fn}`, { cacheTtlMs: 0, headers: { Accept: 'text/csv, application/octet-stream, */*' } });
      return reportRecords(csv);
    }
  }

  public async siteUsage(period: 'D7' | 'D30' | 'D90' | 'D180' = 'D7', onProgress?: (n: number) => void): Promise<IUsageTable> {
    return usageTable(await this.report('getSharePointSiteUsageDetail', period, onProgress));
  }

  public async storageTrend(): Promise<ITrendPoint[]> {
    return trendPoints(await this.report('getSharePointSiteUsageStorage', 'D180'), ['storageusedinbytes', 'storageusedbyte']);
  }

  public async siteCountTrend(): Promise<ITrendPoint[]> {
    return trendPoints(await this.report('getSharePointSiteUsageSiteCounts', 'D180'), ['total'], ['active']);
  }

  // ---- groups

  public async getGroup(groupId: string): Promise<IM365Group> {
    const g = await this.client.get<IRawGroup>(`/v1.0/groups/${assertGuid(groupId)}?$select=id,displayName,mail,visibility,createdDateTime,description,resourceProvisioningOptions`);
    return toGroup(g);
  }

  public async getGroupPeople(groupId: string, kind: 'owners' | 'members'): Promise<{ people: IDirectoryObject[]; truncated: boolean }> {
    const res = await this.client.getAll<IRawDirObject>(`/v1.0/groups/${assertGuid(groupId)}/${kind}?${MEMBER_SELECT}&$top=999`, 5000);
    return { people: res.items.map(dirObject), truncated: res.truncated };
  }

  public async addToGroup(groupId: string, kind: 'owners' | 'members', userId: string): Promise<void> {
    await this.client.send('POST', `/v1.0/groups/${assertGuid(groupId)}/${kind}/$ref`, { '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${assertGuid(userId)}` });
  }

  public async removeFromGroup(groupId: string, kind: 'owners' | 'members', userId: string): Promise<void> {
    await this.client.send('DELETE', `/v1.0/groups/${assertGuid(groupId)}/${kind}/${assertGuid(userId)}/$ref`);
  }

  /**
   * Every Microsoft 365 group with its owners, 100 per page (17,000 groups = 170 paced requests).
   * Only runs when the admin asks for the scan.
   */
  public async listGroupsWithOwners(onProgress?: (n: number) => void, max = 50000): Promise<{ groups: IM365Group[]; truncated: boolean }> {
    const filter = encodeURIComponent("groupTypes/any(c:c eq 'Unified')");
    const res = await this.client.getAll<IRawGroup>(
      `/v1.0/groups?$filter=${filter}&$select=id,displayName,mail,visibility,createdDateTime,resourceProvisioningOptions&$expand=owners($select=id,displayName,userPrincipalName)&$top=100`,
      max,
      onProgress
    );
    return { groups: res.items.map(toGroup), truncated: res.truncated };
  }

  // ---- people

  /** Resolves an e-mail / UPN to a directory user (exact match first, then mail). */
  public async findUser(emailOrUpn: string): Promise<IDirectoryObject | undefined> {
    const v = emailOrUpn.trim();
    if (!v) {
      return undefined;
    }
    try {
      const u = await this.client.get<IRawDirObject>(`/v1.0/users/${encodeURIComponent(v)}?$select=id,displayName,mail,userPrincipalName,userType`);
      return { ...dirObject(u), type: 'user' };
    } catch (e) {
      if ((e as { status?: number }).status !== 404) {
        throw e;
      }
    }
    const q = v.replace(/'/g, "''");
    const res = await this.client.get<{ value?: IRawDirObject[] }>(`/v1.0/users?$filter=${encodeURIComponent(`mail eq '${q}' or otherMails/any(m:m eq '${q}')`)}&$select=id,displayName,mail,userPrincipalName,userType&$count=true`, { headers: { ConsistencyLevel: 'eventual' } });
    const first = (res.value || [])[0];
    return first ? { ...dirObject(first), type: 'user' } : undefined;
  }

  /**
   * All guest accounts. Asks for sign-in activity first; if that is not allowed
   * (AuditLog.Read.All missing, no Entra ID P1, or no reader role) it lists without it.
   */
  public async listGuests(onProgress?: (n: number) => void, max = 100000): Promise<{ guests: IGuest[]; signInKnown: boolean; truncated: boolean }> {
    const filter = `/v1.0/users?$filter=${encodeURIComponent("userType eq 'Guest'")}`;
    const select = '$select=id,displayName,mail,userPrincipalName,createdDateTime,externalUserState,accountEnabled';
    try {
      // Sign-in activity is served in smaller pages.
      const res = await this.client.getAll<IRawUser>(`${filter}&$top=100&${select},signInActivity`, max, onProgress);
      return { guests: res.items.map(toGuest), signInKnown: true, truncated: res.truncated };
    } catch (e) {
      const status = (e as { status?: number }).status;
      if (status !== 400 && status !== 401 && status !== 403) {
        throw e;
      }
    }
    const res = await this.client.getAll<IRawUser>(`${filter}&$top=500&${select}`, max, onProgress);
    return { guests: res.items.map(toGuest), signInKnown: false, truncated: res.truncated };
  }
}

export const STALE_DAYS = 90;


/** A guest is stale when they have not signed in for 90 days (or, without sign-in data, never accepted and were invited 90+ days ago). */
export function isStaleGuest(g: IGuest, signInKnown: boolean, now = Date.now()): boolean {
  if (signInKnown) {
    const last = g.lastSignIn || g.created;
    return !!last && (now - last.getTime()) / 86400000 >= STALE_DAYS;
  }
  return g.state === 'PendingAcceptance' && !!g.created && (now - g.created.getTime()) / 86400000 >= STALE_DAYS;
}
