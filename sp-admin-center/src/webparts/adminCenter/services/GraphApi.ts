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
    features: ['Sync the tenant dashboard live from Microsoft 365 (no CSV or flow needed)', 'Tenant storage trend (180 days)', 'Total vs active sites trend', 'Inactive groups & Teams', 'Teams usage per team', 'Guest counts in Group review'],
    role: 'SharePoint Administrator (or Reports Reader / Global Reader).'
  },
  {
    scope: 'Group.ReadWrite.All',
    why: 'Microsoft 365 groups and teams: owners, members, review, delete and restore.',
    features: ['See and change owners and members of the site’s Microsoft 365 group', 'Group review: no owner, one owner, guest owners, public groups with guests', 'Delete inactive groups; restore deleted groups', 'Offboarding: hand over and remove', 'Copy / compare members', 'Remove guests from all groups', 'Orphaned sites (group owners)'],
    role: 'SharePoint Administrator, Groups Administrator, or owner of the group.'
  },
  {
    scope: 'User.Read.All',
    why: 'Look up people in the directory and list guest accounts.',
    features: ['Guest users review (tenant-wide)', 'Find people when adding group owners / members', 'Offboarding: find the person leaving', 'Orphaned sites: are the owners\u2019 accounts still active?'],
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
    'group.read.all': ['group.readwrite.all', 'directory.read.all', 'directory.readwrite.all'],
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
    // No @odata.type (some $expand responses omit it): a user, as toGroup's owner count assumes.
    type: !t || t.indexOf('user') >= 0 ? 'user' : t.indexOf('group') >= 0 ? 'group' : t.indexOf('serviceprincipal') >= 0 ? 'servicePrincipal' : 'other'
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
const BATCH_SIZE = 20;
const MEMBERSHIP_TTL_MS = 10 * 60 * 1000;

export interface IGroupMembership {
  owners: IDirectoryObject[];
  members: IDirectoryObject[];
  truncated: boolean;
  error?: string;
}

export interface IGroupSummary {
  id: string;
  displayName: string;
  mail: string;
  visibility: string;
  isTeam: boolean;
}

export interface IGroupRef {
  id: string;
  displayName: string;
  mail: string;
}

export interface IDeletedGroup extends IGroupRef {
  visibility: string;
  deleted?: Date;
}

export interface IUserStatus {
  exists: boolean;
  enabled: boolean;
  guest: boolean;
  id?: string;
  displayName?: string;
  /** Could not be checked (throttled or another error): treated as still active. */
  unknown?: boolean;
}

export interface IGroupActivity {
  id: string;
  name: string;
  isDeleted: boolean;
  owner: string;
  type: string;
  lastActivity?: Date;
  members: number;
  guests: number;
  emailsReceived: number;
  activeFiles: number;
  files: number;
  siteBytes: number;
}

export interface ITeamActivity {
  id: string;
  name: string;
  type: string;
  lastActivity?: Date;
  activeUsers: number;
  activeChannels: number;
  messages: number;
  meetings: number;
  guests: number;
}

function num(v: string): number {
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function day(v: string): Date | undefined {
  if (!v) {
    return undefined;
  }
  const d = new Date(v.substring(0, 10) + 'T00:00:00Z');
  return isNaN(d.getTime()) ? undefined : d;
}

/** Rows of the Microsoft 365 groups activity report (JSON or CSV field names). */
export function groupActivityRows(records: Array<{ [key: string]: string }>): IGroupActivity[] {
  return records.map((r) => ({
    id: field(r, ['groupid']).toLowerCase(),
    name: field(r, ['groupdisplayname']),
    isDeleted: /^true$/i.test(field(r, ['isdeleted'])),
    owner: field(r, ['ownerprincipalname']),
    type: field(r, ['grouptype']),
    lastActivity: day(field(r, ['lastactivitydate'])),
    members: num(field(r, ['membercount'])),
    guests: num(field(r, ['externalmembercount', 'guestcount'])),
    emailsReceived: num(field(r, ['exchangereceivedemailcount'])),
    activeFiles: num(field(r, ['sharepointactivefilecount'])),
    files: num(field(r, ['sharepointtotalfilecount'])),
    siteBytes: num(field(r, ['sharepointsitestorageusedbyte', 'sharepointsitestorageusedinbytes']))
  }));
}

/** Rows of the Teams team activity report (JSON or CSV field names). */
export function teamActivityRows(records: Array<{ [key: string]: string }>): ITeamActivity[] {
  return records.map((r) => ({
    id: field(r, ['teamid']).toLowerCase(),
    name: field(r, ['teamname']),
    type: field(r, ['teamtype']),
    lastActivity: day(field(r, ['lastactivitydate'])),
    activeUsers: num(field(r, ['activeusers'])),
    activeChannels: num(field(r, ['activechannels'])),
    messages: num(field(r, ['channelmessages'])) || num(field(r, ['postmessages'])) + num(field(r, ['replymessages'])),
    meetings: num(field(r, ['meetingsorganized'])),
    guests: num(field(r, ['guests']))
  }));
}

export interface IBatchPart {
  status: number;
  body: unknown;
}

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
      `/v1.0/groups?$filter=${filter}&$select=id,displayName,mail,visibility,createdDateTime,description,resourceProvisioningOptions&$expand=owners($select=id,displayName,userPrincipalName,mail,userType)&$top=100`,
      max,
      onProgress
    );
    return { groups: res.items.map(toGroup), truncated: res.truncated };
  }

  // ---- membership of many groups (sites & groups report)

  private membershipCache: { [id: string]: { at: number; value: IGroupMembership } } = {};

  /** Owners and members of one group (2 requests, cached 10 minutes). */
  public async membership(groupId: string, fresh = false): Promise<IGroupMembership> {
    const id = assertGuid(groupId).toLowerCase();
    const hit = this.membershipCache[id];
    if (!fresh && hit && Date.now() - hit.at < MEMBERSHIP_TTL_MS) {
      return hit.value;
    }
    const [owners, members] = await Promise.all([this.getGroupPeople(id, 'owners'), this.getGroupPeople(id, 'members')]);
    const value: IGroupMembership = { owners: owners.people, members: members.people, truncated: owners.truncated || members.truncated };
    this.membershipCache[id] = { at: Date.now(), value };
    return value;
  }

  /** Name, e-mail, privacy and Teams flag of every Microsoft 365 group (999 per request). */
  public async groupDirectory(onProgress?: (n: number) => void): Promise<{ [id: string]: IGroupSummary }> {
    const filter = encodeURIComponent("groupTypes/any(c:c eq 'Unified')");
    const res = await this.client.getAll<IRawGroup>(`/v1.0/groups?$filter=${filter}&$select=id,displayName,mail,visibility,resourceProvisioningOptions&$top=999`, 200000, onProgress);
    const out: { [id: string]: IGroupSummary } = {};
    res.items.forEach((g) => {
      out[String(g.id).toLowerCase()] = { id: g.id, displayName: g.displayName || g.mail || g.id, mail: g.mail || '', visibility: g.visibility || '', isTeam: (g.resourceProvisioningOptions || []).indexOf('Team') >= 0 };
    });
    return out;
  }

  /**
   * Graph JSON batching: up to 20 GETs per request. Throttled (429/503/504) parts are
   * retried after the longest Retry-After they asked for, at most 5 rounds.
   */
  public async batchGet(paths: string[]): Promise<IBatchPart[]> {
    const results: IBatchPart[] = paths.map(() => ({ status: 0, body: undefined }));
    let pending = paths.map((_, i) => i);
    for (let round = 0; pending.length && round < 5; round++) {
      const waitSecs: number[] = [];
      for (let start = 0; start < pending.length; start += BATCH_SIZE) {
        const chunk = pending.slice(start, start + BATCH_SIZE);
        const reply = await this.client.send<{ responses?: Array<{ id: string; status: number; body?: unknown; headers?: { [k: string]: string } }> }>('POST', '/v1.0/$batch', {
          requests: chunk.map((i) => ({ id: String(i), method: 'GET', url: paths[i] }))
        }, true);
        ((reply && reply.responses) || []).forEach((r) => {
          const i = Number(r.id);
          results[i] = { status: Number(r.status), body: r.body };
          if (r.status === 429 || r.status === 503 || r.status === 504) {
            const h = r.headers || {};
            waitSecs.push(Number(h['Retry-After'] || h['retry-after']) || 5);
          }
        });
      }
      pending = pending.filter((i) => results[i].status === 0 || results[i].status === 429 || results[i].status === 503 || results[i].status === 504);
      if (pending.length) {
        await new Promise((r) => setTimeout(r, Math.min(Math.max.apply(null, waitSecs.concat(2)), 60) * 1000));
      }
    }
    return results;
  }

  /**
   * Owners and members of many groups, 10 groups (20 lookups) per batch request, one batch
   * at a time. Groups with more than 999 owners or members get their extra pages read
   * separately. `shouldStop` is checked between batches.
   */
  public async membershipMany(groupIds: string[], onProgress?: (done: number, total: number) => void, shouldStop?: () => boolean): Promise<{ [id: string]: IGroupMembership }> {
    const ids = groupIds.map((g) => assertGuid(g).toLowerCase()).filter((g, i, a) => a.indexOf(g) === i);
    const out: { [id: string]: IGroupMembership } = {};
    const todo = ids.filter((id) => {
      const hit = this.membershipCache[id];
      if (hit && Date.now() - hit.at < MEMBERSHIP_TTL_MS) {
        out[id] = hit.value;
        return false;
      }
      return true;
    });
    let done = ids.length - todo.length;
    if (onProgress) {
      onProgress(done, ids.length);
    }
    const perBatch = BATCH_SIZE / 2;
    for (let start = 0; start < todo.length; start += perBatch) {
      if (shouldStop && shouldStop()) {
        break;
      }
      const chunk = todo.slice(start, start + perBatch);
      const paths: string[] = [];
      chunk.forEach((id) => {
        paths.push(`/groups/${id}/owners?${MEMBER_SELECT}&$top=999`);
        paths.push(`/groups/${id}/members?${MEMBER_SELECT}&$top=999`);
      });
      const parts = await this.batchGet(paths);
      for (let k = 0; k < chunk.length; k++) {
        const id = chunk[k];
        const value: IGroupMembership = { owners: [], members: [], truncated: false };
        for (const [kind, part] of [['owners', parts[2 * k]], ['members', parts[2 * k + 1]]] as Array<['owners' | 'members', IBatchPart]>) {
          if (part.status >= 200 && part.status < 300) {
            const body = (part.body || {}) as { value?: IRawDirObject[]; '@odata.nextLink'?: string };
            let people = (body.value || []).map(dirObject);
            if (body['@odata.nextLink']) {
              const more = await this.client.getAll<IRawDirObject>(body['@odata.nextLink'], 50000);
              people = people.concat(more.items.map(dirObject));
              value.truncated = value.truncated || more.truncated;
            }
            value[kind] = people;
          } else {
            const err = (part.body as { error?: { message?: string } } | undefined) || {};
            value.error = part.status === 404 ? 'Group not found (deleted?)' : part.status === 403 ? 'Access denied to this group' : `${part.status || 'No answer'}: ${(err.error && err.error.message) || 'could not read'}`;
          }
        }
        out[id] = value;
        if (!value.error) {
          this.membershipCache[id] = { at: Date.now(), value };
        }
      }
      done += chunk.length;
      if (onProgress) {
        onProgress(done, ids.length);
      }
    }
    return out;
  }

  // ---- usage reports for groups and Teams

  /** Last activity, members, guests and storage for every Microsoft 365 group (one report). */
  public async groupActivity(period: 'D7' | 'D30' | 'D90' | 'D180' = 'D180', onProgress?: (n: number) => void): Promise<IGroupActivity[]> {
    return groupActivityRows(await this.report('getOffice365GroupsActivityDetail', period, onProgress));
  }

  /** Active users, messages and meetings for every team (one report). */
  public async teamsActivity(period: 'D7' | 'D30' | 'D90' | 'D180' = 'D30', onProgress?: (n: number) => void): Promise<ITeamActivity[]> {
    return teamActivityRows(await this.report('getTeamsTeamActivityDetail', period, onProgress));
  }

  // ---- group lifecycle

  /** Deletes a Microsoft 365 group (it stays restorable for 30 days, with its site and team). */
  public async deleteGroup(groupId: string): Promise<void> {
    await this.client.send('DELETE', `/v1.0/groups/${assertGuid(groupId)}`);
  }

  /** Microsoft 365 groups deleted in the last 30 days. */
  public async deletedGroups(): Promise<IDeletedGroup[]> {
    const res = await this.client.getAll<IRawGroup & { deletedDateTime?: string; groupTypes?: string[] }>(
      '/v1.0/directory/deletedItems/microsoft.graph.group?$select=id,displayName,mail,visibility,description,deletedDateTime,groupTypes&$top=999',
      20000
    );
    return res.items
      .filter((g) => (g.groupTypes || ['Unified']).indexOf('Unified') >= 0)
      .map((g) => ({ id: g.id, displayName: g.displayName || g.mail || g.id, mail: g.mail || '', visibility: g.visibility || '', deleted: g.deletedDateTime ? new Date(g.deletedDateTime) : undefined }));
  }

  public async restoreGroup(groupId: string): Promise<void> {
    await this.client.send('POST', `/v1.0/directory/deletedItems/${assertGuid(groupId)}/restore`);
  }

  // ---- a person's groups (offboarding, guest clean-up)

  /** Microsoft 365 groups a user belongs to (direct membership). */
  public async userGroups(userId: string, kind: 'memberOf' | 'ownedObjects'): Promise<IGroupRef[]> {
    const res = await this.client.getAll<IRawGroup & { groupTypes?: string[] }>(
      `/v1.0/users/${assertGuid(userId)}/${kind}/microsoft.graph.group?$select=id,displayName,mail,groupTypes,visibility&$top=999`,
      5000,
      undefined,
      { cacheTtlMs: 0 }
    );
    return res.items
      .filter((g) => (g.groupTypes || []).indexOf('Unified') >= 0)
      .map((g) => ({ id: g.id.toLowerCase(), displayName: g.displayName || g.mail || g.id, mail: g.mail || '' }));
  }

  /** Owners of many groups: 20 groups per batch request. */
  public async ownersMany(groupIds: string[], onProgress?: (done: number, total: number) => void, shouldStop?: () => boolean): Promise<{ [id: string]: IDirectoryObject[] | undefined }> {
    const ids = groupIds.map((g) => assertGuid(g).toLowerCase()).filter((g, i, a) => a.indexOf(g) === i);
    const out: { [id: string]: IDirectoryObject[] | undefined } = {};
    for (let start = 0; start < ids.length; start += BATCH_SIZE) {
      if (shouldStop && shouldStop()) {
        break;
      }
      const chunk = ids.slice(start, start + BATCH_SIZE);
      const parts = await this.batchGet(chunk.map((id) => `/groups/${id}/owners?${MEMBER_SELECT}&$top=999`));
      chunk.forEach((id, k) => {
        const part = parts[k];
        out[id] = part.status >= 200 && part.status < 300 ? (((part.body || {}) as { value?: IRawDirObject[] }).value || []).map(dirObject) : undefined;
      });
      if (onProgress) {
        onProgress(Math.min(start + chunk.length, ids.length), ids.length);
      }
    }
    return out;
  }

  /**
   * Whether people still have a working account: looked up by object id, UPN or e-mail,
   * 20 per batch request. Missing (deleted) accounts come back as { exists: false }.
   */
  public async userStatus(keys: string[], onProgress?: (done: number, total: number) => void, shouldStop?: () => boolean): Promise<{ [key: string]: IUserStatus }> {
    const unique = keys.map((k) => k.trim().toLowerCase()).filter((k, i, a) => !!k && a.indexOf(k) === i);
    const out: { [key: string]: IUserStatus } = {};
    for (let start = 0; start < unique.length; start += BATCH_SIZE) {
      if (shouldStop && shouldStop()) {
        break;
      }
      const chunk = unique.slice(start, start + BATCH_SIZE);
      const parts = await this.batchGet(chunk.map((k) => `/users/${encodeURIComponent(k)}?$select=id,displayName,mail,userPrincipalName,accountEnabled,userType`));
      chunk.forEach((k, i) => {
        const part = parts[i];
        if (part.status >= 200 && part.status < 300) {
          const u = (part.body || {}) as { id?: string; displayName?: string; accountEnabled?: boolean; userType?: string };
          out[k] = { exists: true, enabled: u.accountEnabled !== false, guest: (u.userType || '').toLowerCase() === 'guest', id: u.id, displayName: u.displayName };
        } else if (part.status === 404) {
          out[k] = { exists: false, enabled: false, guest: false };
        } else {
          out[k] = { exists: true, enabled: true, guest: false, unknown: true };
        }
      });
      if (onProgress) {
        onProgress(Math.min(start + chunk.length, unique.length), unique.length);
      }
    }
    return out;
  }

  /** Microsoft 365 groups whose name starts with the text (for pickers). */
  public async searchGroups(text: string): Promise<IGroupRef[]> {
    const t = text.trim().replace(/'/g, "''");
    if (!t) {
      return [];
    }
    const filter = encodeURIComponent(`groupTypes/any(c:c eq 'Unified') and startswith(displayName,'${t}')`);
    const res = await this.client.get<{ value?: Array<IRawGroup> }>(`/v1.0/groups?$filter=${filter}&$select=id,displayName,mail&$top=25&$count=true&$orderby=displayName`, {
      headers: { ConsistencyLevel: 'eventual' }
    });
    return (res.value || []).map((g) => ({ id: g.id.toLowerCase(), displayName: g.displayName || g.mail || g.id, mail: g.mail || '' }));
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
