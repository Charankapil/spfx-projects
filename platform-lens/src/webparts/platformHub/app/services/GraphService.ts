/* eslint-disable @typescript-eslint/no-explicit-any */
import { ApiClient, isApiError } from '../http/ApiClient';
import {
  IActiveUserCounts, IAppCredential, ICopilotUser, IDirectoryRole, IGroup, IHealthIssue, IHealthOverview, IMessage,
  IMfaRegistration, ISite, ISku, IStoragePoint, ITeamActivity, ITenantSharing, IUserAccount, IUserActivity, IUserInfo
} from '../models';
import { skuFriendlyName, isPowerPlatformSku } from '../logic/skus';

export const GRAPH_RESOURCE = 'https://graph.microsoft.com';
const V1 = 'https://graph.microsoft.com/v1.0';
const BETA = 'https://graph.microsoft.com/beta';

function str(v: any): string {
  return v === undefined || v === null ? '' : String(v);
}

function num(v: any): number {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function maxDate(...dates: (string | undefined)[]): string | undefined {
  let best: string | undefined;
  dates.forEach((d) => {
    if (d && (!best || d > best)) best = d;
  });
  return best;
}

/** Titles are not part of the usage report; derive a readable one from the URL. */
export function titleFromUrl(url: string): string {
  if (!url) return '(name hidden)';
  const parts = url.replace(/\/+$/, '').split('/');
  const last = decodeURIComponent(parts[parts.length - 1] || '');
  return last && !/sharepoint\.com$/i.test(last) ? last : 'Root site';
}

export class GraphService {
  private readonly _api: ApiClient;

  constructor(api: ApiClient) {
    this._api = api;
  }

  private _all<T>(url: string, maxItems?: number, headers?: Record<string, string>): Promise<T[]> {
    return this._api.getAll<T>(GRAPH_RESOURCE, url, { maxItems, headers });
  }

  /* ---------- Reports (beta JSON format avoids the CSV redirect, which browsers block) ---------- */

  private _report(name: string, period: string, maxItems?: number): Promise<any[]> {
    return this._all<any>(`${BETA}/reports/${name}(period='${period}')?$format=application/json`, maxItems);
  }

  public async getReportsConcealed(): Promise<boolean> {
    const s = await this._api.getJson<any>(GRAPH_RESOURCE, `${V1}/admin/reportSettings`);
    return !!s.displayConcealedNames;
  }

  public async getSites(maxItems: number): Promise<ISite[]> {
    const raw = await this._report('getSharePointSiteUsageDetail', 'D30', maxItems);
    return raw.map((r) => {
      const url = str(r.siteUrl);
      return {
        id: str(r.siteId),
        url,
        title: titleFromUrl(url),
        ownerName: str(r.ownerDisplayName),
        ownerUpn: str(r.ownerPrincipalName),
        lastActivity: r.lastActivityDate || undefined,
        template: str(r.rootWebTemplate) || 'Unknown',
        storageUsedBytes: num(r.storageUsedInBytes),
        storageAllocatedBytes: num(r.storageAllocatedInBytes),
        fileCount: num(r.fileCount),
        activeFileCount: num(r.activeFileCount),
        pageViews: num(r.pageViewCount),
        externalSharing: r.externalSharing === true || str(r.externalSharing).toLowerCase() === 'true',
        anonymousLinks: num(r.anonymousLinkCount),
        isDeleted: r.isDeleted === true || str(r.isDeleted).toLowerCase() === 'true'
      };
    });
  }

  public async getStorageTrend(): Promise<IStoragePoint[]> {
    const raw = await this._report('getSharePointSiteUsageStorage', 'D90');
    const byDate: Record<string, number> = {};
    raw.forEach((r) => {
      const d = str(r.reportDate);
      if (d) byDate[d] = (byDate[d] || 0) + num(r.storageUsedInBytes);
    });
    return Object.keys(byDate).sort().map((date) => ({ date, bytes: byDate[date] }));
  }

  public async getSharingSettings(): Promise<ITenantSharing> {
    const s = await this._api.getJson<any>(GRAPH_RESOURCE, `${V1}/admin/sharepoint/settings`);
    return {
      sharingCapability: str(s.sharingCapability),
      isResharingByExternalUsersEnabled: s.isResharingByExternalUsersEnabled,
      sharingDomainRestrictionMode: s.sharingDomainRestrictionMode,
      isUnmanagedSyncAppForTenantRestricted: s.isUnmanagedSyncAppForTenantRestricted,
      oneDriveLoopSharingCapability: s.oneDriveLoopSharingCapability
    };
  }

  public async getActiveUserCounts(): Promise<IActiveUserCounts[]> {
    const raw = await this._report('getOffice365ActiveUserCounts', 'D90');
    return raw
      .map((r) => ({
        date: str(r.reportDate),
        office365: num(r.office365),
        exchange: num(r.exchange),
        oneDrive: num(r.oneDrive),
        sharePoint: num(r.sharePoint),
        teams: num(r.teams)
      }))
      .filter((r) => !!r.date)
      .sort((a, b) => (a.date < b.date ? -1 : 1));
  }

  public async getUserActivity(maxItems: number): Promise<IUserActivity[]> {
    const raw = await this._report('getOffice365ActiveUserDetail', 'D30', maxItems);
    return raw.map((r) => ({
      upn: str(r.userPrincipalName),
      displayName: str(r.displayName),
      products: Array.isArray(r.assignedProducts) ? r.assignedProducts.map(str) : str(r.assignedProducts).split('+').filter(Boolean),
      lastActivity: maxDate(
        r.exchangeLastActivityDate, r.oneDriveLastActivityDate, r.sharePointLastActivityDate,
        r.skypeForBusinessLastActivityDate, r.yammerLastActivityDate, r.teamsLastActivityDate
      ),
      isDeleted: r.isDeleted === true
    }));
  }

  public async getCopilotUsers(): Promise<ICopilotUser[]> {
    const urls = [
      `${BETA}/reports/getMicrosoft365CopilotUsageUserDetail(period='D30')?$format=application/json`,
      `${BETA}/copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='D30')?$format=application/json`
    ];
    let lastError: unknown;
    for (const url of urls) {
      try {
        const raw = await this._all<any>(url);
        return raw.map((r) => ({
          upn: str(r.userPrincipalName),
          displayName: str(r.displayName),
          lastActivity: r.lastActivityDate || undefined
        }));
      } catch (e) {
        lastError = e;
        if (!isApiError(e) || (e.kind !== 'notFound' && e.status !== 400)) break;
      }
    }
    throw lastError;
  }

  public async getTeamsActivity(): Promise<ITeamActivity[]> {
    const raw = await this._report('getTeamsTeamActivityDetail', 'D30');
    return raw
      .filter((r) => !(r.isDeleted === true))
      .map((r) => {
        const d = Array.isArray(r.details) && r.details.length ? r.details[0] : r;
        return {
          id: str(r.teamId),
          name: str(r.teamName) || '(name hidden)',
          teamType: str(r.teamType),
          lastActivity: r.lastActivityDate || undefined,
          activeUsers: num(d.activeUsers),
          guests: num(d.guests),
          activeExternalUsers: num(d.activeExternalUsers),
          activeChannels: num(d.activeChannels),
          messages: num(d.channelMessages) + num(d.postMessages) + num(d.replyMessages),
          meetings: num(d.meetingsOrganized)
        };
      });
  }

  /* ---------- Directory ---------- */

  public async getGroups(maxItems: number): Promise<IGroup[]> {
    const raw = await this._all<any>(
      `${V1}/groups?$select=id,displayName,mail,createdDateTime,visibility,groupTypes,resourceProvisioningOptions&$expand=owners($select=id)&$top=100`,
      maxItems
    );
    return raw
      .filter((g) => Array.isArray(g.groupTypes) && g.groupTypes.indexOf('Unified') >= 0)
      .map((g) => ({
        id: str(g.id),
        displayName: str(g.displayName),
        mail: g.mail || undefined,
        createdTime: g.createdDateTime || undefined,
        visibility: str(g.visibility) || 'Unknown',
        isTeam: Array.isArray(g.resourceProvisioningOptions) && g.resourceProvisioningOptions.indexOf('Team') >= 0,
        ownerCount: Array.isArray(g.owners) ? g.owners.length : 0
      }));
  }

  public async getSubscribedSkus(): Promise<ISku[]> {
    const raw = await this._all<any>(`${V1}/subscribedSkus`);
    return raw.map((s) => {
      const prepaid = s.prepaidUnits || {};
      const plans: string[] = Array.isArray(s.servicePlans) ? s.servicePlans.map((p: any) => str(p.servicePlanName)) : [];
      return {
        skuId: str(s.skuId),
        skuPartNumber: str(s.skuPartNumber),
        friendlyName: skuFriendlyName(str(s.skuPartNumber)),
        enabled: num(prepaid.enabled),
        consumed: num(s.consumedUnits),
        suspended: num(prepaid.suspended),
        warning: num(prepaid.warning),
        isPowerPlatform: isPowerPlatformSku(str(s.skuPartNumber), plans),
        appliesTo: str(s.appliesTo)
      };
    });
  }

  /**
   * Resolves user ids via $batch (20 per request). Ids that return 404 are
   * reported as Deleted, which is how orphaned resources are detected.
   */
  public async resolveUsers(ids: string[], includeSignIn: boolean): Promise<{ users: IUserInfo[]; signInUnavailable: boolean }> {
    const unique = ids.filter((id, i) => !!id && ids.indexOf(id) === i);
    const result: IUserInfo[] = [];
    let withSignIn = includeSignIn;
    const select = (): string =>
      'id,displayName,userPrincipalName,department,country,jobTitle,accountEnabled' + (withSignIn ? ',signInActivity' : '');

    for (let i = 0; i < unique.length; i += 20) {
      let chunk = unique.slice(i, i + 20);
      for (let attempt = 0; attempt < 4 && chunk.length; attempt++) {
        const body = { requests: chunk.map((id, n) => ({ id: String(n), method: 'GET', url: `/users/${id}?$select=${select()}` })) };
        const res = await this._api.postJson<any>(GRAPH_RESOURCE, `${V1}/$batch`, body);
        const responses: any[] = Array.isArray(res.responses) ? res.responses : [];
        const retry: string[] = [];
        let forbiddenSignIn = false;
        let waitMs = 0;
        responses.forEach((r) => {
          const id = chunk[Number(r.id)];
          if (r.status === 200 && r.body) {
            const u = r.body;
            result.push({
              id,
              displayName: u.displayName,
              upn: u.userPrincipalName,
              department: u.department || undefined,
              country: u.country || undefined,
              jobTitle: u.jobTitle || undefined,
              accountEnabled: u.accountEnabled,
              lastSignIn: u.signInActivity ? u.signInActivity.lastSignInDateTime || undefined : undefined,
              status: u.accountEnabled === false ? 'Disabled' : 'Active'
            });
          } else if (r.status === 404) {
            result.push({ id, status: 'Deleted' });
          } else if (r.status === 403 && withSignIn) {
            forbiddenSignIn = true;
            retry.push(id);
          } else if (r.status === 429 || r.status >= 500) {
            const ra = r.headers && (r.headers['Retry-After'] || r.headers['retry-after']);
            waitMs = Math.max(waitMs, (parseInt(ra, 10) || 2) * 1000);
            retry.push(id);
          } else {
            result.push({ id, status: 'Unknown' });
          }
        });
        if (forbiddenSignIn) withSignIn = false; // licence or AuditLog.Read.All missing
        if (waitMs) await new Promise<void>((resolve) => setTimeout(resolve, Math.min(waitMs, 60000)));
        // eslint-disable-next-line require-atomic-updates -- local variable, no shared state
        chunk = retry;
      }
      chunk.forEach((id) => result.push({ id, status: 'Unknown' }));
    }
    return { users: result, signInUnavailable: includeSignIn && !withSignIn };
  }

  public async getDirectoryRoles(): Promise<IDirectoryRole[]> {
    const roles = await this._all<any>(`${V1}/directoryRoles?$select=id,displayName`);
    const out: IDirectoryRole[] = [];
    for (const r of roles) {
      const members = await this._all<any>(`${V1}/directoryRoles/${r.id}/members?$select=id,displayName,userPrincipalName`);
      out.push({
        id: str(r.id),
        displayName: str(r.displayName),
        members: members.map((m) => ({
          id: str(m.id),
          displayName: str(m.displayName),
          upn: str(m.userPrincipalName),
          type: str(m['@odata.type']).replace('#microsoft.graph.', '') || 'user'
        }))
      });
    }
    return out.filter((r) => r.members.length > 0);
  }

  public async getUsers(maxItems: number): Promise<{ users: IUserAccount[]; signInUnavailable: boolean }> {
    const base = 'id,displayName,userPrincipalName,userType,accountEnabled,createdDateTime';
    const map = (u: any): IUserAccount => ({
      id: str(u.id),
      displayName: str(u.displayName),
      upn: str(u.userPrincipalName),
      userType: str(u.userType) || 'Member',
      accountEnabled: u.accountEnabled !== false,
      createdTime: u.createdDateTime || undefined,
      lastSignIn: u.signInActivity ? u.signInActivity.lastSignInDateTime || undefined : undefined
    });
    try {
      const raw = await this._all<any>(`${V1}/users?$select=${base},signInActivity&$top=500`, maxItems);
      return { users: raw.map(map), signInUnavailable: false };
    } catch (e) {
      if (!isApiError(e) || (e.kind !== 'forbidden' && e.status !== 400)) throw e;
      const raw = await this._all<any>(`${V1}/users?$select=${base}&$top=999`, maxItems);
      return { users: raw.map(map), signInUnavailable: true };
    }
  }

  public async getMfaRegistration(maxItems: number): Promise<IMfaRegistration[]> {
    const raw = await this._all<any>(`${V1}/reports/authenticationMethods/userRegistrationDetails?$top=999`, maxItems);
    return raw.map((r) => ({
      id: str(r.id),
      upn: str(r.userPrincipalName),
      displayName: str(r.userDisplayName),
      isMfaRegistered: !!r.isMfaRegistered,
      isAdmin: !!r.isAdmin,
      userType: str(r.userType) || 'member',
      methods: Array.isArray(r.methodsRegistered) ? r.methodsRegistered.map(str) : []
    }));
  }

  public async getAppCredentials(): Promise<IAppCredential[]> {
    const raw = await this._all<any>(`${V1}/applications?$select=id,appId,displayName,passwordCredentials,keyCredentials&$top=999`);
    const now = Date.now();
    const out: IAppCredential[] = [];
    raw.forEach((a) => {
      const push = (c: any, type: 'Secret' | 'Certificate'): void => {
        if (!c || !c.endDateTime) return;
        out.push({
          appId: str(a.appId),
          appDisplayName: str(a.displayName),
          type,
          displayName: str(c.displayName) || (type === 'Secret' ? 'Client secret' : 'Certificate'),
          endDateTime: c.endDateTime,
          daysLeft: Math.floor((Date.parse(c.endDateTime) - now) / 86400000)
        });
      };
      (a.passwordCredentials || []).forEach((c: any) => push(c, 'Secret'));
      (a.keyCredentials || []).forEach((c: any) => push(c, 'Certificate'));
    });
    return out;
  }

  /* ---------- Service health ---------- */

  public async getHealthOverviews(): Promise<IHealthOverview[]> {
    const raw = await this._all<any>(`${V1}/admin/serviceAnnouncement/healthOverviews?$select=service,status`);
    return raw.map((h) => ({ service: str(h.service), status: str(h.status) }));
  }

  public async getHealthIssues(): Promise<IHealthIssue[]> {
    const raw = await this._all<any>(`${V1}/admin/serviceAnnouncement/issues?$filter=isResolved eq false`, 300);
    return raw.map((i) => ({
      id: str(i.id),
      title: str(i.title),
      service: str(i.service),
      status: str(i.status),
      classification: str(i.classification),
      startDateTime: i.startDateTime || undefined
    }));
  }

  public async getMessages(): Promise<IMessage[]> {
    const raw = await this._all<any>(`${V1}/admin/serviceAnnouncement/messages?$top=100`, 300);
    return raw
      .map((m) => ({
        id: str(m.id),
        title: str(m.title),
        services: Array.isArray(m.services) ? m.services.map(str) : [],
        category: str(m.category),
        lastModified: m.lastModifiedDateTime || m.startDateTime || undefined,
        actionRequiredBy: m.actionRequiredByDateTime || undefined,
        isMajorChange: !!m.isMajorChange
      }))
      .sort((a, b) => (str(b.lastModified) > str(a.lastModified) ? 1 : -1));
  }
}
