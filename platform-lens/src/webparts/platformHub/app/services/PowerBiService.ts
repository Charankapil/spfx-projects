/* eslint-disable @typescript-eslint/no-explicit-any */
import { ApiClient, isApiError } from '../http/ApiClient';
import {
  IPbiAccess, IPbiActivity, IPbiApp, IPbiCapacity, IPbiDashboard, IPbiDataflow, IPbiDataset, IPbiReport, IPbiWorkspace
} from '../models';
import { activityWindows, IRawPbiEvent, PbiActivityAggregator } from '../logic/pbiActivity';

/** Entra ID resource for the Power BI REST API ("Power BI Service"). */
export const POWERBI_RESOURCE = 'https://analysis.windows.net/powerbi/api';
const ADMIN = 'https://api.powerbi.com/v1.0/myorg/admin';

function str(v: any): string {
  return v === undefined || v === null ? '' : String(v);
}

export function mapWorkspaceType(t: string): string {
  if (t === 'PersonalGroup' || t === 'Personal') return 'Personal';
  if (t === 'Group') return 'Classic';
  if (t === 'Workspace') return 'Workspace';
  return t || 'Unknown';
}

export function mapStorageMode(m: string): string {
  if (m === 'Abf') return 'Import';
  if (m === 'PremiumFiles') return 'Large model';
  return m || 'Unknown';
}

export interface IPbiContent {
  workspaces: IPbiWorkspace[];
  reports: IPbiReport[];
  datasets: IPbiDataset[];
  dashboards: IPbiDashboard[];
  dataflows: IPbiDataflow[];
}

/** Maps one admin/groups row (with $expand) into normalised workspace + content records. */
export function mapWorkspace(g: any, into: IPbiContent): void {
  const id = str(g.id);
  const name = str(g.name) || '(unnamed)';
  const users: IPbiAccess[] = (Array.isArray(g.users) ? g.users : []).map((u: any) => ({
    name: str(u.displayName || u.emailAddress || u.identifier),
    email: str(u.emailAddress),
    right: str(u.groupUserAccessRight),
    principalType: str(u.principalType) || 'User'
  }));
  const reports: any[] = Array.isArray(g.reports) ? g.reports : [];
  const datasets: any[] = Array.isArray(g.datasets) ? g.datasets : [];
  const dashboards: any[] = Array.isArray(g.dashboards) ? g.dashboards : [];
  const dataflows: any[] = Array.isArray(g.dataflows) ? g.dataflows : [];
  into.workspaces.push({
    id, name,
    type: mapWorkspaceType(str(g.type)),
    state: str(g.state) || 'Active',
    isOnDedicatedCapacity: !!g.isOnDedicatedCapacity,
    capacityId: g.capacityId || undefined,
    users,
    admins: users.filter((u) => u.right === 'Admin').length,
    reports: reports.length,
    datasets: datasets.length,
    dashboards: dashboards.length,
    dataflows: dataflows.length
  });
  reports.forEach((r) => into.reports.push({
    id: str(r.id), name: str(r.name), workspaceId: id, workspaceName: name,
    datasetId: r.datasetId || undefined,
    reportType: str(r.reportType) === 'PaginatedReport' ? 'Paginated' : 'Power BI',
    createdBy: r.createdBy || undefined,
    modifiedBy: r.modifiedBy || undefined,
    createdTime: r.createdDateTime || undefined,
    modifiedTime: r.modifiedDateTime || undefined,
    endorsement: r.endorsementDetails && r.endorsementDetails.endorsement !== 'None' ? str(r.endorsementDetails.endorsement) : undefined
  }));
  datasets.forEach((d) => into.datasets.push({
    id: str(d.id), name: str(d.name), workspaceId: id, workspaceName: name,
    configuredBy: str(d.configuredBy),
    createdTime: d.createdDate || undefined,
    isRefreshable: !!d.isRefreshable,
    gatewayRequired: !!d.isOnPremGatewayRequired,
    storageMode: mapStorageMode(str(d.targetStorageMode)),
    endorsement: d.endorsementDetails && d.endorsementDetails.endorsement !== 'None' ? str(d.endorsementDetails.endorsement) : undefined
  }));
  dashboards.forEach((d) => into.dashboards.push({
    id: str(d.id), name: str(d.displayName || d.name), workspaceId: id, workspaceName: name, isReadOnly: !!d.isReadOnly
  }));
  dataflows.forEach((d) => into.dataflows.push({
    id: str(d.objectId || d.id), name: str(d.name), workspaceId: id, workspaceName: name,
    configuredBy: str(d.configuredBy), modifiedTime: d.modifiedDateTime || undefined
  }));
}

export interface IRefreshable {
  id: string;
  status?: string;
  time?: string;
  refreshCount?: number;
  refreshFailures?: number;
  scheduleEnabled?: boolean;
}

/**
 * Power BI / Fabric admin REST API. Admin calls are rate limited (roughly 50–200
 * requests per hour), so content is read with as few calls as possible:
 * one admin/groups call per 5,000 workspaces returns all reports, semantic
 * models, dashboards, dataflows and access lists.
 */
export class PowerBiService {
  private readonly _api: ApiClient;

  constructor(api: ApiClient) {
    this._api = api;
  }

  private async _skipPaged(path: string, top: number, maxItems: number): Promise<any[]> {
    const out: any[] = [];
    for (let skip = 0; ; skip += top) {
      const sep = path.indexOf('?') >= 0 ? '&' : '?';
      const page = await this._api.getJson<{ value?: any[] }>(POWERBI_RESOURCE, `${ADMIN}/${path}${sep}$top=${top}&$skip=${skip}`);
      const value = Array.isArray(page.value) ? page.value : [];
      value.forEach((v) => out.push(v));
      if (value.length < top || out.length >= maxItems) break;
    }
    return out;
  }

  public async getContent(maxWorkspaces: number): Promise<IPbiContent> {
    const raw = await this._skipPaged('groups?$expand=users,reports,dashboards,datasets,dataflows', 5000, maxWorkspaces);
    const content: IPbiContent = { workspaces: [], reports: [], datasets: [], dashboards: [], dataflows: [] };
    raw.forEach((g) => mapWorkspace(g, content));
    return content;
  }

  public async getApps(): Promise<IPbiApp[]> {
    const raw = await this._skipPaged('apps', 5000, 50000);
    return raw.map((a) => ({
      id: str(a.id), name: str(a.name), workspaceId: a.workspaceId || undefined,
      publishedBy: str(a.publishedBy), lastUpdate: a.lastUpdate || undefined
    }));
  }

  public async getCapacities(): Promise<IPbiCapacity[]> {
    const page = await this._api.getJson<{ value?: any[] }>(POWERBI_RESOURCE, `${ADMIN}/capacities`);
    return (page.value || []).map((c) => ({
      id: str(c.id), name: str(c.displayName), sku: str(c.sku), state: str(c.state), region: str(c.region),
      admins: Array.isArray(c.admins) ? c.admins.map(str) : []
    }));
  }

  /** Refresh history for semantic models on Premium/Fabric capacities. */
  public async getRefreshables(): Promise<IRefreshable[]> {
    const raw = await this._skipPaged('capacities/refreshables?$expand=capacity,group', 1000, 50000);
    return raw.map((r) => ({
      id: str(r.id),
      status: r.lastRefresh ? str(r.lastRefresh.status) : undefined,
      time: r.lastRefresh ? r.lastRefresh.endTime || r.lastRefresh.startTime : undefined,
      refreshCount: typeof r.refreshCount === 'number' ? r.refreshCount : undefined,
      refreshFailures: typeof r.refreshFailures === 'number' ? r.refreshFailures : undefined,
      scheduleEnabled: r.refreshSchedule ? !!r.refreshSchedule.enabled : undefined
    }));
  }

  /**
   * Streams the activity log for the last `days` UTC days (max 30, the API's
   * retention window) into an aggregator. A failed day is reported, not fatal.
   */
  public async getActivity(days: number, maxEvents: number, onDay?: (done: number, total: number) => void): Promise<IPbiActivity> {
    const windows = activityWindows(Math.max(1, Math.min(30, days)));
    const agg = new PbiActivityAggregator();
    const failed: string[] = [];
    let count = 0;
    let firstError: unknown;
    for (let i = 0; i < windows.length; i++) {
      const w = windows[i];
      let url: string | undefined = `${ADMIN}/activityevents?startDateTime='${w.start}'&endDateTime='${w.end}'`;
      try {
        while (url && count < maxEvents) {
          const page: { activityEventEntities?: IRawPbiEvent[]; continuationUri?: string; lastResultSet?: boolean } =
            await this._api.getJson(POWERBI_RESOURCE, url);
          (page.activityEventEntities || []).forEach((e) => {
            if (count < maxEvents) {
              agg.add(e);
              count++;
            }
          });
          url = page.lastResultSet || !page.continuationUri ? undefined : page.continuationUri;
        }
      } catch (e) {
        if (!firstError) firstError = e;
        // Permission problems fail every day the same way: stop instead of retrying 30 times.
        if (isApiError(e) && (e.kind === 'forbidden' || e.kind === 'blocked')) throw e;
        failed.push(w.day);
      }
      if (onDay) onDay(i + 1, windows.length);
    }
    if (failed.length === windows.length && firstError) throw firstError;
    const result = agg.result(windows.length, windows[0].start, windows[windows.length - 1].end, failed);
    result.truncated = count >= maxEvents;
    return result;
  }
}
