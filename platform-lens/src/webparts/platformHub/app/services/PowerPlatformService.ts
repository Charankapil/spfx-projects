/* eslint-disable @typescript-eslint/no-explicit-any */
import { ApiClient, isApiError } from '../http/ApiClient';
import { AgentAuth, FlowState, IAgent, IApp, IConnectorRef, IDlpPolicy, IEnvironment, IFlow } from '../models';
import { makeConnector, uniqueConnectors } from '../logic/connectors';

/** Token audiences (Entra ID resources). */
export const POWERAPPS_RESOURCE = 'https://service.powerapps.com/';
export const FLOW_RESOURCE = 'https://service.flow.microsoft.com/';

const BAP = 'https://api.bap.microsoft.com';
const POWERAPPS = 'https://api.powerapps.com';
const FLOW = 'https://api.flow.microsoft.com';

function str(v: any): string {
  return v === undefined || v === null ? '' : String(v);
}

function toMb(value: any, unit: any): number {
  const n = Number(value) || 0;
  const u = str(unit).toUpperCase();
  if (u === 'GB') return n * 1024;
  if (u === 'KB') return n / 1024;
  return n;
}

const REGIONS: Record<string, string> = {
  unitedstates: 'United States', unitedstatesfirstrelease: 'United States (preview)', europe: 'Europe', unitedkingdom: 'United Kingdom',
  asia: 'Asia', australia: 'Australia', india: 'India', japan: 'Japan', canada: 'Canada', southamerica: 'South America', france: 'France',
  germany: 'Germany', switzerland: 'Switzerland', norway: 'Norway', korea: 'Korea', southafrica: 'South Africa', unitedarabemirates: 'UAE',
  uae: 'UAE', singapore: 'Singapore', sweden: 'Sweden', italy: 'Italy', poland: 'Poland', spain: 'Spain', usgov: 'US Government (GCC)'
};

export function regionLabel(region: string): string {
  const key = region.toLowerCase();
  return REGIONS[key] || region;
}

export function mapEnvironment(e: any): IEnvironment {
  const p = e.properties || {};
  const linked = p.linkedEnvironmentMetadata || {};
  const capacity = { databaseMb: 0, fileMb: 0, logMb: 0 };
  (Array.isArray(p.capacity) ? p.capacity : []).forEach((c: any) => {
    const mb = toMb(c.actualConsumption, c.capacityUnit);
    const t = str(c.capacityType).toLowerCase();
    if (t.indexOf('database') >= 0) capacity.databaseMb += mb;
    else if (t.indexOf('file') >= 0) capacity.fileMb += mb;
    else if (t.indexOf('log') >= 0) capacity.logMb += mb;
  });
  const createdBy = p.createdBy || {};
  const governance = p.governanceConfiguration || {};
  const runtime = (p.states && p.states.runtime && p.states.runtime.id) || p.provisioningState || 'Unknown';
  const instanceUrl = str(linked.instanceUrl).replace(/\/+$/, '');
  return {
    id: str(e.name),
    displayName: str(p.displayName) || str(e.name),
    type: str(p.environmentSku) || (p.isDefault ? 'Default' : 'Unknown'),
    region: regionLabel(str(e.location || p.azureRegion)),
    isDefault: !!p.isDefault,
    isManaged: str(governance.protectionLevel).toLowerCase() === 'standard',
    hasDataverse: !!instanceUrl,
    orgUrl: instanceUrl || undefined,
    dataverseVersion: linked.version ? str(linked.version) : undefined,
    securityGroupId: linked.securityGroupId ? str(linked.securityGroupId) : undefined,
    state: str(runtime),
    createdTime: p.createdTime,
    createdBy: str(createdBy.displayName || createdBy.userPrincipalName || createdBy.email),
    expirationTime: p.expirationTime || undefined,
    capacity
  };
}

function connectionRefsToList(refs: any): any[] {
  if (!refs) return [];
  if (Array.isArray(refs)) return refs;
  return Object.keys(refs).map((k) => refs[k]);
}

export function mapApp(a: any, env: IEnvironment): IApp {
  const p = a.properties || {};
  const owner = p.owner || p.createdBy || {};
  const connectors = uniqueConnectors(
    connectionRefsToList(p.connectionReferences).map((c: any) =>
      makeConnector(c.id, c.displayName, c.apiTier, !!c.isCustomApiConnection)
    )
  );
  const appType = str(p.appType).toLowerCase();
  return {
    id: str(a.name),
    displayName: str(p.displayName) || str(a.name),
    environmentId: env.id,
    environmentName: env.displayName,
    kind: appType.indexOf('model') >= 0 ? 'Model-driven' : 'Canvas',
    owner: { id: owner.id, displayName: owner.displayName, email: owner.email || owner.userPrincipalName },
    createdTime: p.createdTime,
    lastModifiedTime: p.lastModifiedTime,
    lastPublishTime: p.lastPublishTime,
    connectors,
    sharedUsers: typeof p.sharedUsersCount === 'number' ? p.sharedUsersCount : undefined,
    sharedGroups: typeof p.sharedGroupsCount === 'number' ? p.sharedGroupsCount : undefined,
    usesPremium: !!p.usesPremiumApi || connectors.some((c) => c.tier === 'Premium'),
    usesCustom: !!p.usesCustomApi || connectors.some((c) => c.tier === 'Custom'),
    bypassConsent: !!p.bypassConsent
  };
}

export function classifyTrigger(t: any): string {
  if (!t) return 'Unknown';
  const type = str(t.type).toLowerCase();
  const kind = str(t.kind).toLowerCase();
  if (type === 'recurrence') return 'Scheduled';
  if (type === 'request') {
    if (kind === 'http') return 'HTTP request';
    if (kind === 'powerapp' || kind === 'powerappv2') return 'Power Apps';
    if (kind === 'button' || kind === 'skill' || kind === '') return 'Instant';
    return 'Instant';
  }
  if (type.indexOf('openapiconnection') === 0 || type.indexOf('apiconnection') === 0) return 'Automated';
  return str(t.type) || 'Unknown';
}

function mapFlowState(s: any): FlowState {
  const v = str(s).toLowerCase();
  if (v === 'started') return 'On';
  if (v === 'suspended') return 'Suspended';
  return 'Off';
}

export function mapFlow(f: any, env: IEnvironment): IFlow {
  const p = f.properties || {};
  const summary = p.definitionSummary || {};
  const actions: any[] = Array.isArray(summary.actions) ? summary.actions : [];
  const triggers: any[] = Array.isArray(summary.triggers) ? summary.triggers : [];
  const fromRefs = connectionRefsToList(p.connectionReferences).map((c: any) =>
    makeConnector(c.id || (c.api && c.api.id) || c.apiName, c.displayName, c.tier || c.apiTier, !!c.isCustomApiConnection)
  );
  const fromActions: (IConnectorRef | undefined)[] = actions.concat(triggers).map((a: any) => {
    if (a.api && (a.api.id || a.api.name)) return makeConnector(a.api.id || a.api.name, undefined, a.api.tier);
    if (str(a.type).toLowerCase() === 'http') return makeConnector('http', 'HTTP', 'Premium');
    return undefined;
  });
  const connectors = uniqueConnectors(fromRefs.concat(fromActions));
  const creator = p.creator || {};
  return {
    id: str(f.name),
    displayName: str(p.displayName) || str(f.name),
    environmentId: env.id,
    environmentName: env.displayName,
    state: mapFlowState(p.state),
    suspensionReason: p.flowSuspensionReason && p.flowSuspensionReason !== 'None' ? str(p.flowSuspensionReason) : undefined,
    owner: { id: creator.objectId || creator.userId },
    createdTime: p.createdTime,
    lastModifiedTime: p.lastModifiedTime,
    triggerType: classifyTrigger(triggers[0]),
    connectors,
    usesPremium: connectors.some((c) => c.tier === 'Premium'),
    usesCustom: connectors.some((c) => c.tier === 'Custom'),
    actionCount: actions.length
  };
}

function mapScope(t: any): IDlpPolicy['scope'] {
  const v = str(t).toLowerCase();
  if (v.indexOf('except') >= 0) return 'All except';
  if (v.indexOf('only') >= 0 || v === 'singleenvironment') return 'Only selected';
  return 'All environments';
}

export function mapDlpV2(p: any): IDlpPolicy {
  const groups: any[] = Array.isArray(p.connectorGroups) ? p.connectorGroups : [];
  const pick = (cls: string): string[] => {
    const g = groups.filter((x) => str(x.classification).toLowerCase() === cls)[0];
    return g && Array.isArray(g.connectors) ? g.connectors.map((c: any) => str(c.name || c.id)) : [];
  };
  return {
    id: str(p.name || p.id),
    displayName: str(p.displayName),
    scope: mapScope(p.environmentType),
    environmentIds: (Array.isArray(p.environments) ? p.environments : []).map((e: any) => str(e.name || e.id).split('/').pop() || ''),
    businessConnectors: pick('confidential'),
    nonBusinessConnectors: pick('general'),
    blockedConnectors: pick('blocked'),
    defaultGroup: str(p.defaultConnectorsClassification) === 'Confidential' ? 'Business' : str(p.defaultConnectorsClassification) === 'Blocked' ? 'Blocked' : 'Non-business',
    createdBy: p.createdBy ? str(p.createdBy.displayName) : undefined,
    lastModified: p.lastModifiedTime
  };
}

/** Legacy BAP apiPolicies shape (used only if v2 policies are unavailable). */
export function mapDlpV1(p: any): IDlpPolicy {
  const props = p.properties || {};
  const def = props.definition || {};
  const constraints = def.constraints || {};
  const envFilter = constraints.environmentFilter1 && constraints.environmentFilter1.parameters;
  const apiGroups = def.apiGroups || {};
  const list = (g: any): string[] => (g && Array.isArray(g.apis) ? g.apis.map((a: any) => str(a.id).split('/').pop() || '') : []);
  return {
    id: str(p.name),
    displayName: str(props.displayName),
    scope: envFilter ? (str(envFilter.filterType).toLowerCase() === 'exclude' ? 'All except' : 'Only selected') : 'All environments',
    environmentIds: envFilter && Array.isArray(envFilter.environments) ? envFilter.environments.map((e: any) => str(e.name)) : [],
    businessConnectors: list(apiGroups.hbi),
    nonBusinessConnectors: list(apiGroups.lbi),
    blockedConnectors: [],
    defaultGroup: str(def.defaultApiGroup) === 'hbi' ? 'Business' : 'Non-business',
    createdBy: props.createdBy ? str(props.createdBy.displayName) : undefined,
    lastModified: props.lastModifiedTime
  };
}

export interface IAppPermissionSummary {
  users: number;
  groups: number;
  everyone: boolean;
  principals: { name: string; type: string; role: string }[];
}

export class PowerPlatformService {
  private readonly _api: ApiClient;

  constructor(api: ApiClient) {
    this._api = api;
  }

  public async listEnvironments(): Promise<IEnvironment[]> {
    const url = `${BAP}/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments?api-version=2021-04-01&$expand=properties.capacity,properties.addons`;
    const raw = await this._api.getAll<any>(POWERAPPS_RESOURCE, url);
    return raw.map(mapEnvironment);
  }

  public async listApps(env: IEnvironment): Promise<IApp[]> {
    const url = `${POWERAPPS}/providers/Microsoft.PowerApps/scopes/admin/environments/${encodeURIComponent(env.id)}/apps?api-version=2016-11-01&$top=250`;
    const raw = await this._api.getAll<any>(POWERAPPS_RESOURCE, url);
    return raw.map((a) => mapApp(a, env));
  }

  public async listFlows(env: IEnvironment): Promise<IFlow[]> {
    const url = `${FLOW}/providers/Microsoft.ProcessSimple/scopes/admin/environments/${encodeURIComponent(env.id)}/v2/flows?api-version=2016-11-01&$top=50`;
    const raw = await this._api.getAll<any>(FLOW_RESOURCE, url);
    return raw.map((f) => mapFlow(f, env));
  }

  public async listDlpPolicies(): Promise<IDlpPolicy[]> {
    try {
      const raw = await this._api.getAll<any>(POWERAPPS_RESOURCE, `${BAP}/providers/PowerPlatform.Governance/v2/policies`);
      return raw.map(mapDlpV2);
    } catch (e) {
      if (isApiError(e) && (e.kind === 'notFound' || e.status === 400)) {
        const raw = await this._api.getAll<any>(
          POWERAPPS_RESOURCE,
          `${BAP}/providers/Microsoft.BusinessAppPlatform/scopes/admin/apiPolicies?api-version=2016-11-01`
        );
        return raw.map(mapDlpV1);
      }
      throw e;
    }
  }

  public async getAppPermissions(envId: string, appId: string): Promise<IAppPermissionSummary> {
    const url = `${POWERAPPS}/providers/Microsoft.PowerApps/scopes/admin/environments/${encodeURIComponent(envId)}/apps/${encodeURIComponent(appId)}/permissions?api-version=2016-11-01`;
    const raw = await this._api.getAll<any>(POWERAPPS_RESOURCE, url);
    const summary: IAppPermissionSummary = { users: 0, groups: 0, everyone: false, principals: [] };
    raw.forEach((r) => {
      const p = r.properties || {};
      const principal = p.principal || {};
      const type = str(principal.type);
      if (type === 'Tenant') summary.everyone = true;
      else if (type === 'Group') summary.groups++;
      else if (type === 'User') summary.users++;
      summary.principals.push({
        name: type === 'Tenant' ? 'Everyone in the organisation' : str(principal.displayName || principal.email || principal.id),
        type,
        role: str(p.roleName)
      });
    });
    return summary;
  }

  /** Recent runs for one flow (admin scope first, then maker scope). */
  public async getFlowRuns(envId: string, flowId: string): Promise<{ status: string; startTime?: string; endTime?: string }[]> {
    const urls = [
      `${FLOW}/providers/Microsoft.ProcessSimple/scopes/admin/environments/${encodeURIComponent(envId)}/flows/${encodeURIComponent(flowId)}/runs?api-version=2016-11-01&$top=50`,
      `${FLOW}/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(envId)}/flows/${encodeURIComponent(flowId)}/runs?api-version=2016-11-01&$top=50`
    ];
    let lastError: unknown;
    for (const url of urls) {
      try {
        const page = await this._api.getJson<{ value?: any[] }>(FLOW_RESOURCE, url);
        return (page.value || []).map((r) => {
          const p = r.properties || {};
          return { status: str(p.status), startTime: p.startTime, endTime: p.endTime };
        });
      } catch (e) {
        lastError = e;
      }
    }
    throw lastError;
  }
}

/* ---------- Dataverse (Copilot Studio agents, model-driven apps) ---------- */

function mapAuth(v: any): AgentAuth {
  // Copilot Studio bot.authenticationmode: 1 = No authentication, 2 = Integrated (Microsoft), 3 = Manual (custom).
  const n = Number(v);
  if (n === 1) return 'None';
  if (n === 2) return 'Microsoft';
  if (n === 3) return 'Manual';
  return 'Unknown';
}

export class DataverseService {
  private readonly _api: ApiClient;

  constructor(api: ApiClient) {
    this._api = api;
  }

  private static _resource(orgUrl: string): string {
    const m = /^https:\/\/[^/]+/i.exec(orgUrl);
    return m ? m[0] : orgUrl;
  }

  private async _getAll(orgUrl: string, path: string): Promise<any[]> {
    const resource = DataverseService._resource(orgUrl);
    try {
      return await this._api.getAll<any>(resource, `${resource}/api/data/v9.2/${path}`, {
        headers: { Prefer: 'odata.maxpagesize=500', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' }
      });
    } catch (e) {
      // Table not present (e.g. Copilot Studio never provisioned) is not an error for inventory purposes.
      if (isApiError(e) && e.kind === 'notFound') return [];
      throw e;
    }
  }

  public async listAgents(env: IEnvironment): Promise<IAgent[]> {
    if (!env.orgUrl) return [];
    const raw = await this._getAll(
      env.orgUrl,
      'bots?$select=botid,name,createdon,modifiedon,publishedon,authenticationmode,language,statecode' +
        '&$expand=owninguser($select=azureactivedirectoryobjectid,fullname,internalemailaddress)'
    );
    return raw.map((b) => {
      const owner = b.owninguser || {};
      return {
        id: str(b.botid),
        displayName: str(b.name),
        environmentId: env.id,
        environmentName: env.displayName,
        owner: { id: owner.azureactivedirectoryobjectid, displayName: owner.fullname, email: owner.internalemailaddress },
        createdTime: b.createdon,
        modifiedTime: b.modifiedon,
        publishedTime: b.publishedon || undefined,
        authMode: mapAuth(b.authenticationmode),
        language: b.language !== undefined && b.language !== null ? str(b.language) : undefined,
        state: Number(b.statecode) === 0 ? 'Active' : 'Inactive'
      } as IAgent;
    });
  }

  public async listModelDrivenApps(env: IEnvironment): Promise<IApp[]> {
    if (!env.orgUrl) return [];
    const raw = await this._getAll(
      env.orgUrl,
      'appmodules?$select=appmoduleid,name,uniquename,createdon,modifiedon,publishedon,ismanaged' +
        '&$expand=createdby($select=azureactivedirectoryobjectid,fullname,internalemailaddress)'
    );
    return raw.map((m) => {
      const owner = m.createdby || {};
      return {
        id: str(m.appmoduleid),
        displayName: str(m.name || m.uniquename),
        environmentId: env.id,
        environmentName: env.displayName,
        kind: 'Model-driven',
        owner: { id: owner.azureactivedirectoryobjectid, displayName: owner.fullname, email: owner.internalemailaddress },
        createdTime: m.createdon,
        lastModifiedTime: m.modifiedon,
        lastPublishTime: m.publishedon || undefined,
        connectors: [],
        isManaged: !!m.ismanaged,
        usesPremium: true, // model-driven apps always require a premium Power Apps licence
        usesCustom: false
      } as IApp;
    });
  }
}
