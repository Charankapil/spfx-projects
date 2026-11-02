import {
  IAdoptionData, IAppCredential, IDlpPolicy, IFlow, IApp, IIdentityData, IInventory, IMaker, IProgress, IServiceHealth, ISitesData,
  ISku, ISourceStatus, ITeamsData, IUserInfo, SourceState
} from '../models';
import { ApiError, describeError, isApiError } from '../http/ApiClient';

export interface IMakersData {
  makers: IMaker[];
  users: Record<string, IUserInfo>;
  signInUnavailable: boolean;
}

export interface IAppSharing {
  users: number;
  groups: number;
  everyone: boolean;
  principals: { name: string; type: string; role: string }[];
}

export interface IFlowRun {
  status: string;
  startTime?: string;
  endTime?: string;
}

export interface IDataProvider {
  readonly mode: 'live' | 'sample';
  readonly tenantLabel: string;

  getInventory(): Promise<IInventory>;
  getMakers(): Promise<IMakersData>;
  getDlpPolicies(): Promise<IDlpPolicy[]>;
  getSkus(): Promise<ISku[]>;
  getServiceHealth(): Promise<IServiceHealth>;
  getSites(): Promise<ISitesData>;
  getTeams(): Promise<ITeamsData>;
  getAdoption(): Promise<IAdoptionData>;
  getIdentity(): Promise<IIdentityData>;
  /** App registration secrets and certificates (cheaper than the full identity load). */
  getAppCredentials(): Promise<IAppCredential[]>;
  getReportsConcealed(): Promise<boolean>;

  getAppSharing(app: IApp): Promise<IAppSharing>;
  getFlowRuns(flow: IFlow): Promise<IFlowRun[]>;
  /** Looks up "shared with everyone" for many apps; updates the cached inventory in place. */
  scanAppSharing(apps: IApp[], onProgress: (p: IProgress) => void): Promise<number>;

  getSources(): ISourceStatus[];
  getProgress(): IProgress | undefined;
  subscribe(listener: () => void): () => void;
  refresh(): void;
}

export interface ISourceDefinition {
  id: string;
  label: string;
  endpoint: string;
  permission: string;
}

export const SOURCES: ISourceDefinition[] = [
  { id: 'pp-env', label: 'Power Platform environments', endpoint: 'api.bap.microsoft.com', permission: 'PowerApps Service · User, plus the Power Platform administrator role' },
  { id: 'pp-apps', label: 'Power Apps inventory', endpoint: 'api.powerapps.com', permission: 'PowerApps Service · User' },
  { id: 'pp-flows', label: 'Cloud flow inventory', endpoint: 'api.flow.microsoft.com', permission: 'Microsoft Flow Service · Flows.Read.All' },
  { id: 'pp-dlp', label: 'DLP policies', endpoint: 'api.bap.microsoft.com', permission: 'PowerApps Service · User' },
  { id: 'dv', label: 'Dataverse (agents, model-driven apps)', endpoint: '<org>.crm.dynamics.com', permission: 'Dynamics CRM · user_impersonation, plus a Dataverse role in each environment' },
  { id: 'g-users', label: 'Maker and owner lookup', endpoint: 'graph.microsoft.com /users', permission: 'Microsoft Graph · User.Read.All' },
  { id: 'g-signin', label: 'Sign-in activity', endpoint: 'graph.microsoft.com signInActivity', permission: 'Microsoft Graph · AuditLog.Read.All (Entra ID P1)' },
  { id: 'g-skus', label: 'Licences', endpoint: 'graph.microsoft.com /subscribedSkus', permission: 'Microsoft Graph · Organization.Read.All' },
  { id: 'g-health', label: 'Service health', endpoint: 'graph.microsoft.com /admin/serviceAnnouncement', permission: 'Microsoft Graph · ServiceHealth.Read.All' },
  { id: 'g-messages', label: 'Message center', endpoint: 'graph.microsoft.com /admin/serviceAnnouncement/messages', permission: 'Microsoft Graph · ServiceMessage.Read.All' },
  { id: 'g-reports', label: 'Usage reports', endpoint: 'graph.microsoft.com /reports', permission: 'Microsoft Graph · Reports.Read.All, plus a reports reader role' },
  { id: 'g-reportsettings', label: 'Report privacy setting', endpoint: 'graph.microsoft.com /admin/reportSettings', permission: 'Microsoft Graph · ReportSettings.Read.All' },
  { id: 'g-copilot', label: 'Microsoft 365 Copilot usage', endpoint: 'graph.microsoft.com /beta/reports', permission: 'Microsoft Graph · Reports.Read.All' },
  { id: 'g-sharing', label: 'SharePoint sharing settings', endpoint: 'graph.microsoft.com /admin/sharepoint/settings', permission: 'Microsoft Graph · SharePointTenantSettings.Read.All' },
  { id: 'g-groups', label: 'Microsoft 365 groups and teams', endpoint: 'graph.microsoft.com /groups', permission: 'Microsoft Graph · Group.Read.All' },
  { id: 'g-roles', label: 'Admin role assignments', endpoint: 'graph.microsoft.com /directoryRoles', permission: 'Microsoft Graph · RoleManagement.Read.Directory' },
  { id: 'g-mfa', label: 'MFA registration', endpoint: 'graph.microsoft.com /reports/authenticationMethods', permission: 'Microsoft Graph · AuditLog.Read.All (Entra ID P1)' },
  { id: 'g-apps', label: 'App registrations', endpoint: 'graph.microsoft.com /applications', permission: 'Microsoft Graph · Application.Read.All' }
];

/** Shared caching, progress and diagnostics for both providers. */
export abstract class BaseProvider {
  private _cache: Record<string, Promise<unknown>> = {};
  private _listeners: (() => void)[] = [];
  private _progress: IProgress | undefined;
  protected _sources: Record<string, ISourceStatus> = {};

  constructor() {
    this._resetSources();
  }

  private _resetSources(): void {
    this._sources = {};
    SOURCES.forEach((s) => {
      this._sources[s.id] = { ...s, state: 'pending' };
    });
  }

  public getSources(): ISourceStatus[] {
    return SOURCES.map((s) => this._sources[s.id]);
  }

  public getProgress(): IProgress | undefined {
    return this._progress;
  }

  public subscribe(listener: () => void): () => void {
    this._listeners.push(listener);
    return () => {
      this._listeners = this._listeners.filter((l) => l !== listener);
    };
  }

  public refresh(): void {
    this._cache = {};
    this._resetSources();
    this._emit();
  }

  protected _emit(): void {
    this._listeners.slice().forEach((l) => {
      try {
        l();
      } catch {
        /* listener errors must not break loading */
      }
    });
  }

  protected _setProgress(p: IProgress | undefined): void {
    this._progress = p;
    this._emit();
  }

  protected _cached<T>(key: string, loader: () => Promise<T>): Promise<T> {
    if (!this._cache[key]) {
      const p = loader();
      this._cache[key] = p;
      // Failed loads are not cached, so a retry can succeed after consent is granted.
      p.catch(() => {
        if (this._cache[key] === p) delete this._cache[key];
      });
    }
    return this._cache[key] as Promise<T>;
  }

  protected _mark(id: string, state: SourceState, message?: string): void {
    const s = this._sources[id];
    if (!s) return;
    // Never downgrade a source that already succeeded for some calls to "error" by a later partial failure.
    if (s.state === 'ok' && state !== 'ok' && message === undefined) return;
    this._sources[id] = { ...s, state, message, checkedAt: new Date().toISOString() };
    this._emit();
  }

  protected _fail(id: string, e: unknown): void {
    const kind = isApiError(e) ? (e as ApiError).kind : 'error';
    const state: SourceState = kind === 'forbidden' ? 'forbidden' : kind === 'blocked' ? 'blocked' : 'error';
    this._mark(id, state, describeError(e));
  }

  /** Runs `work`, recording success or failure against a data source. */
  protected async _track<T>(id: string, work: () => Promise<T>): Promise<T> {
    try {
      const result = await work();
      this._mark(id, 'ok');
      return result;
    } catch (e) {
      this._fail(id, e);
      throw e;
    }
  }

  /** Like `_track` but returns `fallback` instead of throwing. */
  protected async _optional<T>(id: string, work: () => Promise<T>, fallback: T): Promise<T> {
    try {
      return await this._track(id, work);
    } catch {
      return fallback;
    }
  }
}
