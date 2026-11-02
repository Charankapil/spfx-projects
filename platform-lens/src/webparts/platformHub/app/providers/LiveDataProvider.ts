import { ApiClient, describeError, IHttpTransport, mapWithConcurrency } from '../http/ApiClient';
import {
  IAdoptionData, IAppCredential, IAgent, IApp, IDlpPolicy, IEnvLoadError, IFlow, IIdentityData, IInventory, IMaker, IProgress,
  IServiceHealth, ISitesData, ISku, ITeamsData, IUserInfo
} from '../models';
import { GraphService } from '../services/GraphService';
import { DataverseService, PowerPlatformService } from '../services/PowerPlatformService';
import { BaseProvider, IAppSharing, IDataProvider, IFlowRun, IMakersData } from './IDataProvider';

export interface ILiveProviderOptions {
  tenantLabel: string;
  includeDataverse: boolean;
  /** Upper bound for large user-level lists (users, sites, report rows). */
  maxRows: number;
}

/** Calls Microsoft Graph, Power Platform admin APIs and Dataverse with the signed-in admin's delegated token. */
export class LiveDataProvider extends BaseProvider implements IDataProvider {
  public readonly mode = 'live' as const;
  public readonly tenantLabel: string;
  private readonly _graph: GraphService;
  private readonly _pp: PowerPlatformService;
  private readonly _dv: DataverseService;
  private readonly _options: ILiveProviderOptions;

  constructor(transport: IHttpTransport, options: ILiveProviderOptions) {
    super();
    const api = new ApiClient(transport);
    this._graph = new GraphService(api);
    this._pp = new PowerPlatformService(api);
    this._dv = new DataverseService(api);
    this._options = options;
    this.tenantLabel = options.tenantLabel;
  }

  public getInventory(): Promise<IInventory> {
    return this._cached('inventory', async () => {
      const environments = await this._track('pp-env', () => this._pp.listEnvironments());
      const apps: IApp[] = [];
      const flows: IFlow[] = [];
      const agents: IAgent[] = [];
      const errors: IEnvLoadError[] = [];
      const counters: Record<string, { ok: number; failed: number; last?: unknown }> = {
        'pp-apps': { ok: 0, failed: 0 }, 'pp-flows': { ok: 0, failed: 0 }, dv: { ok: 0, failed: 0 }
      };
      let done = 0;
      this._setProgress({ label: 'Reading environments', done, total: environments.length });

      const attempt = async <T>(source: string, area: IEnvLoadError['area'], envName: string, envId: string, work: () => Promise<T[]>): Promise<T[]> => {
        try {
          const r = await work();
          counters[source].ok++;
          return r;
        } catch (e) {
          counters[source].failed++;
          counters[source].last = e;
          errors.push({ environmentId: envId, environmentName: envName, area, message: describeError(e) });
          return [];
        }
      };

      await mapWithConcurrency(environments, 4, async (env) => {
        const [envApps, envFlows] = await Promise.all([
          attempt('pp-apps', 'Apps', env.displayName, env.id, () => this._pp.listApps(env)),
          attempt('pp-flows', 'Flows', env.displayName, env.id, () => this._pp.listFlows(env))
        ]);
        let modelDriven: IApp[] = [];
        if (this._options.includeDataverse && env.orgUrl) {
          const [bots, mda] = await Promise.all([
            attempt('dv', 'Agents', env.displayName, env.id, () => this._dv.listAgents(env)),
            attempt('dv', 'Model-driven apps', env.displayName, env.id, () => this._dv.listModelDrivenApps(env))
          ]);
          bots.forEach((b) => agents.push(b));
          modelDriven = mda;
        }
        // Canvas apps come from the Power Apps admin API; model-driven apps from Dataverse when it is read.
        envApps.forEach((a) => {
          if (a.kind === 'Canvas' || !modelDriven.length) apps.push(a);
        });
        modelDriven.forEach((a) => apps.push(a));
        envFlows.forEach((f) => flows.push(f));
        done++;
        this._setProgress({ label: `Read ${env.displayName}`, done, total: environments.length });
      });

      Object.keys(counters).forEach((id) => {
        const c = counters[id];
        if (id === 'dv' && !this._options.includeDataverse) {
          this._mark('dv', 'skipped', 'Turned off in the web part settings.');
        } else if (c.ok && !c.failed) {
          this._mark(id, 'ok');
        } else if (c.ok && c.failed) {
          this._mark(id, 'ok', `${c.failed} environment request(s) failed. See the environment errors on the Environments page.`);
        } else if (c.failed) {
          this._fail(id, c.last);
        } else if (id === 'dv') {
          this._mark('dv', 'skipped', 'No environment has a Dataverse database.');
        }
      });
      this._setProgress(undefined);
      return { environments, apps, flows, agents, errors, loadedAt: new Date().toISOString() };
    });
  }

  public getMakers(): Promise<IMakersData> {
    return this._cached('makers', async () => {
      const inv = await this.getInventory();
      const ids: string[] = [];
      inv.apps.forEach((a) => a.owner.id && ids.push(a.owner.id));
      inv.flows.forEach((f) => f.owner.id && ids.push(f.owner.id));
      inv.agents.forEach((b) => b.owner.id && ids.push(b.owner.id));
      const lookup = await this._optional(
        'g-users',
        () => this._graph.resolveUsers(ids, true),
        { users: [] as IUserInfo[], signInUnavailable: true }
      );
      if (lookup.users.length || !ids.length) {
        this._mark('g-signin', lookup.signInUnavailable ? 'forbidden' : 'ok', lookup.signInUnavailable ? 'Sign-in activity needs AuditLog.Read.All and Entra ID P1.' : undefined);
      }
      const users: Record<string, IUserInfo> = {};
      lookup.users.forEach((u) => {
        users[u.id] = u;
      });
      return { makers: buildMakers(inv, users), users, signInUnavailable: lookup.signInUnavailable };
    });
  }

  public getDlpPolicies(): Promise<IDlpPolicy[]> {
    return this._cached('dlp', () => this._track('pp-dlp', () => this._pp.listDlpPolicies()));
  }

  public getSkus(): Promise<ISku[]> {
    return this._cached('skus', () => this._track('g-skus', () => this._graph.getSubscribedSkus()));
  }

  public getServiceHealth(): Promise<IServiceHealth> {
    return this._cached('health', async () => {
      const [overviews, issues, messages] = await Promise.all([
        this._optional('g-health', () => this._graph.getHealthOverviews(), []),
        this._optional('g-health', () => this._graph.getHealthIssues(), []),
        this._optional('g-messages', () => this._graph.getMessages(), [])
      ]);
      return { overviews, issues, messages };
    });
  }

  public getReportsConcealed(): Promise<boolean> {
    return this._cached('concealed', () => this._optional('g-reportsettings', () => this._graph.getReportsConcealed(), false));
  }

  public getSites(): Promise<ISitesData> {
    return this._cached('sites', async () => {
      const [sites, storageTrend, sharing] = await Promise.all([
        this._track('g-reports', () => this._graph.getSites(this._options.maxRows)),
        this._optional('g-reports', () => this._graph.getStorageTrend(), []),
        this._optional('g-sharing', () => this._graph.getSharingSettings(), undefined)
      ]);
      return { sites, storageTrend, sharing };
    });
  }

  public getTeams(): Promise<ITeamsData> {
    return this._cached('teams', async () => {
      const [groups, activity] = await Promise.all([
        this._track('g-groups', () => this._graph.getGroups(this._options.maxRows)),
        this._optional('g-reports', () => this._graph.getTeamsActivity(), [])
      ]);
      return { groups, activity };
    });
  }

  public getAdoption(): Promise<IAdoptionData> {
    return this._cached('adoption', async () => {
      const [counts, users, copilot] = await Promise.all([
        this._track('g-reports', () => this._graph.getActiveUserCounts()),
        this._optional('g-reports', () => this._graph.getUserActivity(this._options.maxRows), []),
        this._optional('g-copilot', () => this._graph.getCopilotUsers().then((u) => ({ u, ok: true })), { u: [], ok: false })
      ]);
      return { counts, users, copilotUsers: copilot.u, copilotUnavailable: !copilot.ok };
    });
  }

  public getIdentity(): Promise<IIdentityData> {
    return this._cached('identity', async () => {
      const [roles, users, mfa, credentials] = await Promise.all([
        this._optional('g-roles', () => this._graph.getDirectoryRoles(), []),
        this._track('g-users', () => this._graph.getUsers(this._options.maxRows)),
        this._optional('g-mfa', () => this._graph.getMfaRegistration(this._options.maxRows).then((m) => ({ m, ok: true })), { m: [], ok: false }),
        this.getAppCredentials().then((c) => c, () => [])
      ]);
      this._mark('g-signin', users.signInUnavailable ? 'forbidden' : 'ok', users.signInUnavailable ? 'Sign-in activity needs AuditLog.Read.All and Entra ID P1.' : undefined);
      return { roles, users: users.users, signInUnavailable: users.signInUnavailable, mfa: mfa.m, mfaUnavailable: !mfa.ok, credentials };
    });
  }

  public getAppCredentials(): Promise<IAppCredential[]> {
    return this._cached('credentials', () => this._track('g-apps', () => this._graph.getAppCredentials()));
  }

  public getAppSharing(app: IApp): Promise<IAppSharing> {
    return this._pp.getAppPermissions(app.environmentId, app.id).then((s) => {
      // Cached inventory objects are updated in place so every page sees the result.
      app.sharedWithEveryone = s.everyone;
      app.sharedUsers = s.users;
      app.sharedGroups = s.groups;
      return s;
    });
  }

  public getFlowRuns(flow: IFlow): Promise<IFlowRun[]> {
    return this._pp.getFlowRuns(flow.environmentId, flow.id);
  }

  public async scanAppSharing(apps: IApp[], onProgress: (p: IProgress) => void): Promise<number> {
    const canvas = apps.filter((a) => a.kind === 'Canvas');
    let done = 0;
    let everyone = 0;
    await mapWithConcurrency(canvas, 6, async (app) => {
      try {
        const s = await this.getAppSharing(app);
        if (s.everyone) everyone++;
      } catch {
        /* individual failures are tolerated; the app keeps its unknown state */
      }
      done++;
      if (done % 5 === 0 || done === canvas.length) onProgress({ label: 'Checking app sharing', done, total: canvas.length });
    });
    this._emit();
    return everyone;
  }
}

/** Aggregates per-maker counts from the inventory and the resolved user directory. */
export function buildMakers(inv: IInventory, users: Record<string, IUserInfo>): IMaker[] {
  const makers: Record<string, IMaker & { envs: Record<string, boolean> }> = {};
  const touch = (id: string | undefined, displayName: string | undefined, email: string | undefined, envId: string, created: string | undefined, kind: 'apps' | 'flows' | 'agents'): void => {
    if (!id) return;
    const u = users[id];
    let m = makers[id];
    if (!m) {
      m = makers[id] = {
        id,
        displayName: (u && u.displayName) || displayName || (u && u.status === 'Deleted' ? 'Deleted user' : id),
        upn: (u && u.upn) || email || '',
        department: (u && u.department) || '',
        country: (u && u.country) || '',
        status: u ? u.status : 'Unknown',
        lastSignIn: u ? u.lastSignIn : undefined,
        apps: 0, flows: 0, agents: 0, environments: 0,
        envs: {}
      };
    }
    m[kind]++;
    m.envs[envId] = true;
    if (created) {
      if (!m.firstCreated || created < m.firstCreated) m.firstCreated = created;
      if (!m.lastCreated || created > m.lastCreated) m.lastCreated = created;
    }
  };
  inv.apps.forEach((a) => touch(a.owner.id, a.owner.displayName, a.owner.email, a.environmentId, a.createdTime, 'apps'));
  inv.flows.forEach((f) => touch(f.owner.id, f.owner.displayName, f.owner.email, f.environmentId, f.createdTime, 'flows'));
  inv.agents.forEach((b) => touch(b.owner.id, b.owner.displayName, b.owner.email, b.environmentId, b.createdTime, 'agents'));
  return Object.keys(makers).map((id) => {
    const m = makers[id];
    const { envs, ...rest } = m;
    return { ...rest, environments: Object.keys(envs).length };
  });
}
