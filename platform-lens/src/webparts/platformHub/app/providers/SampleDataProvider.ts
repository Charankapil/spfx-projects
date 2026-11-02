import {
  AgentAuth, IAdoptionData, IAgent, IApp, IAppCredential, IConnectorRef, IDirectoryRole, IDlpPolicy, IEnvironment, IFlow,
  IGroup, IIdentityData, IInventory, IMfaRegistration, IProgress, IServiceHealth, ISite, ISitesData, ISku, ITeamActivity,
  ITeamsData, IUserAccount, IUserActivity, IUserInfo, FlowState
} from '../models';
import { makeConnector } from '../logic/connectors';
import { BaseProvider, IAppSharing, IDataProvider, IFlowRun, IMakersData, SOURCES } from './IDataProvider';
import { buildMakers } from './LiveDataProvider';
import { regionLabel } from '../services/PowerPlatformService';

/** Small deterministic PRNG (mulberry32) so sample data is stable between renders. */
function rng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const DAY = 86400000;
const FIRST = ['Ava', 'Ben', 'Chloe', 'Dev', 'Elena', 'Farid', 'Grace', 'Hiro', 'Isla', 'Jonas', 'Kemi', 'Liam', 'Maya', 'Noah', 'Olu', 'Priya', 'Quinn', 'Rosa', 'Sam', 'Tara', 'Uma', 'Victor', 'Wen', 'Yara', 'Zoe'];
const LAST = ['Adams', 'Bauer', 'Chen', 'Diaz', 'Evans', 'Fischer', 'Gupta', 'Hughes', 'Ito', 'Jensen', 'Khan', 'Lopez', 'Moreau', 'Novak', 'Okafor', 'Patel', 'Rossi', 'Silva', 'Tanaka', 'Weber'];
const DEPTS = ['Finance', 'Sales', 'Operations', 'HR', 'IT', 'Marketing', 'Legal', 'Customer Service', 'Procurement'];
const COUNTRIES = ['United Kingdom', 'United States', 'Germany', 'India', 'Australia', 'Netherlands'];
const APP_WORDS = ['Expense', 'Leave', 'Asset', 'Visitor', 'Onboarding', 'Invoice', 'Timesheet', 'Incident', 'Purchase', 'Inspection', 'Contract', 'Training', 'Booking', 'Survey', 'Inventory', 'Travel', 'Audit', 'Field', 'Case', 'Budget'];
const APP_SUFFIX = ['Tracker', 'Request', 'Portal', 'Manager', 'Hub', 'Approvals', 'Logger', 'Planner', 'Checker', 'Desk'];
const FLOW_VERBS = ['Notify', 'Sync', 'Archive', 'Approve', 'Copy', 'Escalate', 'Send', 'Update', 'Create', 'Collect'];
const FLOW_OBJECTS = ['new invoices', 'leave requests', 'Planner tasks', 'SharePoint items', 'Forms responses', 'daily report', 'contract renewals', 'Teams messages', 'expense claims', 'CRM leads'];

const STD = ['shared_sharepointonline', 'shared_office365', 'shared_office365users', 'shared_teams', 'shared_approvals', 'shared_excelonlinebusiness', 'shared_onedriveforbusiness', 'shared_planner', 'shared_microsoftforms', 'shared_flowpush'];
const PREM = ['shared_commondataserviceforapps', 'shared_sql', 'shared_http', 'shared_azureblob', 'shared_salesforce', 'shared_servicenow'];
const CUSTOM = ['shared_contoso-20erp-5f1e3c4b9a8d7e6f12', 'shared_contoso-20hr-20api-a1b2c3d4e5f60718'];

export class SampleDataProvider extends BaseProvider implements IDataProvider {
  public readonly mode = 'sample' as const;
  public readonly tenantLabel: string = 'Contoso (sample data)';
  private _inventory: IInventory | undefined;
  private _users: Record<string, IUserInfo> = {};
  private readonly _now: number;

  constructor(now: number = Date.now()) {
    super();
    this._now = now;
    SOURCES.forEach((s) => this._mark(s.id, 'ok', 'Sample data'));
  }

  public refresh(): void {
    super.refresh();
    SOURCES.forEach((s) => this._mark(s.id, 'ok', 'Sample data'));
  }

  private _iso(daysAgo: number): string {
    return new Date(this._now - daysAgo * DAY).toISOString();
  }

  private _delay<T>(value: T): Promise<T> {
    return new Promise<T>((resolve) => setTimeout(() => resolve(value), 250));
  }

  private _build(): IInventory {
    if (this._inventory) return this._inventory;
    const r = rng(42);
    const pick = <T>(arr: T[]): T => arr[Math.floor(r() * arr.length)];
    // Skews creation dates toward the recent past to produce a growth curve.
    const createdDaysAgo = (): number => Math.floor(Math.pow(r(), 1.3) * 720);

    const envDefs: [string, string, boolean, boolean][] = [
      ['Contoso (default)', 'Default', false, false], ['Finance-Prod', 'Production', true, true], ['HR-Prod', 'Production', true, true],
      ['Sales-Prod', 'Production', false, true], ['Operations-Prod', 'Production', true, false], ['Finance-Test', 'Sandbox', true, true],
      ['Sales-UAT', 'Sandbox', false, true], ['Dev - A. Patel', 'Developer', false, true], ['Dev - K. Okafor', 'Developer', false, true],
      ['Dev - unused', 'Developer', false, true], ['Pilot trial', 'Trial', false, true], ['Project Atlas (Teams)', 'Teams', false, true],
      ['Service Desk (Teams)', 'Teams', false, true], ['Copilot Agents', 'Production', true, true]
    ];
    const regions = ['europe', 'unitedstates', 'europe', 'unitedkingdom'];
    const environments: IEnvironment[] = envDefs.map(([name, type, managed, dataverse], i) => ({
      id: `${i === 0 ? 'Default-' : ''}${(0x1000 + i).toString(16)}a7c2-4f1e-9b3d-${(0x5000 + i * 7).toString(16)}e2f9`,
      displayName: name,
      type,
      region: regionLabel(regions[i % regions.length]),
      isDefault: i === 0,
      isManaged: managed,
      hasDataverse: dataverse,
      orgUrl: dataverse ? `https://org${(0xa10 + i).toString(16)}.crm11.dynamics.com` : undefined,
      dataverseVersion: dataverse ? '9.2.26103.00152' : undefined,
      securityGroupId: type === 'Production' && i !== 3 ? `sg-${i}` : undefined,
      state: 'Enabled',
      createdTime: this._iso(i === 0 ? 1600 : 60 + Math.floor(r() * 700)),
      createdBy: i === 0 ? 'System' : `${pick(FIRST)} ${pick(LAST)}`,
      expirationTime: type === 'Trial' ? this._iso(-12) : undefined,
      capacity: dataverse ? { databaseMb: Math.round(400 + r() * 9000), fileMb: Math.round(100 + r() * 14000), logMb: Math.round(r() * 900) } : { databaseMb: 0, fileMb: 0, logMb: 0 }
    }));

    const users: Record<string, IUserInfo> = {};
    const makerIds: string[] = [];
    for (let i = 0; i < 140; i++) {
      const id = `u-${(100000 + i * 7919).toString(16)}-4b1c-9e2a-${i.toString(16).padStart(4, '0')}`;
      const first = FIRST[i % FIRST.length];
      const last = LAST[(i * 7) % LAST.length];
      const roll = r();
      const status = roll < 0.03 ? 'Deleted' : roll < 0.05 ? 'Disabled' : 'Active';
      users[id] = {
        id,
        displayName: `${first} ${last}`,
        upn: `${first}.${last}${i}@contoso.com`.toLowerCase(),
        department: pick(DEPTS),
        country: pick(COUNTRIES),
        accountEnabled: status === 'Active',
        lastSignIn: status === 'Deleted' ? undefined : this._iso(Math.floor(Math.pow(r(), 3) * 200)),
        status
      };
      makerIds.push(id);
    }
    this._users = users;
    // A few prolific makers produce most of the resources, like in real tenants.
    const pickMaker = (): string => makerIds[Math.floor(Math.pow(r(), 2.2) * makerIds.length)];
    const envWeights = [40, 9, 6, 7, 6, 3, 3, 4, 3, 0, 0, 5, 4, 3];
    const totalW = envWeights.reduce((s, w) => s + w, 0);
    const pickEnv = (): IEnvironment => {
      let x = r() * totalW;
      for (let i = 0; i < envWeights.length; i++) {
        x -= envWeights[i];
        if (x < 0) return environments[i];
      }
      return environments[0];
    };
    const conns = (premiumChance: number): IConnectorRef[] => {
      const list: (IConnectorRef | undefined)[] = [];
      const n = 1 + Math.floor(r() * 3);
      for (let k = 0; k < n; k++) list.push(makeConnector(pick(STD)));
      if (r() < premiumChance) list.push(makeConnector(pick(PREM)));
      if (r() < 0.04) list.push(makeConnector(pick(CUSTOM), undefined, undefined, true));
      const seen: Record<string, boolean> = {};
      return list.filter((c): c is IConnectorRef => !!c && !seen[c.name] && (seen[c.name] = true));
    };

    const apps: IApp[] = [];
    for (let i = 0; i < 420; i++) {
      const env = pickEnv();
      const ownerId = pickMaker();
      const owner = users[ownerId];
      const created = createdDaysAgo();
      const modified = Math.max(0, created - Math.floor(r() * created));
      const model = env.hasDataverse && r() < 0.12;
      const c = model ? [] : conns(0.22);
      apps.push({
        id: `app-${i.toString(16).padStart(5, '0')}-${Math.floor(r() * 1e8).toString(16)}`,
        displayName: `${pick(APP_WORDS)} ${pick(APP_SUFFIX)}${r() < 0.15 ? ' v2' : ''}`,
        environmentId: env.id,
        environmentName: env.displayName,
        kind: model ? 'Model-driven' : 'Canvas',
        owner: { id: ownerId, displayName: owner.status === 'Deleted' ? undefined : owner.displayName, email: owner.upn },
        createdTime: this._iso(created),
        lastModifiedTime: this._iso(modified),
        lastPublishTime: this._iso(modified),
        connectors: c,
        sharedUsers: Math.floor(Math.pow(r(), 2) * 80),
        sharedGroups: Math.floor(r() * 3),
        sharedWithEveryone: r() < 0.035,
        isManaged: model ? r() < 0.3 : undefined,
        usesPremium: model || c.some((x) => x.tier === 'Premium'),
        usesCustom: c.some((x) => x.tier === 'Custom')
      });
    }

    const flows: IFlow[] = [];
    const triggers = ['Automated', 'Automated', 'Scheduled', 'Instant', 'Automated', 'Power Apps', 'HTTP request'];
    for (let i = 0; i < 900; i++) {
      const env = pickEnv();
      const ownerId = pickMaker();
      const created = createdDaysAgo();
      const roll = r();
      const state: FlowState = roll < 0.03 ? 'Suspended' : roll < 0.22 ? 'Off' : 'On';
      const c = conns(0.18);
      flows.push({
        id: `flow-${i.toString(16).padStart(5, '0')}-${Math.floor(r() * 1e8).toString(16)}`,
        displayName: `${pick(FLOW_VERBS)} ${pick(FLOW_OBJECTS)}`,
        environmentId: env.id,
        environmentName: env.displayName,
        state,
        suspensionReason: state === 'Suspended' ? pick(['DlpPolicyViolation', 'LicenseExpired', 'ThrottlingLimitExceeded']) : undefined,
        owner: { id: ownerId },
        createdTime: this._iso(created),
        lastModifiedTime: this._iso(Math.max(0, created - Math.floor(r() * created))),
        triggerType: pick(triggers),
        connectors: c,
        usesPremium: c.some((x) => x.tier === 'Premium'),
        usesCustom: c.some((x) => x.tier === 'Custom'),
        actionCount: 2 + Math.floor(r() * 30)
      });
    }

    const agents: IAgent[] = [];
    const agentEnvs = environments.filter((e) => e.hasDataverse && e.type !== 'Trial' && e.displayName !== 'Dev - unused');
    const agentNames = ['IT Helpdesk', 'HR Policy', 'Expense Assistant', 'Sales Coach', 'Onboarding Buddy', 'Facilities', 'Procurement FAQ', 'Travel Desk', 'Legal Intake', 'Customer Returns'];
    for (let i = 0; i < 46; i++) {
      const env = agentEnvs[i % 3 === 0 ? agentEnvs.length - 1 : Math.floor(r() * agentEnvs.length)];
      const ownerId = pickMaker();
      const owner = users[ownerId];
      const created = Math.floor(Math.pow(r(), 1.2) * 300);
      const published = r() < 0.7;
      const authRoll = r();
      const auth: AgentAuth = authRoll < 0.12 ? 'None' : authRoll < 0.85 ? 'Microsoft' : 'Manual';
      agents.push({
        id: `bot-${i.toString(16).padStart(4, '0')}-${Math.floor(r() * 1e8).toString(16)}`,
        displayName: `${pick(agentNames)}${r() < 0.3 ? ' (pilot)' : ''}`,
        environmentId: env.id,
        environmentName: env.displayName,
        owner: { id: ownerId, displayName: owner.status === 'Deleted' ? undefined : owner.displayName, email: owner.upn },
        createdTime: this._iso(created),
        modifiedTime: this._iso(Math.max(0, created - Math.floor(r() * created))),
        publishedTime: published ? this._iso(Math.max(0, created - Math.floor(r() * created))) : undefined,
        authMode: auth,
        language: '1033',
        state: r() < 0.9 ? 'Active' : 'Inactive'
      });
    }

    this._inventory = { environments, apps, flows, agents, errors: [], loadedAt: new Date(this._now).toISOString() };
    return this._inventory;
  }

  public getInventory(): Promise<IInventory> {
    return this._cached('inventory', () => this._delay(this._build()));
  }

  public getMakers(): Promise<IMakersData> {
    return this._cached('makers', async () => {
      const inv = await this.getInventory();
      return { makers: buildMakers(inv, this._users), users: this._users, signInUnavailable: false };
    });
  }

  public getDlpPolicies(): Promise<IDlpPolicy[]> {
    return this._cached('dlp', async () => {
      const inv = await this.getInventory();
      const id = (name: string): string => (inv.environments.filter((e) => e.displayName === name)[0] || { id: '' }).id;
      return this._delay<IDlpPolicy[]>([
        {
          id: 'pol-1', displayName: 'Tenant baseline', scope: 'All except', environmentIds: [id('Operations-Prod'), id('Copilot Agents'), id('Sales-UAT')],
          businessConnectors: ['shared_sharepointonline', 'shared_office365', 'shared_teams', 'shared_approvals', 'shared_commondataserviceforapps'],
          nonBusinessConnectors: ['shared_twitter', 'shared_dropbox', 'shared_gmail'], blockedConnectors: ['shared_http', 'shared_ftp'],
          defaultGroup: 'Non-business', createdBy: 'Grace Hughes', lastModified: this._iso(41)
        },
        {
          id: 'pol-2', displayName: 'Finance strict', scope: 'Only selected', environmentIds: [id('Finance-Prod'), id('Finance-Test')],
          businessConnectors: ['shared_sharepointonline', 'shared_sql', 'shared_commondataserviceforapps', 'shared_office365'],
          nonBusinessConnectors: [], blockedConnectors: ['shared_twitter', 'shared_dropbox', 'shared_gmail', 'shared_googledrive', 'shared_http'],
          defaultGroup: 'Blocked', createdBy: 'Victor Rossi', lastModified: this._iso(120)
        },
        {
          id: 'pol-3', displayName: 'Agents environment', scope: 'Only selected', environmentIds: [id('Copilot Agents')],
          businessConnectors: ['shared_sharepointonline', 'shared_commondataserviceforapps', 'shared_azureopenai'],
          nonBusinessConnectors: [], blockedConnectors: ['shared_http'], defaultGroup: 'Blocked', createdBy: 'Grace Hughes', lastModified: this._iso(9)
        }
      ]);
    });
  }

  public getSkus(): Promise<ISku[]> {
    return this._cached('skus', () =>
      this._delay<ISku[]>([
        { skuId: 's1', skuPartNumber: 'SPE_E3', friendlyName: 'Microsoft 365 E3', enabled: 1200, consumed: 1131, suspended: 0, warning: 0, isPowerPlatform: false, appliesTo: 'User' },
        { skuId: 's2', skuPartNumber: 'SPE_E5', friendlyName: 'Microsoft 365 E5', enabled: 150, consumed: 142, suspended: 0, warning: 0, isPowerPlatform: false, appliesTo: 'User' },
        { skuId: 's3', skuPartNumber: 'Microsoft_365_Copilot', friendlyName: 'Microsoft 365 Copilot', enabled: 300, consumed: 287, suspended: 0, warning: 0, isPowerPlatform: false, appliesTo: 'User' },
        { skuId: 's4', skuPartNumber: 'POWERAPPS_PER_USER', friendlyName: 'Power Apps Premium', enabled: 120, consumed: 96, suspended: 0, warning: 0, isPowerPlatform: true, appliesTo: 'User' },
        { skuId: 's5', skuPartNumber: 'POWER_AUTOMATE_PREMIUM', friendlyName: 'Power Automate Premium', enabled: 60, consumed: 58, suspended: 0, warning: 0, isPowerPlatform: true, appliesTo: 'User' },
        { skuId: 's6', skuPartNumber: 'FLOW_PER_PROCESS', friendlyName: 'Power Automate Process', enabled: 5, consumed: 3, suspended: 0, warning: 0, isPowerPlatform: true, appliesTo: 'Company' },
        { skuId: 's7', skuPartNumber: 'Microsoft_Copilot_Studio', friendlyName: 'Copilot Studio', enabled: 2, consumed: 0, suspended: 0, warning: 0, isPowerPlatform: true, appliesTo: 'Company' },
        { skuId: 's8', skuPartNumber: 'POWERAPPS_VIRAL', friendlyName: 'Power Apps (trial / viral)', enabled: 10000, consumed: 212, suspended: 0, warning: 0, isPowerPlatform: true, appliesTo: 'User' },
        { skuId: 's9', skuPartNumber: 'FLOW_FREE', friendlyName: 'Power Automate Free', enabled: 10000, consumed: 341, suspended: 0, warning: 0, isPowerPlatform: true, appliesTo: 'User' },
        { skuId: 's10', skuPartNumber: 'CDS_DB_CAPACITY', friendlyName: 'Dataverse database capacity', enabled: 20, consumed: 0, suspended: 0, warning: 0, isPowerPlatform: true, appliesTo: 'Company' },
        { skuId: 's11', skuPartNumber: 'POWER_BI_PRO', friendlyName: 'Power BI Pro', enabled: 400, consumed: 371, suspended: 0, warning: 0, isPowerPlatform: false, appliesTo: 'User' },
        { skuId: 's12', skuPartNumber: 'AAD_PREMIUM_P2', friendlyName: 'Microsoft Entra ID P2', enabled: 150, consumed: 150, suspended: 0, warning: 10, isPowerPlatform: false, appliesTo: 'User' }
      ])
    );
  }

  public getServiceHealth(): Promise<IServiceHealth> {
    return this._cached('health', () => {
      const services = ['Exchange Online', 'Microsoft Teams', 'SharePoint Online', 'OneDrive for Business', 'Power Apps', 'Power Automate', 'Microsoft Copilot Studio', 'Dynamics 365 Apps', 'Microsoft Entra', 'Power BI'];
      return this._delay<IServiceHealth>({
        overviews: services.map((s) => ({ service: s, status: s === 'Power Automate' ? 'serviceDegradation' : s === 'Microsoft Teams' ? 'serviceRestored' : 'serviceOperational' })),
        issues: [
          { id: 'MF812345', title: 'Some users may see delays when flows with SharePoint triggers run', service: 'Power Automate', status: 'serviceDegradation', classification: 'advisory', startDateTime: this._iso(1) },
          { id: 'TM809876', title: 'Users may be unable to join scheduled meetings from Outlook', service: 'Microsoft Teams', status: 'serviceRestored', classification: 'incident', startDateTime: this._iso(2) }
        ],
        messages: [
          { id: 'MC1101', title: 'Managed Environments: sharing limits will apply to agents', services: ['Microsoft Copilot Studio'], category: 'planForChange', lastModified: this._iso(3), actionRequiredBy: this._iso(-30), isMajorChange: true },
          { id: 'MC1098', title: 'Power Apps: classic designer retirement', services: ['Power Apps'], category: 'planForChange', lastModified: this._iso(6), isMajorChange: true },
          { id: 'MC1093', title: 'Power Automate: new run history experience', services: ['Power Automate'], category: 'stayInformed', lastModified: this._iso(9), isMajorChange: false },
          { id: 'MC1090', title: 'SharePoint: updated site lifecycle management policies', services: ['SharePoint Online'], category: 'stayInformed', lastModified: this._iso(12), isMajorChange: false },
          { id: 'MC1087', title: 'Dataverse: storage capacity reports move to the new admin center', services: ['Dynamics 365 Apps'], category: 'stayInformed', lastModified: this._iso(15), isMajorChange: false }
        ]
      });
    });
  }

  public getReportsConcealed(): Promise<boolean> {
    return this._delay(false);
  }

  public getSites(): Promise<ISitesData> {
    return this._cached('sites', () => {
      const r = rng(7);
      const pick = <T>(arr: T[]): T => arr[Math.floor(r() * arr.length)];
      const templates = ['Group', 'Group', 'Team Site', 'Communication Site', 'Group', 'Team Channel', 'Classic Team Site'];
      const names = ['finance', 'hr-hub', 'sales-emea', 'project-atlas', 'it-service-desk', 'legal', 'marketing-brand', 'ops-north', 'intranet', 'procurement', 'board-papers', 'customer-success', 'engineering', 'policies', 'training'];
      const sites: ISite[] = [];
      for (let i = 0; i < 260; i++) {
        const name = `${pick(names)}${i > 14 ? '-' + i : ''}`;
        const used = Math.floor(Math.pow(r(), 3) * 900 * 1024 * 1024 * 1024);
        const files = Math.floor(used / (1.6 * 1024 * 1024)) + Math.floor(r() * 300);
        const inactive = r() < 0.22;
        sites.push({
          id: `site-${i}`,
          url: `https://contoso.sharepoint.com/sites/${name}`,
          title: name,
          ownerName: `${pick(FIRST)} ${pick(LAST)}`,
          ownerUpn: '',
          lastActivity: inactive ? (r() < 0.3 ? undefined : this._iso(120 + Math.floor(r() * 500))) : this._iso(Math.floor(r() * 30)),
          template: pick(templates),
          storageUsedBytes: used,
          storageAllocatedBytes: 25 * 1024 * 1024 * 1024 * 1024,
          fileCount: files,
          activeFileCount: inactive ? 0 : Math.floor(files * r() * 0.2),
          pageViews: inactive ? 0 : Math.floor(Math.pow(r(), 2) * 6000),
          externalSharing: r() < 0.3,
          anonymousLinks: r() < 0.1 ? Math.floor(r() * 40) : 0,
          isDeleted: false
        });
      }
      const total = sites.reduce((s, x) => s + x.storageUsedBytes, 0);
      const storageTrend = [];
      for (let d = 89; d >= 0; d--) storageTrend.push({ date: this._iso(d).slice(0, 10), bytes: Math.round(total * (1 - d * 0.0011) + Math.sin(d) * total * 0.0015) });
      return this._delay<ISitesData>({ sites, storageTrend, sharing: { sharingCapability: 'externalUserAndGuestSharing', isResharingByExternalUsersEnabled: false, sharingDomainRestrictionMode: 'none' } });
    });
  }

  public getTeams(): Promise<ITeamsData> {
    return this._cached('teams', () => {
      const r = rng(11);
      const pick = <T>(arr: T[]): T => arr[Math.floor(r() * arr.length)];
      const groups: IGroup[] = [];
      const activity: ITeamActivity[] = [];
      for (let i = 0; i < 180; i++) {
        const isTeam = r() < 0.72;
        const name = `${pick(DEPTS)} ${pick(['Team', 'Project', 'Leadership', 'Community', 'Working Group', 'Programme'])} ${i}`;
        const owners = r() < 0.08 ? 0 : r() < 0.25 ? 1 : 2 + Math.floor(r() * 3);
        const id = `grp-${i}`;
        groups.push({ id, displayName: name, mail: `${name.replace(/\s+/g, '').toLowerCase()}@contoso.com`, createdTime: this._iso(Math.floor(Math.pow(r(), 1.5) * 900)), visibility: r() < 0.7 ? 'Private' : 'Public', isTeam, ownerCount: owners });
        if (isTeam) {
          const quiet = r() < 0.25;
          activity.push({
            id, name, teamType: r() < 0.7 ? 'Private' : 'Public',
            lastActivity: quiet ? (r() < 0.4 ? undefined : this._iso(60 + Math.floor(r() * 300))) : this._iso(Math.floor(r() * 20)),
            activeUsers: quiet ? 0 : 2 + Math.floor(r() * 60), guests: r() < 0.25 ? 1 + Math.floor(r() * 8) : 0,
            activeExternalUsers: r() < 0.15 ? 1 + Math.floor(r() * 4) : 0, activeChannels: quiet ? 0 : 1 + Math.floor(r() * 9),
            messages: quiet ? 0 : Math.floor(Math.pow(r(), 2) * 1500), meetings: quiet ? 0 : Math.floor(r() * 40)
          });
        }
      }
      return this._delay<ITeamsData>({ groups, activity });
    });
  }

  public getAdoption(): Promise<IAdoptionData> {
    return this._cached('adoption', () => {
      const r = rng(5);
      const counts = [];
      for (let d = 89; d >= 0; d--) {
        const weekday = new Date(this._now - d * DAY).getDay();
        const f = weekday === 0 || weekday === 6 ? 0.35 : 1;
        const trend = 1 + (89 - d) * 0.0012;
        counts.push({
          date: this._iso(d).slice(0, 10),
          office365: Math.round(1180 * f * trend + r() * 30),
          exchange: Math.round(1120 * f * trend + r() * 30),
          teams: Math.round(1050 * f * trend + r() * 30),
          sharePoint: Math.round(760 * f * trend + r() * 30),
          oneDrive: Math.round(820 * f * trend + r() * 30)
        });
      }
      const users: IUserActivity[] = [];
      for (let i = 0; i < 1270; i++) {
        const inactive = r() < 0.08;
        users.push({
          upn: `user${i}@contoso.com`,
          displayName: `${FIRST[i % FIRST.length]} ${LAST[(i * 3) % LAST.length]}`,
          products: i < 142 ? ['MICROSOFT 365 E5'] : ['MICROSOFT 365 E3'].concat(i % 4 === 0 ? ['MICROSOFT 365 COPILOT'] : []),
          lastActivity: inactive ? (r() < 0.5 ? undefined : this._iso(31 + Math.floor(r() * 200))) : this._iso(Math.floor(r() * 10)),
          isDeleted: false
        });
      }
      const copilotUsers = users.filter((u) => u.products.indexOf('MICROSOFT 365 COPILOT') >= 0).map((u) => ({
        upn: u.upn, displayName: u.displayName, lastActivity: r() < 0.72 ? this._iso(Math.floor(r() * 25)) : undefined
      }));
      return this._delay<IAdoptionData>({ counts, users, copilotUsers, copilotUnavailable: false });
    });
  }

  public getIdentity(): Promise<IIdentityData> {
    return this._cached('identity', () => {
      const r = rng(19);
      const users: IUserAccount[] = [];
      const mfa: IMfaRegistration[] = [];
      for (let i = 0; i < 1460; i++) {
        const guest = i >= 1300;
        const name = `${FIRST[i % FIRST.length]} ${LAST[(i * 7) % LAST.length]}`;
        const upn = guest ? `${name.replace(' ', '.').toLowerCase()}_partner.com#EXT#@contoso.onmicrosoft.com` : `${name.replace(' ', '.').toLowerCase()}${i}@contoso.com`;
        const stale = r() < (guest ? 0.45 : 0.06);
        users.push({
          id: `id-${i}`, displayName: name, upn, userType: guest ? 'Guest' : 'Member', accountEnabled: r() > 0.03,
          createdTime: this._iso(Math.floor(r() * 1500)),
          lastSignIn: stale ? (r() < 0.4 ? undefined : this._iso(91 + Math.floor(r() * 400))) : this._iso(Math.floor(r() * 20))
        });
        if (!guest) {
          const reg = r() < 0.91;
          mfa.push({ id: `id-${i}`, upn, displayName: name, isMfaRegistered: reg, isAdmin: i < 14, userType: 'member', methods: reg ? (r() < 0.7 ? ['microsoftAuthenticatorPush'] : ['mobilePhone']) : [] });
        }
      }
      const member = (i: number): { id: string; displayName: string; upn: string; type: string } => ({ id: users[i].id, displayName: users[i].displayName, upn: users[i].upn, type: 'user' });
      const roles: IDirectoryRole[] = [
        { id: 'r1', displayName: 'Global Administrator', members: [member(0), member(1), member(2), member(3), member(4)] },
        { id: 'r2', displayName: 'Power Platform Administrator', members: [member(5), member(6)] },
        { id: 'r3', displayName: 'SharePoint Administrator', members: [member(7), member(8), member(9)] },
        { id: 'r4', displayName: 'Dynamics 365 Administrator', members: [member(6)] },
        { id: 'r5', displayName: 'Global Reader', members: [member(10), member(11), member(12)] },
        { id: 'r6', displayName: 'Teams Administrator', members: [member(13)] },
        { id: 'r7', displayName: 'Reports Reader', members: [member(10)] }
      ];
      const credentials: IAppCredential[] = [];
      const apps = ['Contoso ERP connector', 'HR API', 'Power BI gateway', 'DocuSign integration', 'Backup service', 'SharePoint migration', 'Ticketing bridge', 'Data export job'];
      apps.forEach((a, i) => {
        const days = [-40, 6, 19, 75, 210, 340, 28, 400][i];
        credentials.push({ appId: `app-reg-${i}`, appDisplayName: a, type: i % 3 === 0 ? 'Certificate' : 'Secret', displayName: i % 3 === 0 ? 'CN=' + a : 'prod', endDateTime: this._iso(-days), daysLeft: days });
      });
      return this._delay<IIdentityData>({ roles, users, signInUnavailable: false, mfa, mfaUnavailable: false, credentials });
    });
  }

  public async getAppCredentials(): Promise<IAppCredential[]> {
    return (await this.getIdentity()).credentials;
  }

  public getAppSharing(app: IApp): Promise<IAppSharing> {
    const principals = [{ name: app.owner.displayName || 'Owner', type: 'User', role: 'Owner' }];
    if (app.sharedWithEveryone) principals.push({ name: 'Everyone in the organisation', type: 'Tenant', role: 'CanView' });
    for (let i = 0; i < Math.min(app.sharedGroups || 0, 3); i++) principals.push({ name: `${DEPTS[i]} team`, type: 'Group', role: 'CanView' });
    return this._delay({ users: app.sharedUsers || 0, groups: app.sharedGroups || 0, everyone: !!app.sharedWithEveryone, principals });
  }

  public getFlowRuns(flow: IFlow): Promise<IFlowRun[]> {
    const r = rng(flow.id.length * 31 + flow.actionCount);
    const runs: IFlowRun[] = [];
    for (let i = 0; i < 30; i++) {
      const start = this._now - i * DAY * (0.4 + r());
      runs.push({ status: flow.state === 'Suspended' && i < 3 ? 'Failed' : r() < 0.08 ? 'Failed' : 'Succeeded', startTime: new Date(start).toISOString(), endTime: new Date(start + 40000 * r()).toISOString() });
    }
    return this._delay(runs);
  }

  public async scanAppSharing(apps: IApp[], onProgress: (p: IProgress) => void): Promise<number> {
    const canvas = apps.filter((a) => a.kind === 'Canvas');
    for (let i = 0; i < canvas.length; i += 60) {
      await this._delay(undefined);
      onProgress({ label: 'Checking app sharing', done: Math.min(canvas.length, i + 60), total: canvas.length });
    }
    return canvas.filter((a) => a.sharedWithEveryone).length;
  }
}
