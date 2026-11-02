import { IAgent, IApp, IAppCredential, IDlpPolicy, IEnvironment, IFinding, IFlow, IGroup, IUserInfo, Severity } from '../models';
import { daysSince } from './format';

/* ---------- Deep links into the Microsoft admin and maker portals ---------- */

export const links = {
  environment: (envId: string): string => `https://admin.powerplatform.microsoft.com/environments/${envId}/hub`,
  app: (envId: string, appId: string): string => `https://make.powerapps.com/environments/${envId}/apps/${appId}/details`,
  modelApp: (envId: string): string => `https://make.powerapps.com/environments/${envId}/apps`,
  flow: (envId: string, flowId: string): string => `https://make.powerautomate.com/environments/${envId}/flows/${flowId}/details`,
  agent: (envId: string, botId: string): string => `https://copilotstudio.microsoft.com/environments/${envId}/bots/${botId}/overview`,
  dlp: (): string => 'https://admin.powerplatform.microsoft.com/security/dataprotection',
  group: (id: string): string => `https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Overview/groupId/${id}`,
  appRegistration: (appId: string): string =>
    `https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Credentials/appId/${appId}`
};

export interface IGovernanceInput {
  environments: IEnvironment[];
  apps: IApp[];
  flows: IFlow[];
  agents: IAgent[];
  users: Record<string, IUserInfo>;
  dlp?: IDlpPolicy[];
  groups?: IGroup[];
  credentials?: IAppCredential[];
  staleDays?: number;
}

export interface IRuleInfo {
  id: string;
  severity: Severity;
  category: string;
  title: string;
  action: string;
}

/** Every rule the engine evaluates, with the recommended action shown in the UI. */
export const RULES: IRuleInfo[] = [
  { id: 'orphan-app', severity: 'High', category: 'Ownership', title: 'App owner has left or is disabled', action: 'Assign a new owner, or archive the app.' },
  { id: 'orphan-flow', severity: 'High', category: 'Ownership', title: 'Flow owner has left or is disabled', action: 'Add a co-owner or service account before the flow stops.' },
  { id: 'orphan-agent', severity: 'High', category: 'Ownership', title: 'Agent owner has left or is disabled', action: 'Reassign ownership in Copilot Studio.' },
  { id: 'app-everyone', severity: 'High', category: 'Sharing', title: 'App shared with everyone', action: 'Confirm the app is meant for the whole organisation, or restrict it to a group.' },
  { id: 'agent-noauth', severity: 'High', category: 'Security', title: 'Published agent without authentication', action: 'Turn on Microsoft authentication unless the agent is intentionally public.' },
  { id: 'env-nodlp', severity: 'High', category: 'Data loss prevention', title: 'Environment not covered by any DLP policy', action: 'Add the environment to a tenant DLP policy.' },
  { id: 'env-nosg', severity: 'Medium', category: 'Environments', title: 'Production environment without a security group', action: 'Restrict access with an Entra security group.' },
  { id: 'flow-suspended', severity: 'Medium', category: 'Reliability', title: 'Flow suspended', action: 'Check the suspension reason (licence, DLP or throttling) and fix it.' },
  { id: 'default-premium', severity: 'Medium', category: 'Environments', title: 'Premium or custom connectors in the default environment', action: 'Move business-critical solutions to a dedicated environment.' },
  { id: 'group-ownerless', severity: 'Medium', category: 'Microsoft 365', title: 'Microsoft 365 group or team without an owner', action: 'Assign at least two owners.' },
  { id: 'cred-expiring', severity: 'Medium', category: 'Security', title: 'App registration secret or certificate expiring', action: 'Rotate the credential before it expires and update connectors that use it.' },
  { id: 'env-unmanaged', severity: 'Low', category: 'Environments', title: 'Production environment is not a Managed Environment', action: 'Enable Managed Environments for sharing limits, usage insights and solution checker.' },
  { id: 'app-stale', severity: 'Low', category: 'Lifecycle', title: 'App not modified for a long time', action: 'Confirm the app is still used; archive or delete it if not.' },
  { id: 'env-empty', severity: 'Low', category: 'Lifecycle', title: 'Developer or trial environment with no resources', action: 'Delete unused environments to free capacity.' }
];

const RULE_MAP: Record<string, IRuleInfo> = {};
RULES.forEach((r) => {
  RULE_MAP[r.id] = r;
});

function ownerLeft(id: string | undefined, users: Record<string, IUserInfo>): IUserInfo | undefined {
  if (!id) return undefined;
  const u = users[id];
  return u && (u.status === 'Deleted' || u.status === 'Disabled') ? u : undefined;
}

function ownerName(ref: { id?: string; displayName?: string; email?: string }, users: Record<string, IUserInfo>): string {
  const u = ref.id ? users[ref.id] : undefined;
  return ref.displayName || (u && (u.displayName || u.upn)) || ref.email || (u && u.status === 'Deleted' ? 'Deleted user' : ref.id || 'Unknown');
}

/** True when a DLP policy applies to the environment. */
export function isCoveredByDlp(env: IEnvironment, policies: IDlpPolicy[]): boolean {
  return policies.some((p) => {
    if (p.scope === 'All environments') return true;
    const listed = p.environmentIds.indexOf(env.id) >= 0;
    return p.scope === 'All except' ? !listed : listed;
  });
}

export function evaluateGovernance(input: IGovernanceInput): IFinding[] {
  const out: IFinding[] = [];
  const staleDays = input.staleDays || 180;
  const add = (ruleId: string, resourceType: string, resourceName: string, environmentName: string, owner: string, detail: string, link?: string): void => {
    const rule = RULE_MAP[ruleId];
    out.push({
      id: `${ruleId}:${resourceType}:${resourceName}:${environmentName}:${out.length}`,
      ruleId,
      severity: rule.severity,
      category: rule.category,
      title: rule.title,
      resourceType,
      resourceName,
      environmentName,
      owner,
      detail,
      link
    });
  };

  input.apps.forEach((a) => {
    const left = ownerLeft(a.owner.id, input.users);
    if (left) add('orphan-app', 'App', a.displayName, a.environmentName, ownerName(a.owner, input.users), `Owner is ${left.status.toLowerCase()}.`, a.kind === 'Canvas' ? links.app(a.environmentId, a.id) : links.modelApp(a.environmentId));
    if (a.sharedWithEveryone) add('app-everyone', 'App', a.displayName, a.environmentName, ownerName(a.owner, input.users), 'Shared with "Everyone in the organisation".', links.app(a.environmentId, a.id));
    const age = daysSince(a.lastModifiedTime);
    if (a.kind === 'Canvas' && age !== undefined && age > staleDays) add('app-stale', 'App', a.displayName, a.environmentName, ownerName(a.owner, input.users), `Last modified ${age} days ago.`, links.app(a.environmentId, a.id));
  });

  input.flows.forEach((f) => {
    const left = ownerLeft(f.owner.id, input.users);
    if (left && f.state !== 'Off') add('orphan-flow', 'Flow', f.displayName, f.environmentName, ownerName(f.owner, input.users), `Owner is ${left.status.toLowerCase()}; flow is ${f.state.toLowerCase()}.`, links.flow(f.environmentId, f.id));
    if (f.state === 'Suspended') add('flow-suspended', 'Flow', f.displayName, f.environmentName, ownerName(f.owner, input.users), f.suspensionReason ? `Reason: ${f.suspensionReason}.` : 'Suspended by the platform.', links.flow(f.environmentId, f.id));
  });

  input.agents.forEach((b) => {
    const left = ownerLeft(b.owner.id, input.users);
    if (left) add('orphan-agent', 'Agent', b.displayName, b.environmentName, ownerName(b.owner, input.users), `Owner is ${left.status.toLowerCase()}.`, links.agent(b.environmentId, b.id));
    if (b.authMode === 'None' && b.publishedTime) add('agent-noauth', 'Agent', b.displayName, b.environmentName, ownerName(b.owner, input.users), 'Anyone with the link can chat with this agent.', links.agent(b.environmentId, b.id));
  });

  const appsByEnv: Record<string, number> = {};
  const flowsByEnv: Record<string, number> = {};
  const premiumDefault: Record<string, number> = {};
  input.apps.forEach((a) => {
    appsByEnv[a.environmentId] = (appsByEnv[a.environmentId] || 0) + 1;
    if (a.usesPremium || a.usesCustom) premiumDefault[a.environmentId] = (premiumDefault[a.environmentId] || 0) + 1;
  });
  input.flows.forEach((f) => {
    flowsByEnv[f.environmentId] = (flowsByEnv[f.environmentId] || 0) + 1;
    if (f.usesPremium || f.usesCustom) premiumDefault[f.environmentId] = (premiumDefault[f.environmentId] || 0) + 1;
  });
  const agentsByEnv: Record<string, number> = {};
  input.agents.forEach((b) => {
    agentsByEnv[b.environmentId] = (agentsByEnv[b.environmentId] || 0) + 1;
  });

  input.environments.forEach((e) => {
    const type = e.type.toLowerCase();
    if (input.dlp && !isCoveredByDlp(e, input.dlp)) add('env-nodlp', 'Environment', e.displayName, e.displayName, e.createdBy || '', 'No DLP policy applies, so any connector can be combined.', links.environment(e.id));
    if (type === 'production' && !e.isDefault && !e.securityGroupId) add('env-nosg', 'Environment', e.displayName, e.displayName, e.createdBy || '', 'Every licensed user in the tenant can be added to this environment.', links.environment(e.id));
    if (type === 'production' && !e.isManaged) add('env-unmanaged', 'Environment', e.displayName, e.displayName, e.createdBy || '', 'Managed Environment features are off.', links.environment(e.id));
    if (e.isDefault && premiumDefault[e.id]) add('default-premium', 'Environment', e.displayName, e.displayName, '', `${premiumDefault[e.id]} apps and flows use premium or custom connectors.`, links.environment(e.id));
    const total = (appsByEnv[e.id] || 0) + (flowsByEnv[e.id] || 0) + (agentsByEnv[e.id] || 0);
    if ((type === 'developer' || type === 'trial') && total === 0) add('env-empty', 'Environment', e.displayName, e.displayName, e.createdBy || '', 'No apps, flows or agents.', links.environment(e.id));
  });

  (input.groups || []).forEach((g) => {
    if (g.ownerCount === 0) add('group-ownerless', g.isTeam ? 'Team' : 'Group', g.displayName, '—', '', 'No owners.', links.group(g.id));
  });

  (input.credentials || []).forEach((c) => {
    if (c.daysLeft <= 30 && c.daysLeft >= -90) add('cred-expiring', 'App registration', c.appDisplayName, '—', '', c.daysLeft < 0 ? `${c.type} "${c.displayName}" expired ${-c.daysLeft} days ago.` : `${c.type} "${c.displayName}" expires in ${c.daysLeft} days.`, links.appRegistration(c.appId));
  });

  const order: Record<Severity, number> = { High: 0, Medium: 1, Low: 2 };
  return out.sort((a, b) => order[a.severity] - order[b.severity] || (a.title < b.title ? -1 : a.title > b.title ? 1 : 0));
}

export function ruleAction(ruleId: string): string {
  return RULE_MAP[ruleId] ? RULE_MAP[ruleId].action : '';
}
