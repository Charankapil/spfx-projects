import {
  IAdoptionData, IAppCredential, IAgent, IApp, IDlpPolicy, IEnvironment, IFinding, IFlow, IIdentityData, IInventory, ISitesData, ISku, ITeamsData
} from '../models';
import { IMakersData } from '../providers/IDataProvider';
import { IConnectorUsage, IEnvRow, ISkuRow } from './columns';
import { isCoveredByDlp, evaluateGovernance } from '../logic/governance';
import { daysSince, fmtBytes, fmtMb, fmtNumber, fmtPercent } from '../logic/format';
import { sumBy } from '../logic/aggregate';
import { Tone } from './components/Layout';

export interface ISummaryItem {
  label: string;
  value: string;
  sub?: string;
  tone?: Tone;
}

export interface IPPData {
  inv: IInventory;
  makers: IMakersData;
}

/**
 * Applies the environment filter. Managed model-driven apps (installed by Microsoft
 * or ISV solutions, e.g. "Power Platform Environment Settings") are never counted,
 * so every page and export uses the same app totals.
 */
export function filterInventory(inv: IInventory, envFilter: string[]): IInventory {
  const has = (id: string): boolean => !envFilter.length || envFilter.indexOf(id) >= 0;
  return {
    ...inv,
    environments: inv.environments.filter((e) => has(e.id)),
    apps: inv.apps.filter((a) => !a.isManaged && has(a.environmentId)),
    flows: inv.flows.filter((f) => has(f.environmentId)),
    agents: inv.agents.filter((b) => has(b.environmentId))
  };
}

/* ---------- Environments ---------- */

export function envRows(inv: IInventory, dlp: IDlpPolicy[] | undefined): IEnvRow[] {
  const count = (list: { environmentId: string }[]): Record<string, number> => {
    const m: Record<string, number> = {};
    list.forEach((x) => {
      m[x.environmentId] = (m[x.environmentId] || 0) + 1;
    });
    return m;
  };
  const apps = count(inv.apps);
  const flows = count(inv.flows);
  const agents = count(inv.agents);
  const makers: Record<string, Record<string, boolean>> = {};
  const addMaker = (envId: string, id?: string): void => {
    if (!id) return;
    (makers[envId] = makers[envId] || {})[id] = true;
  };
  inv.apps.forEach((a) => addMaker(a.environmentId, a.owner.id));
  inv.flows.forEach((f) => addMaker(f.environmentId, f.owner.id));
  inv.agents.forEach((b) => addMaker(b.environmentId, b.owner.id));
  return inv.environments.map((e) => ({
    ...e,
    apps: apps[e.id] || 0,
    flows: flows[e.id] || 0,
    agents: agents[e.id] || 0,
    makers: Object.keys(makers[e.id] || {}).length,
    dlpCovered: dlp ? isCoveredByDlp(e, dlp) : undefined
  }));
}

export function envSummary(rows: IEnvRow[], dlpKnown: boolean): ISummaryItem[] {
  const prod = rows.filter((e) => e.type === 'Production');
  const uncovered = rows.filter((e) => e.dlpCovered === false).length;
  const db = sumBy(rows, (e) => e.capacity.databaseMb);
  return [
    { label: 'Environments', value: fmtNumber(rows.length), sub: `${prod.length} production` },
    { label: 'With Dataverse', value: fmtNumber(rows.filter((e) => e.hasDataverse).length) },
    { label: 'Managed', value: fmtNumber(rows.filter((e) => e.isManaged).length), sub: fmtPercent(rows.filter((e) => e.isManaged).length, rows.length) + ' of all' },
    {
      label: 'Without DLP', value: dlpKnown ? fmtNumber(uncovered) : '—', tone: dlpKnown ? (uncovered ? 'critical' : 'good') : undefined,
      sub: dlpKnown ? (uncovered ? 'Not covered by any policy' : 'All covered') : 'DLP not readable'
    },
    { label: 'Database used', value: fmtMb(db), sub: `File ${fmtMb(sumBy(rows, (e) => e.capacity.fileMb))}` }
  ];
}

/* ---------- Apps, flows, agents ---------- */

function isLeft(id: string | undefined, data: IMakersData): boolean {
  const u = id ? data.users[id] : undefined;
  return !!u && (u.status === 'Deleted' || u.status === 'Disabled');
}

export function appSummary(apps: IApp[], makers: IMakersData): ISummaryItem[] {
  const canvas = apps.filter((a) => a.kind === 'Canvas');
  const orphans = apps.filter((a) => isLeft(a.owner.id, makers)).length;
  const everyoneKnown = apps.some((a) => a.sharedWithEveryone !== undefined);
  const everyone = apps.filter((a) => a.sharedWithEveryone).length;
  const stale = canvas.filter((a) => (daysSince(a.lastModifiedTime) || 0) > 180).length;
  return [
    { label: 'Apps', value: fmtNumber(apps.length), sub: `${fmtNumber(canvas.length)} canvas · ${fmtNumber(apps.length - canvas.length)} model-driven` },
    { label: 'Premium', value: fmtNumber(apps.filter((a) => a.usesPremium).length), sub: fmtPercent(apps.filter((a) => a.usesPremium).length, apps.length) + ' of apps' },
    { label: 'Orphaned', value: fmtNumber(orphans), tone: orphans ? 'critical' : 'good', sub: orphans ? 'Owner left or disabled' : 'All owners active' },
    { label: 'Shared with everyone', value: everyoneKnown ? fmtNumber(everyone) : '—', tone: everyoneKnown ? (everyone ? 'warning' : 'good') : undefined, sub: everyoneKnown ? undefined : 'Run a sharing scan' },
    { label: 'Not modified 180d', value: fmtNumber(stale), sub: 'Archive candidates' }
  ];
}

export function flowSummary(flows: IFlow[], makers: IMakersData): ISummaryItem[] {
  const on = flows.filter((f) => f.state === 'On').length;
  const suspended = flows.filter((f) => f.state === 'Suspended').length;
  const orphans = flows.filter((f) => f.state !== 'Off' && isLeft(f.owner.id, makers)).length;
  return [
    { label: 'Cloud flows', value: fmtNumber(flows.length), sub: `${fmtNumber(on)} on` },
    { label: 'Suspended', value: fmtNumber(suspended), tone: suspended ? 'critical' : 'good', sub: suspended ? 'Need attention' : 'None' },
    { label: 'Turned off', value: fmtNumber(flows.filter((f) => f.state === 'Off').length) },
    { label: 'Premium', value: fmtNumber(flows.filter((f) => f.usesPremium).length), sub: fmtPercent(flows.filter((f) => f.usesPremium).length, flows.length) + ' of flows' },
    { label: 'Orphaned (running)', value: fmtNumber(orphans), tone: orphans ? 'critical' : 'good' }
  ];
}

export function agentSummary(agents: IAgent[], makers: IMakersData): ISummaryItem[] {
  const published = agents.filter((b) => !!b.publishedTime).length;
  const noAuth = agents.filter((b) => b.authMode === 'None' && b.publishedTime).length;
  const orphans = agents.filter((b) => isLeft(b.owner.id, makers)).length;
  return [
    { label: 'Agents', value: fmtNumber(agents.length), sub: `${fmtNumber(new Set(agents.map((a) => a.environmentId)).size)} environments` },
    { label: 'Published', value: fmtNumber(published), sub: fmtPercent(published, agents.length) + ' of agents' },
    { label: 'Public (no auth)', value: fmtNumber(noAuth), tone: noAuth ? 'critical' : 'good', sub: noAuth ? 'Published without sign-in' : 'None' },
    { label: 'Orphaned', value: fmtNumber(orphans), tone: orphans ? 'critical' : 'good' }
  ];
}

export function makerSummary(data: IMakersData): ISummaryItem[] {
  const m = data.makers;
  const left = m.filter((x) => x.status === 'Deleted' || x.status === 'Disabled');
  const newMakers = m.filter((x) => (daysSince(x.firstCreated) || 9999) <= 30).length;
  const inactive = m.filter((x) => x.status === 'Active' && x.lastSignIn && (daysSince(x.lastSignIn) || 0) > 90).length;
  return [
    { label: 'Makers', value: fmtNumber(m.length), sub: `${fmtNumber(new Set(m.map((x) => x.department).filter(Boolean)).size)} departments` },
    { label: 'New in 30 days', value: fmtNumber(newMakers) },
    { label: 'Left or disabled', value: fmtNumber(left.length), tone: left.length ? 'critical' : 'good', sub: `Own ${fmtNumber(sumBy(left, (x) => x.apps + x.flows + x.agents))} resources` },
    { label: 'No sign-in 90d', value: data.signInUnavailable ? '—' : fmtNumber(inactive), sub: data.signInUnavailable ? 'Needs AuditLog.Read.All' : undefined }
  ];
}

/* ---------- Connectors & DLP ---------- */

export function connectorUsage(apps: IApp[], flows: IFlow[], dlp: IDlpPolicy[] | undefined): IConnectorUsage[] {
  const map: Record<string, IConnectorUsage & { m: Record<string, boolean>; e: Record<string, boolean> }> = {};
  const add = (list: IApp[] | IFlow[], kind: 'apps' | 'flows'): void => {
    (list as (IApp | IFlow)[]).forEach((item) => {
      item.connectors.forEach((c) => {
        let u = map[c.name];
        if (!u) u = map[c.name] = { name: c.name, displayName: c.displayName, tier: c.tier, apps: 0, flows: 0, makers: 0, environments: 0, businessIn: [], blockedIn: [], m: {}, e: {} };
        u[kind]++;
        if (item.owner.id) u.m[item.owner.id] = true;
        u.e[item.environmentId] = true;
      });
    });
  };
  add(apps, 'apps');
  add(flows, 'flows');
  return Object.keys(map)
    .map((k) => {
      const { m, e, ...u } = map[k];
      const business = (dlp || []).filter((p) => p.businessConnectors.indexOf(k) >= 0).map((p) => p.displayName);
      const blocked = (dlp || []).filter((p) => p.blockedConnectors.indexOf(k) >= 0).map((p) => p.displayName);
      return { ...u, makers: Object.keys(m).length, environments: Object.keys(e).length, businessIn: business, blockedIn: blocked };
    })
    .sort((a, b) => b.apps + b.flows - (a.apps + a.flows));
}

/* ---------- Licensing ---------- */

export function skuRows(skus: ISku[], prices: Record<string, number>): ISkuRow[] {
  return skus
    .map((s) => ({ ...s, available: Math.max(0, s.enabled - s.consumed), unitPrice: prices[s.skuPartNumber] }))
    .sort((a, b) => b.consumed - a.consumed);
}

export function unusedCost(rows: ISkuRow[]): number {
  return sumBy(rows, (r) => (r.unitPrice ? r.available * r.unitPrice : 0));
}

/* ---------- Microsoft 365 summaries ---------- */

export function siteSummary(data: ISitesData): ISummaryItem[] {
  const s = data.sites.filter((x) => !x.isDeleted);
  const inactive = s.filter((x) => (daysSince(x.lastActivity) === undefined ? true : (daysSince(x.lastActivity) || 0) > 90)).length;
  const anyone = s.filter((x) => x.anonymousLinks > 0).length;
  return [
    { label: 'Sites', value: fmtNumber(s.length) },
    { label: 'Storage used', value: fmtBytes(sumBy(s, (x) => x.storageUsedBytes)) },
    { label: 'Inactive 90d', value: fmtNumber(inactive), tone: inactive ? 'warning' : 'good', sub: fmtPercent(inactive, s.length) + ' of sites' },
    { label: 'Externally shared', value: fmtNumber(s.filter((x) => x.externalSharing).length) },
    { label: 'With Anyone links', value: fmtNumber(anyone), tone: anyone ? 'warning' : undefined }
  ];
}

export function teamsSummary(data: ITeamsData): ISummaryItem[] {
  const teams = data.groups.filter((g) => g.isTeam);
  const ownerless = data.groups.filter((g) => g.ownerCount === 0).length;
  const single = data.groups.filter((g) => g.ownerCount === 1).length;
  const inactive = data.activity.filter((t) => (daysSince(t.lastActivity) === undefined ? true : (daysSince(t.lastActivity) || 0) > 90)).length;
  return [
    { label: 'Microsoft 365 groups', value: fmtNumber(data.groups.length) },
    { label: 'Teams', value: fmtNumber(teams.length) },
    { label: 'Ownerless', value: fmtNumber(ownerless), tone: ownerless ? 'critical' : 'good', sub: `${fmtNumber(single)} with one owner` },
    { label: 'Inactive teams 90d', value: fmtNumber(inactive), tone: inactive ? 'warning' : undefined },
    { label: 'Teams with guests', value: fmtNumber(data.activity.filter((t) => t.guests > 0).length) }
  ];
}

export function inactiveLicensed(data: IAdoptionData): IAdoptionData['users'] {
  return data.users.filter((u) => !u.isDeleted && u.products.length > 0 && (daysSince(u.lastActivity) === undefined || (daysSince(u.lastActivity) || 0) > 30));
}

export function adoptionSummary(data: IAdoptionData, skus: ISku[] | undefined): ISummaryItem[] {
  const last = data.counts[data.counts.length - 1];
  const inactive = inactiveLicensed(data).length;
  const copilotSeats = (skus || []).filter((s) => /COPILOT/i.test(s.skuPartNumber) && !/STUDIO/i.test(s.skuPartNumber)).reduce((s, x) => s + x.consumed, 0);
  const copilotActive = data.copilotUsers.filter((u) => (daysSince(u.lastActivity) || 9999) <= 30).length;
  return [
    { label: 'Active users (daily)', value: last ? fmtNumber(last.office365) : '—', sub: last ? `on ${last.date}` : undefined },
    { label: 'Teams daily active', value: last ? fmtNumber(last.teams) : '—' },
    { label: 'Licensed, inactive 30d', value: fmtNumber(inactive), tone: inactive ? 'warning' : 'good', sub: 'Licence reclaim candidates' },
    {
      label: 'Copilot active 30d', value: data.copilotUnavailable ? '—' : fmtNumber(copilotActive),
      sub: data.copilotUnavailable ? 'Report not available' : copilotSeats ? `of ${fmtNumber(copilotSeats)} seats (${fmtPercent(copilotActive, copilotSeats)})` : undefined
    }
  ];
}

export function identitySummary(data: IIdentityData): ISummaryItem[] {
  const guests = data.users.filter((u) => u.userType === 'Guest');
  const stale = data.users.filter((u) => u.accountEnabled && (daysSince(u.lastSignIn) === undefined || (daysSince(u.lastSignIn) || 0) > 90));
  const admins = new Set<string>();
  data.roles.forEach((r) => r.members.forEach((m) => admins.add(m.id)));
  const ga = data.roles.filter((r) => r.displayName === 'Global Administrator')[0];
  const noMfa = data.mfa.filter((m) => !m.isMfaRegistered && m.userType.toLowerCase() === 'member').length;
  const expiring = data.credentials.filter((c) => c.daysLeft >= 0 && c.daysLeft <= 30).length;
  return [
    { label: 'Admin role holders', value: fmtNumber(admins.size), sub: ga ? `${ga.members.length} global admins` : undefined, tone: ga && ga.members.length > 4 ? 'warning' : undefined },
    { label: 'Guests', value: fmtNumber(guests.length), sub: data.signInUnavailable ? undefined : `${fmtNumber(guests.filter((g) => stale.indexOf(g) >= 0).length)} inactive 90d` },
    { label: 'Stale accounts 90d', value: data.signInUnavailable ? '—' : fmtNumber(stale.length), tone: data.signInUnavailable ? undefined : stale.length ? 'warning' : 'good', sub: data.signInUnavailable ? 'Needs Entra ID P1' : 'Enabled, no sign-in' },
    { label: 'Members without MFA', value: data.mfaUnavailable ? '—' : fmtNumber(noMfa), tone: data.mfaUnavailable ? undefined : noMfa ? 'critical' : 'good', sub: data.mfaUnavailable ? 'Needs Entra ID P1' : noMfa ? 'Not registered for MFA' : 'All registered' },
    { label: 'Credentials expiring 30d', value: fmtNumber(expiring), tone: expiring ? 'warning' : 'good', sub: expiring ? 'Rotate before they expire' : 'None due' }
  ];
}

/* ---------- Governance ---------- */

export function findings(pp: IPPData, dlp: IDlpPolicy[] | undefined, teams: ITeamsData | undefined, credentials: IAppCredential[] | undefined): IFinding[] {
  return evaluateGovernance({
    environments: pp.inv.environments,
    apps: pp.inv.apps,
    flows: pp.inv.flows,
    agents: pp.inv.agents,
    users: pp.makers.users,
    dlp,
    groups: teams ? teams.groups : undefined,
    credentials
  });
}

export function findingSummary(list: IFinding[]): ISummaryItem[] {
  const high = list.filter((f) => f.severity === 'High').length;
  const med = list.filter((f) => f.severity === 'Medium').length;
  return [
    { label: 'Findings', value: fmtNumber(list.length) },
    { label: 'High', value: fmtNumber(high), tone: high ? 'critical' : 'good' },
    { label: 'Medium', value: fmtNumber(med), tone: med ? 'warning' : 'good' },
    { label: 'Low', value: fmtNumber(list.length - high - med) }
  ];
}

export function envNameMap(envs: IEnvironment[]): Record<string, string> {
  const m: Record<string, string> = {};
  envs.forEach((e) => {
    m[e.id] = e.displayName;
  });
  return m;
}
