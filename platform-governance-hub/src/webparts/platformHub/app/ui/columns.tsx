import * as React from 'react';
import { Link } from '@fluentui/react';
import {
  IAgent, IApp, IAppCredential, IConnectorRef, IDirectoryRole, IDlpPolicy, IEnvironment, IFinding, IFlow, IGroup, IHealthIssue,
  IMaker, IMessage, IMfaRegistration, IPrincipalRef, ISite, ISku, ITeamActivity, IUserAccount, IUserActivity, IUserInfo
} from '../models';
import { IColumnSpec } from './components/DataGrid';
import { Pill, severityTone } from './components/Layout';
import { daysSince, fmtAge, fmtBytes, fmtMb, fmtNumber, safeUrl } from '../logic/format';
import { ruleAction } from '../logic/governance';

export function ownerLabel(ref: IPrincipalRef, users: Record<string, IUserInfo>): string {
  const u = ref.id ? users[ref.id] : undefined;
  if (u && u.status === 'Deleted') return ref.displayName ? `${ref.displayName} (deleted)` : 'Deleted user';
  return ref.displayName || (u && (u.displayName || u.upn)) || ref.email || (ref.id ? 'Unresolved user' : 'Unknown');
}

export function ownerStatus(ref: IPrincipalRef, users: Record<string, IUserInfo>): string {
  const u = ref.id ? users[ref.id] : undefined;
  return u ? u.status : 'Unknown';
}

function statusCell(status: string): React.ReactNode {
  if (status === 'Deleted') return <Pill tone="critical" text="Left org" />;
  if (status === 'Disabled') return <Pill tone="warning" text="Disabled" />;
  if (status === 'Active') return <Pill tone="good" text="Active" />;
  return <Pill tone="neutral" text="Unknown" />;
}

const connectorList = (c: IConnectorRef[]): string => c.map((x) => x.displayName).join(', ');

/* ---------- Power Platform ---------- */

export interface IEnvRow extends IEnvironment {
  apps: number;
  flows: number;
  agents: number;
  makers: number;
  dlpCovered?: boolean;
}

export const envColumns: IColumnSpec<IEnvRow>[] = [
  { key: 'name', header: 'Environment', value: (e) => e.displayName, minWidth: 160, width: 32 },
  { key: 'type', header: 'Type', value: (e) => e.type, maxWidth: 100 },
  { key: 'region', header: 'Region', value: (e) => e.region, maxWidth: 110 },
  { key: 'dv', header: 'Dataverse', value: (e) => e.hasDataverse, maxWidth: 80 },
  { key: 'managed', header: 'Managed', value: (e) => e.isManaged, maxWidth: 80 },
  { key: 'sg', header: 'Security group', value: (e) => !!e.securityGroupId, maxWidth: 100 },
  {
    key: 'dlp', header: 'DLP', value: (e) => (e.dlpCovered === undefined ? '' : e.dlpCovered ? 'Covered' : 'Not covered'),
    render: (e) => (e.dlpCovered === undefined ? '—' : e.dlpCovered ? <Pill tone="good" text="Covered" /> : <Pill tone="critical" text="None" />), maxWidth: 100
  },
  { key: 'apps', header: 'Apps', value: (e) => e.apps, isNumeric: true },
  { key: 'flows', header: 'Flows', value: (e) => e.flows, isNumeric: true },
  { key: 'agents', header: 'Agents', value: (e) => e.agents, isNumeric: true },
  { key: 'makers', header: 'Makers', value: (e) => e.makers, isNumeric: true },
  { key: 'db', header: 'Database', value: (e) => Math.round(e.capacity.databaseMb), render: (e) => fmtMb(e.capacity.databaseMb), isNumeric: true },
  { key: 'file', header: 'File', value: (e) => Math.round(e.capacity.fileMb), render: (e) => fmtMb(e.capacity.fileMb), isNumeric: true },
  { key: 'log', header: 'Log', value: (e) => Math.round(e.capacity.logMb), render: (e) => fmtMb(e.capacity.logMb), isNumeric: true, exportOnly: true },
  { key: 'created', header: 'Created', value: (e) => e.createdTime, isDate: true },
  { key: 'createdBy', header: 'Created by', value: (e) => e.createdBy, exportOnly: true },
  { key: 'state', header: 'State', value: (e) => e.state, exportOnly: true },
  { key: 'url', header: 'Dataverse URL', value: (e) => e.orgUrl, exportOnly: true, width: 40 },
  { key: 'id', header: 'Environment ID', value: (e) => e.id, exportOnly: true, width: 40 }
];

export function appColumns(users: Record<string, IUserInfo>): IColumnSpec<IApp>[] {
  return [
    { key: 'name', header: 'App', value: (a) => a.displayName, minWidth: 170, width: 32 },
    { key: 'kind', header: 'Type', value: (a) => a.kind, maxWidth: 100 },
    { key: 'env', header: 'Environment', value: (a) => a.environmentName, minWidth: 120 },
    { key: 'owner', header: 'Owner', value: (a) => ownerLabel(a.owner, users), minWidth: 130 },
    { key: 'ownerStatus', header: 'Owner status', value: (a) => ownerStatus(a.owner, users), render: (a) => statusCell(ownerStatus(a.owner, users)), maxWidth: 100 },
    { key: 'connectors', header: 'Connectors', value: (a) => connectorList(a.connectors), minWidth: 160, width: 40 },
    { key: 'premium', header: 'Premium', value: (a) => a.usesPremium, maxWidth: 70 },
    { key: 'users', header: 'Shared users', value: (a) => a.sharedUsers, isNumeric: true },
    {
      key: 'everyone', header: 'Everyone', value: (a) => (a.sharedWithEveryone === undefined ? '' : a.sharedWithEveryone),
      render: (a) => (a.sharedWithEveryone ? <Pill tone="critical" text="Yes" /> : a.sharedWithEveryone === false ? 'No' : '—'), maxWidth: 80
    },
    { key: 'modified', header: 'Modified', value: (a) => a.lastModifiedTime, isDate: true },
    { key: 'created', header: 'Created', value: (a) => a.createdTime, isDate: true, exportOnly: true },
    { key: 'published', header: 'Published', value: (a) => a.lastPublishTime, isDate: true, exportOnly: true },
    { key: 'custom', header: 'Custom connector', value: (a) => a.usesCustom, exportOnly: true },
    { key: 'groups', header: 'Shared groups', value: (a) => a.sharedGroups, exportOnly: true },
    { key: 'managed', header: 'Managed solution', value: (a) => a.isManaged, exportOnly: true },
    { key: 'ownerEmail', header: 'Owner email', value: (a) => a.owner.email || (a.owner.id && users[a.owner.id] ? users[a.owner.id].upn : ''), exportOnly: true },
    { key: 'id', header: 'App ID', value: (a) => a.id, exportOnly: true, width: 40 },
    { key: 'envId', header: 'Environment ID', value: (a) => a.environmentId, exportOnly: true, width: 40 }
  ];
}

export function flowColumns(users: Record<string, IUserInfo>): IColumnSpec<IFlow>[] {
  return [
    { key: 'name', header: 'Flow', value: (f) => f.displayName, minWidth: 180, width: 36 },
    {
      key: 'state', header: 'State', value: (f) => f.state,
      render: (f) => <Pill tone={f.state === 'On' ? 'good' : f.state === 'Suspended' ? 'critical' : 'neutral'} text={f.state} />, maxWidth: 100
    },
    { key: 'trigger', header: 'Trigger', value: (f) => f.triggerType, maxWidth: 110 },
    { key: 'env', header: 'Environment', value: (f) => f.environmentName, minWidth: 120 },
    { key: 'owner', header: 'Owner', value: (f) => ownerLabel(f.owner, users), minWidth: 130 },
    { key: 'ownerStatus', header: 'Owner status', value: (f) => ownerStatus(f.owner, users), render: (f) => statusCell(ownerStatus(f.owner, users)), maxWidth: 100 },
    { key: 'connectors', header: 'Connectors', value: (f) => connectorList(f.connectors), minWidth: 160, width: 40 },
    { key: 'premium', header: 'Premium', value: (f) => f.usesPremium, maxWidth: 70 },
    { key: 'actions', header: 'Actions', value: (f) => f.actionCount, isNumeric: true },
    { key: 'modified', header: 'Modified', value: (f) => f.lastModifiedTime, isDate: true },
    { key: 'created', header: 'Created', value: (f) => f.createdTime, isDate: true, exportOnly: true },
    { key: 'reason', header: 'Suspension reason', value: (f) => f.suspensionReason, exportOnly: true },
    { key: 'custom', header: 'Custom connector', value: (f) => f.usesCustom, exportOnly: true },
    { key: 'id', header: 'Flow ID', value: (f) => f.id, exportOnly: true, width: 40 },
    { key: 'envId', header: 'Environment ID', value: (f) => f.environmentId, exportOnly: true, width: 40 }
  ];
}

export function agentColumns(users: Record<string, IUserInfo>): IColumnSpec<IAgent>[] {
  return [
    { key: 'name', header: 'Agent', value: (b) => b.displayName, minWidth: 170, width: 32 },
    { key: 'env', header: 'Environment', value: (b) => b.environmentName, minWidth: 120 },
    { key: 'owner', header: 'Owner', value: (b) => ownerLabel(b.owner, users), minWidth: 130 },
    { key: 'ownerStatus', header: 'Owner status', value: (b) => ownerStatus(b.owner, users), render: (b) => statusCell(ownerStatus(b.owner, users)), maxWidth: 100 },
    {
      key: 'auth', header: 'Authentication', value: (b) => b.authMode,
      render: (b) => (b.authMode === 'None' ? <Pill tone={b.publishedTime ? 'critical' : 'warning'} text="None" /> : b.authMode), maxWidth: 110
    },
    { key: 'published', header: 'Published', value: (b) => b.publishedTime, isDate: true },
    { key: 'state', header: 'State', value: (b) => b.state, maxWidth: 80 },
    { key: 'modified', header: 'Modified', value: (b) => b.modifiedTime, isDate: true },
    { key: 'created', header: 'Created', value: (b) => b.createdTime, isDate: true, exportOnly: true },
    { key: 'id', header: 'Agent ID', value: (b) => b.id, exportOnly: true, width: 40 }
  ];
}

export const makerColumns: IColumnSpec<IMaker>[] = [
  { key: 'name', header: 'Maker', value: (m) => m.displayName, minWidth: 150, width: 28 },
  { key: 'upn', header: 'User name', value: (m) => m.upn, minWidth: 170, width: 34 },
  { key: 'dept', header: 'Department', value: (m) => m.department, minWidth: 110 },
  { key: 'country', header: 'Country', value: (m) => m.country, maxWidth: 120 },
  { key: 'status', header: 'Status', value: (m) => m.status, render: (m) => statusCell(m.status), maxWidth: 100 },
  { key: 'apps', header: 'Apps', value: (m) => m.apps, isNumeric: true },
  { key: 'flows', header: 'Flows', value: (m) => m.flows, isNumeric: true },
  { key: 'agents', header: 'Agents', value: (m) => m.agents, isNumeric: true },
  { key: 'envs', header: 'Environments', value: (m) => m.environments, isNumeric: true },
  { key: 'signin', header: 'Last sign-in', value: (m) => m.lastSignIn, isDate: true },
  { key: 'first', header: 'First created', value: (m) => m.firstCreated, isDate: true, exportOnly: true },
  { key: 'last', header: 'Last created', value: (m) => m.lastCreated, isDate: true },
  { key: 'id', header: 'Object ID', value: (m) => m.id, exportOnly: true, width: 40 }
];

export interface IConnectorUsage {
  name: string;
  displayName: string;
  tier: string;
  apps: number;
  flows: number;
  makers: number;
  environments: number;
  businessIn: string[];
  blockedIn: string[];
}

export const connectorColumns: IColumnSpec<IConnectorUsage>[] = [
  { key: 'name', header: 'Connector', value: (c) => c.displayName, minWidth: 160, width: 30 },
  {
    key: 'tier', header: 'Tier', value: (c) => c.tier,
    render: (c) => (c.tier === 'Custom' ? <Pill tone="warning" text="Custom" /> : c.tier), maxWidth: 90
  },
  { key: 'apps', header: 'Apps', value: (c) => c.apps, isNumeric: true },
  { key: 'flows', header: 'Flows', value: (c) => c.flows, isNumeric: true },
  { key: 'makers', header: 'Makers', value: (c) => c.makers, isNumeric: true },
  { key: 'envs', header: 'Environments', value: (c) => c.environments, isNumeric: true },
  { key: 'business', header: 'Business in', value: (c) => c.businessIn.join(', '), minWidth: 140 },
  {
    key: 'blocked', header: 'Blocked in', value: (c) => c.blockedIn.join(', '),
    render: (c) => (c.blockedIn.length ? <Pill tone="warning" text={c.blockedIn.join(', ')} /> : '—'), minWidth: 140
  },
  { key: 'id', header: 'Connector ID', value: (c) => c.name, exportOnly: true, width: 40 }
];

export function dlpColumns(envNames: Record<string, string>): IColumnSpec<IDlpPolicy>[] {
  return [
    { key: 'name', header: 'Policy', value: (p) => p.displayName, minWidth: 160, width: 30 },
    { key: 'scope', header: 'Scope', value: (p) => p.scope, maxWidth: 120 },
    { key: 'envs', header: 'Environments', value: (p) => p.environmentIds.map((id) => envNames[id] || id).join(', '), minWidth: 180, width: 50 },
    { key: 'business', header: 'Business', value: (p) => p.businessConnectors.length, isNumeric: true },
    { key: 'nonbusiness', header: 'Non-business', value: (p) => p.nonBusinessConnectors.length, isNumeric: true },
    { key: 'blocked', header: 'Blocked', value: (p) => p.blockedConnectors.length, isNumeric: true },
    { key: 'default', header: 'Default group', value: (p) => p.defaultGroup, maxWidth: 110 },
    { key: 'modified', header: 'Modified', value: (p) => p.lastModified, isDate: true },
    { key: 'by', header: 'Created by', value: (p) => p.createdBy, exportOnly: true },
    { key: 'blockedList', header: 'Blocked connectors', value: (p) => p.blockedConnectors.join(', '), exportOnly: true, width: 50 }
  ];
}

export interface ISkuRow extends ISku {
  available: number;
  unitPrice?: number;
}

export const skuColumns: IColumnSpec<ISkuRow>[] = [
  { key: 'name', header: 'Licence', value: (s) => s.friendlyName, minWidth: 190, width: 36 },
  { key: 'part', header: 'SKU', value: (s) => s.skuPartNumber, minWidth: 150, width: 30 },
  { key: 'enabled', header: 'Purchased', value: (s) => s.enabled, isNumeric: true },
  { key: 'consumed', header: 'Assigned', value: (s) => s.consumed, isNumeric: true },
  { key: 'available', header: 'Available', value: (s) => s.available, isNumeric: true },
  { key: 'pct', header: 'Use', value: (s) => (s.enabled ? Math.round((s.consumed / s.enabled) * 100) : 0), render: (s) => (s.enabled ? `${Math.round((s.consumed / s.enabled) * 100)}%` : '—'), isNumeric: true },
  { key: 'warning', header: 'In warning', value: (s) => s.warning, isNumeric: true },
  { key: 'appliesTo', header: 'Applies to', value: (s) => s.appliesTo, exportOnly: true }
];

/* ---------- Governance ---------- */

export const findingColumns: IColumnSpec<IFinding>[] = [
  { key: 'sev', header: 'Severity', value: (f) => f.severity, render: (f) => <Pill tone={severityTone(f.severity)} text={f.severity} />, maxWidth: 90 },
  { key: 'title', header: 'Finding', value: (f) => f.title, minWidth: 200, width: 40 },
  { key: 'type', header: 'Resource type', value: (f) => f.resourceType, maxWidth: 110 },
  {
    key: 'name', header: 'Resource', value: (f) => f.resourceName,
    render: (f) => (safeUrl(f.link) ? <Link href={safeUrl(f.link)} target="_blank" rel="noopener noreferrer" onClick={(e) => e.stopPropagation()}>{f.resourceName}</Link> : f.resourceName), minWidth: 160, width: 32
  },
  { key: 'env', header: 'Environment', value: (f) => f.environmentName, minWidth: 110 },
  { key: 'owner', header: 'Owner', value: (f) => f.owner, minWidth: 120 },
  { key: 'detail', header: 'Detail', value: (f) => f.detail, minWidth: 200, width: 50 },
  { key: 'category', header: 'Category', value: (f) => f.category, exportOnly: true },
  { key: 'action', header: 'Recommended action', value: (f) => ruleAction(f.ruleId), exportOnly: true, width: 60 },
  { key: 'link', header: 'Link', value: (f) => f.link, exportOnly: true, width: 60 }
];

/* ---------- Microsoft 365 ---------- */

export const siteColumns: IColumnSpec<ISite>[] = [
  { key: 'title', header: 'Site', value: (s) => s.title, render: (s) => (safeUrl(s.url) ? <Link href={safeUrl(s.url)} target="_blank" rel="noopener noreferrer" onClick={(e) => e.stopPropagation()}>{s.title}</Link> : s.title), minWidth: 170, width: 30 },
  { key: 'template', header: 'Template', value: (s) => s.template, maxWidth: 140 },
  { key: 'owner', header: 'Owner', value: (s) => s.ownerName, minWidth: 120 },
  { key: 'storage', header: 'Storage', value: (s) => s.storageUsedBytes, render: (s) => fmtBytes(s.storageUsedBytes), isNumeric: true },
  { key: 'files', header: 'Files', value: (s) => s.fileCount, isNumeric: true },
  { key: 'views', header: 'Page views (30d)', value: (s) => s.pageViews, isNumeric: true },
  { key: 'external', header: 'External sharing', value: (s) => s.externalSharing, maxWidth: 100 },
  { key: 'anon', header: 'Anyone links', value: (s) => s.anonymousLinks, isNumeric: true },
  {
    key: 'activity', header: 'Last activity', value: (s) => s.lastActivity,
    render: (s) => {
      const d = daysSince(s.lastActivity);
      return d === undefined || d > 90 ? <Pill tone="warning" text={s.lastActivity ? fmtAge(s.lastActivity) : 'No activity'} /> : fmtAge(s.lastActivity);
    }, isDate: true, minWidth: 110
  },
  { key: 'url', header: 'URL', value: (s) => s.url, exportOnly: true, width: 60 },
  { key: 'activeFiles', header: 'Active files (30d)', value: (s) => s.activeFileCount, exportOnly: true }
];

export const groupColumns: IColumnSpec<IGroup>[] = [
  { key: 'name', header: 'Group', value: (g) => g.displayName, minWidth: 180, width: 36 },
  { key: 'team', header: 'Team', value: (g) => g.isTeam, maxWidth: 70 },
  { key: 'vis', header: 'Privacy', value: (g) => g.visibility, maxWidth: 90 },
  {
    key: 'owners', header: 'Owners', value: (g) => g.ownerCount,
    render: (g) => (g.ownerCount === 0 ? <Pill tone="critical" text="None" /> : g.ownerCount === 1 ? <Pill tone="warning" text="1" /> : fmtNumber(g.ownerCount)), isNumeric: true
  },
  { key: 'created', header: 'Created', value: (g) => g.createdTime, isDate: true },
  { key: 'mail', header: 'Email', value: (g) => g.mail, minWidth: 170, width: 36 },
  { key: 'id', header: 'Group ID', value: (g) => g.id, exportOnly: true, width: 40 }
];

export const teamActivityColumns: IColumnSpec<ITeamActivity>[] = [
  { key: 'name', header: 'Team', value: (t) => t.name, minWidth: 180, width: 36 },
  { key: 'type', header: 'Privacy', value: (t) => t.teamType, maxWidth: 90 },
  { key: 'active', header: 'Active users', value: (t) => t.activeUsers, isNumeric: true },
  { key: 'guests', header: 'Guests', value: (t) => t.guests, isNumeric: true },
  { key: 'ext', header: 'Active external', value: (t) => t.activeExternalUsers, isNumeric: true },
  { key: 'channels', header: 'Active channels', value: (t) => t.activeChannels, isNumeric: true },
  { key: 'messages', header: 'Messages', value: (t) => t.messages, isNumeric: true },
  { key: 'meetings', header: 'Meetings', value: (t) => t.meetings, isNumeric: true },
  { key: 'last', header: 'Last activity', value: (t) => t.lastActivity, isDate: true }
];

export const inactiveUserColumns: IColumnSpec<IUserActivity>[] = [
  { key: 'name', header: 'User', value: (u) => u.displayName, minWidth: 150, width: 28 },
  { key: 'upn', header: 'User name', value: (u) => u.upn, minWidth: 180, width: 36 },
  { key: 'products', header: 'Licences', value: (u) => u.products.join(', '), minWidth: 200, width: 50 },
  { key: 'last', header: 'Last activity', value: (u) => u.lastActivity, isDate: true, render: (u) => (u.lastActivity ? fmtAge(u.lastActivity) : 'No activity') }
];

export interface IRoleMemberRow {
  role: string;
  name: string;
  upn: string;
  type: string;
}

export function flattenRoles(roles: IDirectoryRole[]): IRoleMemberRow[] {
  const rows: IRoleMemberRow[] = [];
  roles.forEach((r) => r.members.forEach((m) => rows.push({ role: r.displayName, name: m.displayName, upn: m.upn, type: m.type })));
  return rows;
}

export const roleColumns: IColumnSpec<IRoleMemberRow>[] = [
  { key: 'role', header: 'Role', value: (r) => r.role, minWidth: 180, width: 34 },
  { key: 'name', header: 'Member', value: (r) => r.name, minWidth: 150, width: 28 },
  { key: 'upn', header: 'User name', value: (r) => r.upn, minWidth: 180, width: 36 },
  { key: 'type', header: 'Type', value: (r) => r.type, maxWidth: 110 }
];

export const accountColumns: IColumnSpec<IUserAccount>[] = [
  { key: 'name', header: 'User', value: (u) => u.displayName, minWidth: 150, width: 28 },
  { key: 'upn', header: 'User name', value: (u) => u.upn, minWidth: 200, width: 44 },
  { key: 'type', header: 'Type', value: (u) => u.userType, maxWidth: 80 },
  { key: 'enabled', header: 'Enabled', value: (u) => u.accountEnabled, maxWidth: 70 },
  { key: 'created', header: 'Created', value: (u) => u.createdTime, isDate: true },
  {
    key: 'signin', header: 'Last sign-in', value: (u) => u.lastSignIn, isDate: true,
    render: (u) => {
      const d = daysSince(u.lastSignIn);
      return d === undefined ? <Pill tone="warning" text="Never" /> : d > 90 ? <Pill tone="warning" text={fmtAge(u.lastSignIn)} /> : fmtAge(u.lastSignIn);
    }
  },
  { key: 'id', header: 'Object ID', value: (u) => u.id, exportOnly: true, width: 40 }
];

export const mfaColumns: IColumnSpec<IMfaRegistration>[] = [
  { key: 'name', header: 'User', value: (m) => m.displayName, minWidth: 150, width: 28 },
  { key: 'upn', header: 'User name', value: (m) => m.upn, minWidth: 200, width: 40 },
  { key: 'admin', header: 'Admin', value: (m) => m.isAdmin, maxWidth: 70 },
  { key: 'mfa', header: 'MFA registered', value: (m) => m.isMfaRegistered, render: (m) => (m.isMfaRegistered ? 'Yes' : <Pill tone={m.isAdmin ? 'critical' : 'warning'} text="No" />), maxWidth: 110 },
  { key: 'methods', header: 'Methods', value: (m) => m.methods.join(', '), minWidth: 160, width: 40 }
];

export const credentialColumns: IColumnSpec<IAppCredential>[] = [
  { key: 'app', header: 'App registration', value: (c) => c.appDisplayName, minWidth: 180, width: 34 },
  { key: 'type', header: 'Type', value: (c) => c.type, maxWidth: 90 },
  { key: 'name', header: 'Credential', value: (c) => c.displayName, minWidth: 120 },
  { key: 'end', header: 'Expires', value: (c) => c.endDateTime, isDate: true },
  {
    key: 'days', header: 'Days left', value: (c) => c.daysLeft, isNumeric: true,
    render: (c) => (c.daysLeft < 0 ? <Pill tone="critical" text="Expired" /> : c.daysLeft <= 30 ? <Pill tone="warning" text={String(c.daysLeft)} /> : String(c.daysLeft))
  },
  { key: 'appId', header: 'Application ID', value: (c) => c.appId, exportOnly: true, width: 40 }
];

export const issueColumns: IColumnSpec<IHealthIssue>[] = [
  { key: 'service', header: 'Service', value: (i) => i.service, minWidth: 130 },
  { key: 'title', header: 'Issue', value: (i) => i.title, minWidth: 240, width: 60 },
  { key: 'class', header: 'Type', value: (i) => i.classification, maxWidth: 90 },
  { key: 'status', header: 'Status', value: (i) => i.status, maxWidth: 140 },
  { key: 'start', header: 'Started', value: (i) => i.startDateTime, isDate: true },
  { key: 'id', header: 'ID', value: (i) => i.id, maxWidth: 90 }
];

export const messageColumns: IColumnSpec<IMessage>[] = [
  { key: 'title', header: 'Message', value: (m) => m.title, minWidth: 240, width: 60 },
  { key: 'services', header: 'Services', value: (m) => m.services.join(', '), minWidth: 140 },
  { key: 'category', header: 'Category', value: (m) => m.category, maxWidth: 110 },
  { key: 'major', header: 'Major change', value: (m) => m.isMajorChange, maxWidth: 90 },
  { key: 'action', header: 'Act by', value: (m) => m.actionRequiredBy, isDate: true },
  { key: 'updated', header: 'Updated', value: (m) => m.lastModified, isDate: true },
  { key: 'id', header: 'ID', value: (m) => m.id, maxWidth: 80 }
];
