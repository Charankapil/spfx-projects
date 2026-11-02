/**
 * Normalised data model shared by every page, provider and export.
 * Raw API payloads are mapped into these shapes by the services so that the
 * UI never depends on undocumented field names directly.
 */

export interface IPrincipalRef {
  /** Entra ID object id when known. */
  id?: string;
  displayName?: string;
  email?: string;
}

export type ConnectorTier = 'Standard' | 'Premium' | 'Custom';

export interface IConnectorRef {
  /** Short connector name, e.g. "shared_sharepointonline". */
  name: string;
  displayName: string;
  tier: ConnectorTier;
}

export interface ICapacity {
  databaseMb: number;
  fileMb: number;
  logMb: number;
}

export interface IEnvironment {
  id: string;
  displayName: string;
  /** Production | Sandbox | Trial | Developer | Default | Teams | ... */
  type: string;
  region: string;
  isDefault: boolean;
  isManaged: boolean;
  hasDataverse: boolean;
  orgUrl?: string;
  dataverseVersion?: string;
  securityGroupId?: string;
  state: string;
  createdTime?: string;
  createdBy?: string;
  expirationTime?: string;
  capacity: ICapacity;
}

export type AppKind = 'Canvas' | 'Model-driven';

export interface IApp {
  id: string;
  displayName: string;
  environmentId: string;
  environmentName: string;
  kind: AppKind;
  owner: IPrincipalRef;
  createdTime?: string;
  lastModifiedTime?: string;
  lastPublishTime?: string;
  connectors: IConnectorRef[];
  sharedUsers?: number;
  sharedGroups?: number;
  /** Only known after a sharing scan (or from sample data). */
  sharedWithEveryone?: boolean;
  isManaged?: boolean;
  usesPremium: boolean;
  usesCustom: boolean;
  bypassConsent?: boolean;
}

export type FlowState = 'On' | 'Off' | 'Suspended';

export interface IFlow {
  id: string;
  displayName: string;
  environmentId: string;
  environmentName: string;
  state: FlowState;
  suspensionReason?: string;
  owner: IPrincipalRef;
  createdTime?: string;
  lastModifiedTime?: string;
  triggerType: string;
  connectors: IConnectorRef[];
  usesPremium: boolean;
  usesCustom: boolean;
  actionCount: number;
}

export type AgentAuth = 'None' | 'Microsoft' | 'Manual' | 'Unknown';

export interface IAgent {
  id: string;
  displayName: string;
  environmentId: string;
  environmentName: string;
  owner: IPrincipalRef;
  createdTime?: string;
  modifiedTime?: string;
  publishedTime?: string;
  authMode: AgentAuth;
  language?: string;
  state: 'Active' | 'Inactive';
}

export interface IEnvLoadError {
  environmentId: string;
  environmentName: string;
  area: 'Apps' | 'Flows' | 'Agents' | 'Model-driven apps';
  message: string;
}

export interface IInventory {
  environments: IEnvironment[];
  apps: IApp[];
  flows: IFlow[];
  agents: IAgent[];
  errors: IEnvLoadError[];
  loadedAt: string;
}

export type UserStatus = 'Active' | 'Disabled' | 'Deleted' | 'Unknown';

export interface IUserInfo {
  id: string;
  displayName?: string;
  upn?: string;
  department?: string;
  country?: string;
  jobTitle?: string;
  accountEnabled?: boolean;
  lastSignIn?: string;
  status: UserStatus;
}

export interface IMaker {
  id: string;
  displayName: string;
  upn: string;
  department: string;
  country: string;
  status: UserStatus;
  lastSignIn?: string;
  apps: number;
  flows: number;
  agents: number;
  environments: number;
  firstCreated?: string;
  lastCreated?: string;
}

export interface IDlpPolicy {
  id: string;
  displayName: string;
  scope: 'All environments' | 'All except' | 'Only selected';
  environmentIds: string[];
  businessConnectors: string[];
  nonBusinessConnectors: string[];
  blockedConnectors: string[];
  defaultGroup: string;
  createdBy?: string;
  lastModified?: string;
}

export interface ISku {
  skuId: string;
  skuPartNumber: string;
  friendlyName: string;
  enabled: number;
  consumed: number;
  suspended: number;
  warning: number;
  isPowerPlatform: boolean;
  appliesTo: string;
}

export type Severity = 'High' | 'Medium' | 'Low';

export interface IFinding {
  id: string;
  ruleId: string;
  severity: Severity;
  category: string;
  title: string;
  resourceType: string;
  resourceName: string;
  environmentName: string;
  owner: string;
  detail: string;
  link?: string;
}

/* ---------- Microsoft 365 ---------- */

export interface ISite {
  id: string;
  url: string;
  title: string;
  ownerName: string;
  ownerUpn: string;
  lastActivity?: string;
  template: string;
  storageUsedBytes: number;
  storageAllocatedBytes: number;
  fileCount: number;
  activeFileCount: number;
  pageViews: number;
  externalSharing: boolean;
  anonymousLinks: number;
  isDeleted: boolean;
}

export interface IStoragePoint {
  date: string;
  bytes: number;
}

export interface ITenantSharing {
  sharingCapability: string;
  isResharingByExternalUsersEnabled?: boolean;
  sharingDomainRestrictionMode?: string;
  isUnmanagedSyncAppForTenantRestricted?: boolean;
  oneDriveLoopSharingCapability?: string;
}

export interface IGroup {
  id: string;
  displayName: string;
  mail?: string;
  createdTime?: string;
  visibility: string;
  isTeam: boolean;
  ownerCount: number;
}

export interface ITeamActivity {
  id: string;
  name: string;
  teamType: string;
  lastActivity?: string;
  activeUsers: number;
  guests: number;
  activeExternalUsers: number;
  activeChannels: number;
  messages: number;
  meetings: number;
}

export interface IActiveUserCounts {
  date: string;
  office365: number;
  exchange: number;
  oneDrive: number;
  sharePoint: number;
  teams: number;
}

export interface IUserActivity {
  upn: string;
  displayName: string;
  products: string[];
  lastActivity?: string;
  isDeleted: boolean;
}

export interface ICopilotUser {
  upn: string;
  displayName: string;
  lastActivity?: string;
}

export interface IDirectoryRoleMember {
  id: string;
  displayName: string;
  upn: string;
  type: string;
}

export interface IDirectoryRole {
  id: string;
  displayName: string;
  members: IDirectoryRoleMember[];
}

export interface IUserAccount {
  id: string;
  displayName: string;
  upn: string;
  userType: string;
  accountEnabled: boolean;
  createdTime?: string;
  lastSignIn?: string;
}

export interface IMfaRegistration {
  id: string;
  upn: string;
  displayName: string;
  isMfaRegistered: boolean;
  isAdmin: boolean;
  userType: string;
  methods: string[];
}

export interface IAppCredential {
  appId: string;
  appDisplayName: string;
  type: 'Secret' | 'Certificate';
  displayName: string;
  endDateTime: string;
  daysLeft: number;
}

export interface IIdentityData {
  roles: IDirectoryRole[];
  users: IUserAccount[];
  /** True when sign-in activity could not be read (licence or permission). */
  signInUnavailable: boolean;
  mfa: IMfaRegistration[];
  mfaUnavailable: boolean;
  credentials: IAppCredential[];
}

export interface IHealthOverview {
  service: string;
  status: string;
}

export interface IHealthIssue {
  id: string;
  title: string;
  service: string;
  status: string;
  classification: string;
  startDateTime?: string;
}

export interface IMessage {
  id: string;
  title: string;
  services: string[];
  category: string;
  lastModified?: string;
  actionRequiredBy?: string;
  isMajorChange: boolean;
}

export interface IServiceHealth {
  overviews: IHealthOverview[];
  issues: IHealthIssue[];
  messages: IMessage[];
}

export interface ITeamsData {
  groups: IGroup[];
  activity: ITeamActivity[];
}

export interface IAdoptionData {
  counts: IActiveUserCounts[];
  users: IUserActivity[];
  copilotUsers: ICopilotUser[];
  copilotUnavailable: boolean;
}

export interface ISitesData {
  sites: ISite[];
  storageTrend: IStoragePoint[];
  sharing?: ITenantSharing;
}

/* ---------- Data source diagnostics ---------- */

export type SourceState = 'pending' | 'ok' | 'forbidden' | 'blocked' | 'error' | 'skipped';

export interface ISourceStatus {
  id: string;
  label: string;
  endpoint: string;
  permission: string;
  state: SourceState;
  message?: string;
  checkedAt?: string;
}

export interface IProgress {
  label: string;
  done: number;
  total: number;
}
