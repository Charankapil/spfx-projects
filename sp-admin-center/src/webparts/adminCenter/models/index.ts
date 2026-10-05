export interface ITarget {
  /** Absolute URL of the web being administered (no trailing slash). */
  webUrl: string;
  /** Absolute URL of the site collection that web belongs to. */
  siteUrl: string;
  title: string;
}

export interface IWebInfo {
  id: string;
  title: string;
  description: string;
  url: string;
  serverRelativeUrl: string;
  created?: Date;
  lastModified?: Date;
  template: string;
  language: number;
  hasUniquePermissions: boolean;
  siteLogoUrl?: string;
}

export interface ISiteInfo {
  id: string;
  url: string;
  readOnly: boolean;
  storageBytes: number;
  /** 0..1 of the site's quota that is used. */
  storageFraction: number;
  groupId?: string;
  hubSiteId?: string;
  classification?: string;
  ownerTitle?: string;
  ownerEmail?: string;
}

export interface ISubWeb {
  id: string;
  title: string;
  url: string;
  template: string;
  created?: Date;
  lastModified?: Date;
  hasUniquePermissions: boolean;
}

export type ListKind = 'Library' | 'List' | 'Other';

export interface IListInfo {
  id: string;
  title: string;
  description: string;
  kind: ListKind;
  baseTemplate: number;
  itemCount: number;
  created?: Date;
  lastModified?: Date;
  versioning: boolean;
  minorVersions: boolean;
  moderation: boolean;
  hidden: boolean;
  noCrawl: boolean;
  uniquePermissions: boolean;
  isSystem: boolean;
  url: string;
  rootFolder: string;
}

export type UserKind = 'Member' | 'Guest' | 'Group' | 'OrgWide' | 'System';

export interface IUserInfo {
  id: number;
  title: string;
  email: string;
  loginName: string;
  principalType: number;
  isSiteAdmin: boolean;
  kind: UserKind;
}

export interface IGroupInfo {
  id: number;
  title: string;
  description: string;
  ownerTitle: string;
  role?: 'Owners' | 'Members' | 'Visitors';
  allowMembersEditMembership: boolean;
}

export interface IRoleAssignment {
  principalId: number;
  principalTitle: string;
  principalLogin: string;
  principalType: number;
  roles: string[];
}

export interface IRecycleItem {
  id: string;
  title: string;
  leafName: string;
  dirName: string;
  itemType: number;
  stage: 1 | 2;
  size: number;
  deletedDate?: Date;
  deletedBy: string;
  author: string;
}

export interface ISiteRow {
  title: string;
  url: string;
  template: string;
  created?: Date;
  lastModified?: Date;
  groupConnected: boolean;
  manageable: boolean;
}

export interface IFileRow {
  title: string;
  path: string;
  size: number;
  extension: string;
  modified?: Date;
  author: string;
}

export interface IFileTypeCount {
  type: string;
  count: number;
}

export interface IChangeRow {
  time?: Date;
  changeType: string;
  objectKind: string;
  /** Numeric / GUID id of the changed object; the view resolves it to a name. */
  objectId: string;
  severity: 'security' | 'structure' | 'info';
}

export type Severity = 'critical' | 'warning' | 'info' | 'good';

export interface IFinding {
  id: string;
  severity: Severity;
  title: string;
  detail: string;
  /** Where in the app the admin can act on it. */
  view?: ViewKey;
  /** Points deducted from the health score. */
  penalty: number;
}

export type ViewKey =
  | 'overview'
  | 'tenant'
  | 'sites'
  | 'people'
  | 'content'
  | 'storage'
  | 'growth'
  | 'recycle'
  | 'activity'
  | 'health'
  | 'settings';

export interface IActionLogEntry {
  at: Date;
  action: string;
  target: string;
  ok: boolean;
  message?: string;
}
