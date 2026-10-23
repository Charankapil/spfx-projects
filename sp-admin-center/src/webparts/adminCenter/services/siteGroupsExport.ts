import { ISiteRow } from '../models';
import { IDirectoryObject, IGroupMembership, IGroupSummary } from './GraphApi';
import { siteType } from './siteTypes';

/**
 * The "sites with their owners and members" export. Pure functions: the view gathers
 * sites (search), group names (one Graph listing), Microsoft 365 group memberships
 * (batched Graph reads) and SharePoint group memberships for the other sites (one
 * SharePoint request per site), and these turn them into CSV rows.
 */

export type ExportLayout = 'site' | 'person';

/** Who owns and belongs to a site: from its Microsoft 365 group, or from its SharePoint groups. */
export interface ISiteAccess {
  source: 'm365' | 'sharepoint';
  owners: IDirectoryObject[];
  members: IDirectoryObject[];
  visitors: IDirectoryObject[];
  /** Site collection administrators (only when asked for). */
  admins?: IDirectoryObject[];
  /** SharePoint group names (SharePoint source). */
  ownerGroup?: string;
  memberGroup?: string;
  visitorGroup?: string;
  truncated: boolean;
  error?: string;
}

export function fromGroupMembership(m: IGroupMembership): ISiteAccess {
  return { source: 'm365', owners: m.owners, members: m.members, visitors: [], truncated: m.truncated, error: m.error };
}

export const SITE_HEADERS = [
  'Site URL',
  'Site title',
  'Site type',
  'Membership from',
  'Microsoft 365 group',
  'Group name',
  'Group e-mail',
  'Privacy',
  'Teams',
  'Owners (count)',
  'Owners',
  'Members (count)',
  'Members',
  'Visitors (count)',
  'Visitors',
  'Site admins',
  'Note'
];
export const PERSON_HEADERS = ['Site URL', 'Site title', 'Membership from', 'Group', 'Role', 'Name', 'E-mail / UPN', 'User type'];

export function personText(p: IDirectoryObject): string {
  const id = p.mail || p.userPrincipalName || '';
  return id && id.toLowerCase() !== p.displayName.toLowerCase() ? `${p.displayName} <${id}>` : p.displayName;
}

function userType(p: IDirectoryObject): string {
  if (p.type === 'group') {
    return 'Group';
  }
  if (p.type === 'servicePrincipal') {
    return 'App';
  }
  return (p.userType || '').toLowerCase() === 'guest' ? 'Guest' : 'Member';
}

function sourceText(a: ISiteAccess | undefined): string {
  if (!a) {
    return '';
  }
  return a.source === 'm365' ? 'Microsoft 365 group' : 'SharePoint groups';
}

export function siteGroupRows(
  sites: ISiteRow[],
  groups: { [id: string]: IGroupSummary },
  accessOf: (s: ISiteRow) => ISiteAccess | undefined,
  layout: ExportLayout
): unknown[][] {
  const rows: unknown[][] = [];
  sites.forEach((s) => {
    const g = s.groupId ? groups[s.groupId] : undefined;
    const a = accessOf(s);
    const type = siteType(s.template, s.groupConnected);
    const notes: string[] = [];
    if (!s.groupId && s.relatedGroupId) {
      notes.push('Teams channel site: membership is the channel\u2019s');
    }
    if (!a) {
      notes.push('Not read (export stopped early)');
    } else if (a.error) {
      notes.push(a.error);
    } else if (a.truncated) {
      notes.push('Very large group: list truncated');
    }
    if (layout === 'site') {
      rows.push([
        s.url,
        s.title,
        type,
        sourceText(a),
        s.groupId ? 'Yes' : 'No',
        g ? g.displayName : '',
        g ? g.mail : '',
        g ? g.visibility : '',
        g ? (g.isTeam ? 'Yes' : 'No') : '',
        a && !a.error ? a.owners.length : '',
        a ? a.owners.map(personText).join('; ') : '',
        a && !a.error ? a.members.length : '',
        a ? a.members.map(personText).join('; ') : '',
        a && !a.error && a.source === 'sharepoint' ? a.visitors.length : '',
        a ? a.visitors.map(personText).join('; ') : '',
        a && a.admins ? a.admins.map(personText).join('; ') : '',
        notes.join('. ')
      ]);
      return;
    }
    if (!a || a.error) {
      rows.push([s.url, s.title, sourceText(a), g ? g.displayName : '', '', '', '', notes.join('. ')]);
      return;
    }
    const groupName = (role: 'owners' | 'members' | 'visitors'): string =>
      a.source === 'm365' ? (g ? g.displayName : 'Microsoft 365 group') : (role === 'owners' ? a.ownerGroup : role === 'members' ? a.memberGroup : a.visitorGroup) || '';
    const add = (people: IDirectoryObject[], role: string, group: string): void =>
      people.forEach((p) => rows.push([s.url, s.title, sourceText(a), group, role, p.displayName, p.mail || p.userPrincipalName || '', userType(p)]));
    add(a.owners, 'Owner', groupName('owners'));
    add(a.members, 'Member', groupName('members'));
    add(a.visitors, 'Visitor', groupName('visitors'));
    add(a.admins || [], 'Site collection admin', '');
    if (!a.owners.length && !a.members.length && !a.visitors.length && !(a.admins || []).length) {
      rows.push([s.url, s.title, sourceText(a), '', 'No owners or members', '', '', notes.join('. ')]);
    }
  });
  return rows;
}
