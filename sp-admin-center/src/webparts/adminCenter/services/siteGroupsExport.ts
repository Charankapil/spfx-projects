import { ISiteRow } from '../models';
import { IDirectoryObject, IGroupMembership, IGroupSummary } from './GraphApi';
import { siteType } from './siteTypes';

/**
 * The "sites and their Microsoft 365 groups" export. Pure functions: the view
 * gathers sites (search), group names (one Graph listing) and memberships
 * (batched Graph reads), and these turn them into CSV rows.
 */

export type ExportLayout = 'site' | 'person';

export const SITE_HEADERS = ['Site URL', 'Site title', 'Site type', 'Microsoft 365 group', 'Group name', 'Group e-mail', 'Privacy', 'Teams', 'Owners (count)', 'Owners', 'Members (count)', 'Members', 'Note'];
export const PERSON_HEADERS = ['Site URL', 'Site title', 'Group name', 'Group e-mail', 'Role', 'Name', 'E-mail / UPN', 'User type'];

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

export function siteGroupRows(
  sites: ISiteRow[],
  groups: { [id: string]: IGroupSummary },
  memberships: { [id: string]: IGroupMembership | undefined },
  layout: ExportLayout
): unknown[][] {
  const rows: unknown[][] = [];
  sites.forEach((s) => {
    const g = s.groupId ? groups[s.groupId] : undefined;
    const m = s.groupId ? memberships[s.groupId] : undefined;
    const type = siteType(s.template, s.groupConnected);
    const note = !s.groupId
      ? s.relatedGroupId
        ? 'Teams channel site: membership is the channel’s, not the group’s'
        : ''
      : !m
        ? 'Membership not read (export stopped early)'
        : m.error || (m.truncated ? 'Very large group: list truncated' : '');
    if (layout === 'site') {
      rows.push([
        s.url,
        s.title,
        type,
        s.groupId ? 'Yes' : 'No',
        g ? g.displayName : '',
        g ? g.mail : '',
        g ? g.visibility : '',
        g ? (g.isTeam ? 'Yes' : 'No') : '',
        m && !m.error ? m.owners.length : '',
        m ? m.owners.map(personText).join('; ') : '',
        m && !m.error ? m.members.length : '',
        m ? m.members.map(personText).join('; ') : '',
        note
      ]);
      return;
    }
    if (!s.groupId || !m || m.error) {
      rows.push([s.url, s.title, g ? g.displayName : '', g ? g.mail : '', s.groupId ? '' : 'Not group-connected', '', '', note]);
      return;
    }
    m.owners.forEach((p) => rows.push([s.url, s.title, g ? g.displayName : '', g ? g.mail : '', 'Owner', p.displayName, p.mail || p.userPrincipalName || '', userType(p)]));
    m.members.forEach((p) => rows.push([s.url, s.title, g ? g.displayName : '', g ? g.mail : '', 'Member', p.displayName, p.mail || p.userPrincipalName || '', userType(p)]));
    if (!m.owners.length && !m.members.length) {
      rows.push([s.url, s.title, g ? g.displayName : '', g ? g.mail : '', 'No owners or members', '', '', note]);
    }
  });
  return rows;
}
