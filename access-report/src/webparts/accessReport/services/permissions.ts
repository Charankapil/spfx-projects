/**
 * Who a report is about, and what a SharePoint permission mask means.
 *
 * A subject is stored by its SharePoint login name (claim), which is all SharePoint REST
 * needs to answer for it - no directory lookup, so no Microsoft Graph:
 *   i:0#.f|membership|ann@contoso.com                      a person (guests too)
 *   c:0t.c|tenant|<object id>                               a security group
 *   c:0o.c|federateddirectoryclaimprovider|<object id>      a Microsoft 365 group (members)
 *   c:0o.c|federateddirectoryclaimprovider|<object id>_o    ... its owners
 *   c:0-.f|rolemanager|spo-grid-all-users/<tenant id>       Everyone except external users
 *   c:0(.s|true                                             Everyone
 */

export interface ISubject {
  /** Shown in the report's "Who" column. */
  label: string;
  /** SharePoint login name (claim). */
  login: string;
}

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const M365_CLAIM = '|federateddirectoryclaimprovider|';

/** E-mail / UPN -> membership claim; a claim is kept as it is. */
export function toLogin(input: string): string {
  const v = input.trim();
  if (v.indexOf('|') >= 0) {
    return v;
  }
  return 'i:0#.f|membership|' + v.toLowerCase();
}

/** People are identity claims ("i:..."); everything else is answered from the site's permissions. */
export function isPersonLogin(login: string): boolean {
  return /^i:/i.test(login.trim());
}

export function isM365Login(login: string): boolean {
  return login.toLowerCase().indexOf(M365_CLAIM) >= 0;
}

export function isGuid(value: string): boolean {
  return GUID.test(value.trim());
}

export function securityGroupLogin(objectId: string): string {
  return 'c:0t.c|tenant|' + objectId.trim().toLowerCase();
}

export function m365GroupLogin(objectId: string): string {
  return 'c:0o.c' + M365_CLAIM + objectId.trim().toLowerCase();
}

export function everyoneExceptExternalLogin(tenantId: string): string {
  return 'c:0-.f|rolemanager|spo-grid-all-users/' + tenantId.trim().toLowerCase();
}

export const EVERYONE_LOGIN = 'c:0(.s|true';

/**
 * The login as it is stored: trimmed, and a Microsoft 365 group always as its members claim
 * (the owners claim is checked alongside it, so storing "..._o" would look for "..._o_o").
 */
export function normaliseLogin(input: string): string {
  const v = toLogin(input);
  return isM365Login(v) ? v.replace(/_o$/i, '') : v;
}

/** The claims a group is known by inside a site: one, or members + owners for a Microsoft 365 group. */
export function groupClaims(login: string): string[] {
  const v = login.trim();
  return isM365Login(v) ? [v, v + '_o'] : [v];
}

export function subjectKind(login: string): string {
  const v = login.trim().toLowerCase();
  if (isPersonLogin(v)) {
    return /#ext#/.test(v) ? 'Guest' : 'Person';
  }
  if (isM365Login(v)) {
    return 'Microsoft 365 group';
  }
  if (v.indexOf('c:0t.c|tenant|') === 0) {
    return 'Security group';
  }
  if (v === EVERYONE_LOGIN || v.indexOf('spo-grid-all-users') >= 0) {
    return 'Everyone claim';
  }
  return 'Group / claim';
}

export type ManualType = 'person' | 'security' | 'm365' | 'claim';

/** A typed e-mail, object id or claim as a subject, or the reason it is not one. */
export function manualSubject(type: ManualType, value: string, label: string): ISubject | string {
  const v = value.trim();
  if (!v) {
    return 'Enter a value.';
  }
  if (type === 'person') {
    if (v.indexOf('|') < 0 && !/^[^@\s]+@[^@\s]+$/.test(v)) {
      return 'Enter the sign-in name (e-mail address) of the person.';
    }
    return { login: normaliseLogin(v), label: label.trim() || v };
  }
  if (type === 'claim') {
    return v.indexOf('|') < 0 ? 'A claim looks like c:0t.c|tenant|<id>.' : { login: normaliseLogin(v), label: label.trim() || v };
  }
  if (!isGuid(v)) {
    return 'Enter the group’s object id (a GUID from Entra ID).';
  }
  return { login: type === 'm365' ? m365GroupLogin(v) : securityGroupLogin(v), label: label.trim() || `${type === 'm365' ? 'Microsoft 365 group' : 'Security group'} ${v.toLowerCase()}` };
}

export const NO_ACCESS = 'No access';

/**
 * A permission mask (the Low / High halves SharePoint returns) as one of five plain levels.
 * PermissionKind numbers used: ViewListItems 1, AddListItems 2, EditListItems 3,
 * DeleteListItems 4, AddAndCustomizePages 19, ManagePermissions 26; FullMask = every bit.
 * The built-in levels land where their names say: Full Control, Design (19) -> "Design /
 * manage", Edit and Contribute -> "Edit / contribute", Read and View Only -> "Read", and
 * Limited Access (no ViewListItems) -> no access.
 *
 * The Power Automate flow computes exactly this with integer arithmetic (flow/definition.ts);
 * the tests check the two against each other.
 */
export function permissionLevel(low: number, high: number): string {
  const has = (kind: number): boolean => Math.floor(low / Math.pow(2, kind - 1)) % 2 === 1;
  if (low >= 4294967295 && high >= 2147483647) {
    return 'Full control';
  }
  if (has(19) || has(26)) {
    return 'Design / manage';
  }
  if (has(2) && has(3) && has(4)) {
    return 'Edit / contribute';
  }
  if (has(1)) {
    return 'Read';
  }
  return NO_ACCESS;
}
