export type ScanPermission = 'owners' | 'everyone' | 'people';

export interface IScanUser {
  email?: string;
  loginName?: string;
}

/** Splits the "allowed people" setting (one address per line, or separated by ; or ,) into lowercase entries. */
export function parseAllowedPeople(text: string | undefined): string[] {
  return (text || '')
    .split(/[\n;,]/)
    .map((n) => n.trim().toLowerCase())
    .filter((n) => n.length > 0);
}

/** True when the signed-in user is named in the list, by email or by sign-in name (`i:0#.f|membership|name@contoso.com`). */
export function isAllowedPerson(user: IScanUser, allowed: string[]): boolean {
  const email = (user.email || '').trim().toLowerCase();
  const login = (user.loginName || '').trim().toLowerCase();
  const loginTail = login.indexOf('|') >= 0 ? login.substring(login.lastIndexOf('|') + 1) : login;
  return allowed.some((raw) => {
    const entry = raw.trim().toLowerCase();
    return entry.length > 0 && (entry === email || entry === login || entry === loginTail);
  });
}

/**
 * Who may run a scan. Only the "people" mode ignores site ownership: it is the
 * way to reserve scanning for named administrators. It is a convenience gate in
 * the page, not a security boundary: results and scans only ever use the
 * signed-in user's own SharePoint permissions.
 */
export function canRunScan(permission: ScanPermission, isOwner: boolean, isListed: boolean): boolean {
  if (permission === 'everyone') {
    return true;
  }
  if (permission === 'people') {
    return isListed;
  }
  return isOwner;
}
