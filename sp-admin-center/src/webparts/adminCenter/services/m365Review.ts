import { IDirectoryObject, IGroupActivity, IGroupRef, IM365Group, IUserStatus } from './GraphApi';

/**
 * Pure rules behind the Microsoft 365 pages (group review, offboarding, compare members,
 * orphaned sites). No network: the views gather data, these decide what it means.
 */

export const DAY_MS = 86400000;

export function daysSince(d: Date | undefined, now = Date.now()): number | undefined {
  return d ? Math.max(0, Math.floor((now - d.getTime()) / DAY_MS)) : undefined;
}

export function isGuestPrincipal(p: IDirectoryObject): boolean {
  return (p.userType || '').toLowerCase() === 'guest' || /#ext#/i.test(p.userPrincipalName || '');
}

export type ReviewFlag = 'No owner' | 'One owner' | 'Guest owner' | 'Public with guests' | 'No description';

export const REVIEW_FLAGS: ReviewFlag[] = ['No owner', 'One owner', 'Guest owner', 'Public with guests', 'No description'];

/** What needs attention on a group. `activity` (usage report) adds the guest count. */
export function reviewFlags(g: IM365Group, activity?: IGroupActivity): ReviewFlag[] {
  const flags: ReviewFlag[] = [];
  const owners = (g.owners || []).filter((o) => o.type === 'user');
  if (owners.length === 0) {
    flags.push('No owner');
  } else if (owners.length === 1) {
    flags.push('One owner');
  }
  if (owners.some(isGuestPrincipal)) {
    flags.push('Guest owner');
  }
  if ((g.visibility || '').toLowerCase() === 'public' && activity && activity.guests > 0) {
    flags.push('Public with guests');
  }
  if (!(g.description || '').trim()) {
    flags.push('No description');
  }
  return flags;
}

// ---- offboarding -------------------------------------------------------------------

export interface IOffboardRow {
  group: IGroupRef;
  isOwner: boolean;
  isMember: boolean;
  /** Owners other than the person leaving (users only); undefined when not read. */
  otherOwners?: number;
}

/** One row per group the person owns or belongs to, owners first, sole-owner groups at the top. */
export function planOffboarding(owned: IGroupRef[], memberOf: IGroupRef[], owners: { [groupId: string]: IDirectoryObject[] | undefined }, leaverId: string): IOffboardRow[] {
  const rows: { [id: string]: IOffboardRow } = {};
  owned.forEach((g) => {
    rows[g.id] = { group: g, isOwner: true, isMember: false };
  });
  memberOf.forEach((g) => {
    rows[g.id] = rows[g.id] ? { ...rows[g.id], isMember: true } : { group: g, isOwner: false, isMember: true };
  });
  const me = leaverId.toLowerCase();
  Object.keys(rows).forEach((id) => {
    const list = owners[id];
    if (rows[id].isOwner && list) {
      rows[id].otherOwners = list.filter((o) => o.type === 'user' && o.id.toLowerCase() !== me).length;
    }
  });
  return Object.keys(rows)
    .map((k) => rows[k])
    .sort((a, b) => {
      const rank = (r: IOffboardRow): number => (r.isOwner ? (r.otherOwners === 0 ? 0 : 1) : 2);
      return rank(a) - rank(b) || a.group.displayName.localeCompare(b.group.displayName);
    });
}

// ---- compare members ------------------------------------------------------------------

export function compareMembers(a: IDirectoryObject[], b: IDirectoryObject[]): { onlyA: IDirectoryObject[]; onlyB: IDirectoryObject[]; both: IDirectoryObject[] } {
  const ids = (list: IDirectoryObject[]): { [id: string]: boolean } => {
    const m: { [id: string]: boolean } = {};
    list.forEach((p) => (m[p.id.toLowerCase()] = true));
    return m;
  };
  const inA = ids(a);
  const inB = ids(b);
  const byName = (x: IDirectoryObject, y: IDirectoryObject): number => x.displayName.localeCompare(y.displayName);
  return {
    onlyA: a.filter((p) => !inB[p.id.toLowerCase()]).sort(byName),
    onlyB: b.filter((p) => !inA[p.id.toLowerCase()]).sort(byName),
    both: a.filter((p) => inB[p.id.toLowerCase()]).sort(byName)
  };
}

// ---- orphaned sites -------------------------------------------------------------------

export type OwnerState = 'No owners' | 'All owners have left' | 'One active owner' | 'OK';

/** The key used to look a person up in Microsoft Graph: object id when known, else e-mail / UPN. */
export function personKey(p: IDirectoryObject, byId: boolean): string {
  return (byId ? p.id : p.userPrincipalName || p.mail || '').toLowerCase();
}

/**
 * Whether a site still has someone in charge. Only people count (groups such as
 * "Everyone" are not owners); an owner whose account is deleted or disabled has left.
 */
export function ownerState(owners: IDirectoryObject[], status: (p: IDirectoryObject) => IUserStatus | undefined): { state: OwnerState; active: IDirectoryObject[]; left: IDirectoryObject[] } {
  const people = owners.filter((o) => o.type === 'user');
  const left = people.filter((p) => {
    const s = status(p);
    return !!s && !s.unknown && (!s.exists || !s.enabled);
  });
  const active = people.filter((p) => left.indexOf(p) < 0);
  const state: OwnerState = people.length === 0 ? 'No owners' : active.length === 0 ? 'All owners have left' : active.length === 1 ? 'One active owner' : 'OK';
  return { state, active, left };
}
