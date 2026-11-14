/**
 * Pure permission logic: no network, easy to test.
 *
 * Ported from the Access report in the SharePoint Admin Center web part
 * (sp-admin-center/src/webparts/adminCenter/services/AccessScan.ts) and extended so an
 * unattended run can also say *how* a person gets their access (which Entra ID group,
 * which SharePoint group) instead of only the effective level.
 */

export type PrincipalKind = 'user' | 'security' | 'm365';

/** A sign-in name SharePoint may know a principal by, and the group it stands for ('' = the principal itself). */
export interface Identity {
  login: string;
  via: string;
  /** A claim everyone holds (Everyone, Everyone except external users), not a directory group. */
  implicit?: boolean;
}

export interface Principal {
  /** Stable across runs: "user:<object id>" or "group:<object id>". */
  key: string;
  kind: PrincipalKind;
  /** Shown in the report. */
  label: string;
  /** What the admin typed in the principals list. */
  input: string;
  /** The person's own SharePoint login (users only). */
  login?: string;
  guest: boolean;
  /** Claims that give this principal access: its own (groups), and the groups it belongs to. */
  identities: Identity[];
}

export interface Bits {
  low: number;
  high: number;
}

export const NO_BITS: Bits = { low: 0, high: 0 };
export const FULL_BITS: Bits = { low: 4294967295, high: 2147483647 };

export function orBits(a: Bits, b: Bits): Bits {
  // >>> 0 keeps the 32-bit halves unsigned.
  return { low: (a.low | b.low) >>> 0, high: (a.high | b.high) >>> 0 };
}

export function bitsOf(bp: { Low?: unknown; High?: unknown } | undefined | null): Bits {
  if (!bp) {
    return NO_BITS;
  }
  return { low: Number(bp.Low) >>> 0 || 0, high: Number(bp.High) >>> 0 || 0 };
}

/** SP.PermissionKind numbers (1-based bit positions across Low then High). */
export const PermissionKind = {
  viewListItems: 1,
  addListItems: 2,
  editListItems: 3,
  deleteListItems: 4,
  openItems: 6,
  manageLists: 12,
  addAndCustomizePages: 19,
  applyStyleSheets: 21,
  managePermissions: 26,
  manageWeb: 31
} as const;

export function hasPermission(b: Bits, kind: number): boolean {
  return kind <= 32 ? Math.floor(b.low / Math.pow(2, kind - 1)) % 2 === 1 : Math.floor(b.high / Math.pow(2, kind - 33)) % 2 === 1;
}

export const NO_ACCESS = 'No access';

/** The access levels, weakest first (used for sorting and for "more / less access" in the changes sheet). */
export const LEVELS = [NO_ACCESS, 'View only', 'Read', 'Contribute', 'Edit', 'Design', 'Full control'];

/**
 * A permission mask in plain words. Limited Access (opening a single shared item) is not
 * access to the site or library itself, so it counts as no access.
 */
export function permissionLevel(b: Bits): string {
  const has = (k: number): boolean => hasPermission(b, k);
  const K = PermissionKind;
  if (has(K.managePermissions) && has(K.manageWeb)) {
    return 'Full control';
  }
  const write = has(K.addListItems) && has(K.editListItems) && has(K.deleteListItems);
  if (write && has(K.addAndCustomizePages) && has(K.applyStyleSheets)) {
    return 'Design';
  }
  if (write && has(K.manageLists)) {
    return 'Edit';
  }
  if (write) {
    return 'Contribute';
  }
  if (has(K.viewListItems)) {
    return has(K.openItems) ? 'Read' : 'View only';
  }
  return NO_ACCESS;
}

/** Masks of the built-in levels, for role assignments that come back without their mask. */
const ROLE_TYPE_BITS: Record<number, Bits> = {
  2: { low: 138612833, high: 176 }, // Reader
  3: { low: 1011028719, high: 432 }, // Contributor
  4: { low: 1012866047, high: 432 }, // Web designer
  5: FULL_BITS, // Administrator
  6: { low: 1011030767, high: 432 } // Editor
};

export function toLogin(input: string): string {
  const v = input.trim();
  return v.indexOf('|') >= 0 ? v : 'i:0#.f|membership|' + v.toLowerCase();
}

/** Claims a group is known by inside SharePoint (members, and for Microsoft 365 groups also owners). */
export function groupClaims(kind: PrincipalKind, objectId: string): string[] {
  const id = objectId.trim().toLowerCase();
  if (kind === 'security') {
    return [`c:0t.c|tenant|${id}`];
  }
  return [`c:0o.c|federateddirectoryclaimprovider|${id}`, `c:0o.c|federateddirectoryclaimprovider|${id}_o`];
}

export const EVERYONE_LOGIN = 'c:0(.s|true';
export const EEEU_PREFIX = 'c:0-.f|rolemanager|spo-grid-all-users/';
export const ALL_USERS_WINDOWS = 'c:0!.s|windows';

/** Claims every signed-in person holds without being a member of anything. */
export function implicitIdentities(guest: boolean): Identity[] {
  const out: Identity[] = [{ login: EVERYONE_LOGIN, via: 'Everyone', implicit: true }];
  if (!guest) {
    out.push({ login: EEEU_PREFIX, via: 'Everyone except external users', implicit: true });
    out.push({ login: ALL_USERS_WINDOWS, via: 'All users', implicit: true });
  }
  return out;
}

/** A principal's foothold in one site collection: a site user entry that is the principal, or a group it belongs to. */
export interface Link {
  /** Id of the entry in the site's user list. */
  id: number;
  /** '' when the entry is the principal itself, else the group's name. */
  via: string;
  isSiteAdmin: boolean;
  /** `via` is a claim everyone holds, not a directory group. */
  implicit?: boolean;
}

export interface SpGroupLink {
  id: number;
  title: string;
  via: string;
  implicit?: boolean;
}

/** "Group: Finance" for a directory group, "Everyone" for a claim everyone holds. */
function viaText(l: { via: string; implicit?: boolean }): string {
  return l.implicit ? l.via : `Group: ${l.via}`;
}

export interface RawRoleDefinition {
  Name?: string;
  RoleTypeKind?: number;
  BasePermissions?: { Low?: unknown; High?: unknown };
}

export interface RawAssignment {
  PrincipalId?: number;
  RoleDefinitionBindings?: RawRoleDefinition[] | { results?: RawRoleDefinition[] };
}

export interface Explanation {
  bits: Bits;
  /** Permission level names as the site calls them ("Contribute", "Marketing editors"...). */
  levels: string[];
  /** One entry per route: "Direct", "SharePoint group: X", "Group: G -> SharePoint group: X", "Everyone -> SharePoint group: X". */
  how: string[];
}

function bindings(a: RawAssignment): RawRoleDefinition[] {
  const b = a.RoleDefinitionBindings;
  if (Array.isArray(b)) {
    return b;
  }
  return b && Array.isArray(b.results) ? b.results : [];
}

function roleBits(r: RawRoleDefinition): Bits {
  const own = bitsOf(r.BasePermissions);
  if (own.low || own.high) {
    return own;
  }
  return ROLE_TYPE_BITS[Number(r.RoleTypeKind)] || NO_BITS;
}

const uniq = (xs: string[]): string[] => xs.filter((x, i) => xs.indexOf(x) === i);

/**
 * What a principal has on one object, read from the object's role assignments: grants to
 * the principal (or a group it belongs to), and grants to SharePoint groups those are in.
 * Limited Access is left out. undefined = nothing.
 */
export function explainAssignments(assignments: RawAssignment[], links: Link[], spGroups: SpGroupLink[]): Explanation | undefined {
  let bits = NO_BITS;
  const levels: string[] = [];
  const how: string[] = [];
  for (const a of assignments) {
    const pid = Number(a.PrincipalId);
    const roles = bindings(a).filter((r) => r.RoleTypeKind !== 1 && !/limited access/i.test(String(r.Name || '')));
    if (!roles.length) {
      continue;
    }
    const routes: string[] = [];
    links.filter((l) => l.id === pid).forEach((l) => routes.push(l.via ? viaText(l) : 'Direct'));
    spGroups.filter((g) => g.id === pid).forEach((g) => routes.push(g.via ? `${viaText(g)} -> SharePoint group: ${g.title}` : `SharePoint group: ${g.title}`));
    if (!routes.length) {
      continue;
    }
    roles.forEach((r) => {
      bits = orBits(bits, roleBits(r));
      levels.push(String(r.Name || 'Unnamed level'));
    });
    routes.forEach((r) => how.push(r));
  }
  if (!how.length) {
    return undefined;
  }
  return { bits, levels: uniq(levels), how: uniq(how) };
}

/** Site collection administrators have full control of everything in it. */
export function adminExplanation(links: Link[]): Explanation | undefined {
  const admins = links.filter((l) => l.isSiteAdmin);
  if (!admins.length) {
    return undefined;
  }
  return {
    bits: FULL_BITS,
    levels: ['Site collection administrator'],
    how: uniq(admins.map((l) => (l.via ? `Site collection administrator (${l.implicit ? l.via : `group: ${l.via}`})` : 'Site collection administrator')))
  };
}

export function mergeExplanations(a: Explanation | undefined, b: Explanation | undefined): Explanation | undefined {
  if (!a || !b) {
    return a || b;
  }
  return { bits: orBits(a.bits, b.bits), levels: uniq(a.levels.concat(b.levels)), how: uniq(a.how.concat(b.how)) };
}
