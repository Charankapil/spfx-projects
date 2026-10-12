/**
 * Read-only locks, shared by the web part (which applies them) and the banner
 * extension (which shows their message). Pure functions only: no network.
 *
 * A lock makes a site, library/list or folder read-only by switching every
 * edit-capable permission on that object to Read, after saving the exact
 * original permissions so that unlocking restores them.
 */

export type LockScope = 'web' | 'list' | 'folder';

/** What the banner needs to know, stored in the extension's settings on the web. */
export interface IBannerLock {
  id: string;
  scope: LockScope;
  /** Server-relative path of the locked web, library root folder or folder (no trailing slash). */
  path: string;
  title: string;
  message: string;
  /** Epoch seconds. */
  at: number;
}

export interface IBinding {
  id: number;
  name: string;
  /** SharePoint RoleTypeKind: 0 none/custom, 1 guest (limited access), 2 reader, 3 contributor, 4 web designer, 5 administrator, 6 editor. */
  kind: number;
  hidden: boolean;
}

export interface IAssignment {
  principalId: number;
  title: string;
  login: string;
  principalType: number;
  bindings: IBinding[];
}

export interface IPlannedChange {
  principalId: number;
  title: string;
  /** Permission level names before / after, for the preview. */
  before: string[];
  after: string[];
  add: number[];
  remove: number[];
}

export const READER_KIND = 2;

/** Limited Access and the hidden system levels are managed by SharePoint itself and never touched. */
export function isSystemBinding(b: IBinding): boolean {
  return b.hidden || b.kind === 1;
}

/**
 * Changes that make an object read-only: every principal (other than the ones to keep)
 * whose permission levels include anything other than Read gets Read and loses the rest.
 * Read is added before anything is removed, so nobody loses access in between.
 */
export function planReadOnly(assignments: IAssignment[], readId: number, readName: string, keepPrincipalIds: number[]): IPlannedChange[] {
  const out: IPlannedChange[] = [];
  assignments.forEach((a) => {
    if (keepPrincipalIds.indexOf(a.principalId) >= 0) {
      return;
    }
    const levels = a.bindings.filter((b) => !isSystemBinding(b));
    const writable = levels.filter((b) => b.kind !== READER_KIND);
    if (writable.length === 0) {
      return;
    }
    const hasRead = levels.some((b) => b.kind === READER_KIND || b.id === readId);
    out.push({
      principalId: a.principalId,
      title: a.title,
      before: levels.map((b) => b.name),
      after: levels.filter((b) => b.kind === READER_KIND).map((b) => b.name).concat(hasRead ? [] : [readName]),
      add: hasRead ? [] : [readId],
      remove: writable.map((b) => b.id)
    });
  });
  return out;
}

export interface ISavedAssignment {
  principalId: number;
  title: string;
  roleDefIds: number[];
}

/**
 * Changes that undo a lock on an object that had its own permissions before the lock:
 * every saved principal gets its original levels back, and Read is removed again where
 * the lock added it. Grants made by someone else after the lock are left alone.
 */
export function planRestore(saved: ISavedAssignment[], current: IAssignment[], readId: number): IPlannedChange[] {
  const out: IPlannedChange[] = [];
  saved.forEach((s) => {
    const now = current.filter((c) => c.principalId === s.principalId)[0];
    const nowIds = now ? now.bindings.filter((b) => !isSystemBinding(b)).map((b) => b.id) : [];
    const add = s.roleDefIds.filter((id) => nowIds.indexOf(id) < 0);
    const remove = s.roleDefIds.indexOf(readId) < 0 && nowIds.indexOf(readId) >= 0 ? [readId] : [];
    if (add.length || remove.length) {
      out.push({ principalId: s.principalId, title: s.title, before: [], after: [], add, remove });
    }
  });
  return out;
}

function norm(path: string): string {
  let p = path;
  try {
    p = decodeURIComponent(path);
  } catch {
    /* keep as is */
  }
  return p.replace(/\/+$/, '').toLowerCase();
}

function within(path: string, root: string): boolean {
  const p = norm(path);
  const r = norm(root);
  return !!r && (p === r || p.indexOf(r + '/') === 0);
}

/**
 * The lock whose message applies to the page at `pathname` (+ `search`), most specific first.
 * Library views put the folder being browsed in the id / RootFolder query parameter.
 */
export function matchLock(locks: IBannerLock[], pathname: string, search: string): IBannerLock | undefined {
  const candidates: string[] = [pathname];
  const q = /[?&](?:id|RootFolder)=([^&#]+)/i.exec(search || '');
  if (q) {
    try {
      candidates.unshift(decodeURIComponent(q[1].replace(/\+/g, ' ')));
    } catch {
      candidates.unshift(q[1]);
    }
  }
  const hits = locks.filter((l) => candidates.some((c) => within(c, l.path)));
  hits.sort((a, b) => norm(b.path).length - norm(a.path).length);
  return hits[0];
}
