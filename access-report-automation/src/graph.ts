import { TokenProvider } from './auth.js';
import { Http, HttpError } from './http.js';
import { groupClaims, Identity, implicitIdentities, Principal, PrincipalKind, toLogin } from './permissions.js';

/** Microsoft Graph, app-only: the list of sites, and who the people and groups in the report are. */

export interface SiteRef {
  url: string;
  title: string;
}

export interface PrincipalRequest {
  type: 'user' | 'group';
  /** User: e-mail or UPN. Group: object id, e-mail or exact display name. */
  id: string;
}

export interface UnresolvedPrincipal {
  input: string;
  type: string;
  reason: string;
}

interface RawUser {
  id: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  userType?: string;
  accountEnabled?: boolean;
}

interface RawGroup {
  '@odata.type'?: string;
  id: string;
  displayName?: string;
  mail?: string;
  groupTypes?: string[];
  securityEnabled?: boolean;
}

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function groupKind(g: RawGroup): PrincipalKind {
  return (g.groupTypes || []).indexOf('Unified') >= 0 ? 'm365' : 'security';
}

function q(v: string): string {
  return v.replace(/'/g, "''");
}

export class Graph {
  private readonly base: string;

  constructor(private http: Http, private tokens: TokenProvider, graphBase: string) {
    this.base = graphBase.replace(/\/+$/, '');
  }

  private async get<T>(pathOrUrl: string, headers: Record<string, string> = {}): Promise<T> {
    const url = /^https?:\/\//i.test(pathOrUrl) ? pathOrUrl : this.base + pathOrUrl;
    if (url.toLowerCase().indexOf(this.base.toLowerCase() + '/') !== 0) {
      throw new Error('Refusing to follow a Graph link to another address.');
    }
    const call = async (force: boolean): Promise<T> => {
      const token = await this.tokens.getToken(this.base, force);
      return (await this.http.request<T>(url, { headers: { Accept: 'application/json', Authorization: `Bearer ${token}`, ...headers } })).data;
    };
    try {
      return await call(false);
    } catch (e) {
      if (e instanceof HttpError && e.status === 401) {
        return call(true);
      }
      throw e;
    }
  }

  private async getAll<T>(path: string, headers: Record<string, string> = {}, onPage?: (n: number) => void, max = 1000000): Promise<T[]> {
    const out: T[] = [];
    let next: string | undefined = path;
    for (let page = 0; next && page < 5000 && out.length < max; page++) {
      const j: { value?: T[]; '@odata.nextLink'?: string } = await this.get(next, headers);
      (j.value || []).forEach((r) => out.push(r));
      if (onPage) {
        onPage(out.length);
      }
      next = j['@odata.nextLink'];
    }
    return out;
  }

  /** Every site collection in the tenant (all geographies), OneDrive left out unless asked for. */
  public async listSites(includeOneDrive: boolean, onPage?: (n: number) => void): Promise<SiteRef[]> {
    type Row = { webUrl?: string; displayName?: string; name?: string; isPersonalSite?: boolean };
    let rows: Row[];
    try {
      rows = await this.getAll<Row>('/v1.0/sites/getAllSites?$select=id,webUrl,displayName,name,isPersonalSite&$top=999', {}, onPage);
    } catch (e) {
      if (!(e instanceof HttpError && e.status === 400)) {
        throw e;
      }
      // The query options were refused: ask for the plain list.
      rows = await this.getAll<Row>('/v1.0/sites/getAllSites', {}, onPage);
    }
    const seen = new Set<string>();
    const out: SiteRef[] = [];
    for (const r of rows) {
      const url = String(r.webUrl || '').replace(/\/+$/, '');
      if (!/^https?:\/\//i.test(url)) {
        continue;
      }
      const personal = r.isPersonalSite === true || /-my\.sharepoint\.[a-z.]+\/personal\//i.test(url);
      if (personal && !includeOneDrive) {
        continue;
      }
      const key = url.toLowerCase();
      if (!seen.has(key)) {
        seen.add(key);
        out.push({ url, title: String(r.displayName || r.name || url) });
      }
    }
    out.sort((a, b) => (a.url.toLowerCase() < b.url.toLowerCase() ? -1 : 1));
    return out;
  }

  private async findUser(id: string): Promise<RawUser | undefined> {
    const select = '$select=id,displayName,mail,userPrincipalName,userType,accountEnabled';
    try {
      return await this.get<RawUser>(`/v1.0/users/${encodeURIComponent(id)}?${select}`);
    } catch (e) {
      // 400: the text is not something Graph accepts as a key (e.g. it contains a space).
      if (!(e instanceof HttpError && (e.status === 404 || e.status === 400))) {
        throw e;
      }
    }
    const filter = encodeURIComponent(`mail eq '${q(id)}' or otherMails/any(m:m eq '${q(id)}')`);
    const res = await this.get<{ value?: RawUser[] }>(`/v1.0/users?$filter=${filter}&${select}&$count=true&$top=5`, { ConsistencyLevel: 'eventual' });
    const found = res.value || [];
    if (found.length > 1) {
      throw new Error(`${found.length} accounts share this e-mail address; use the sign-in name (UPN) instead`);
    }
    return found[0];
  }

  private async findGroup(id: string): Promise<RawGroup | undefined> {
    const select = '$select=id,displayName,mail,groupTypes,securityEnabled';
    if (GUID.test(id)) {
      try {
        return await this.get<RawGroup>(`/v1.0/groups/${id.toLowerCase()}?${select}`);
      } catch (e) {
        if (e instanceof HttpError && e.status === 404) {
          return undefined;
        }
        throw e;
      }
    }
    const filter = encodeURIComponent(`displayName eq '${q(id)}' or mail eq '${q(id)}'`);
    const res = await this.get<{ value?: RawGroup[] }>(`/v1.0/groups?$filter=${filter}&${select}&$top=5`);
    const found = res.value || [];
    if (found.length > 1) {
      throw new Error(`${found.length} groups are called this; use the group's object id instead`);
    }
    return found[0];
  }

  /** Groups a directory object belongs to, nested ones included. */
  private async groupsOf(kind: 'users' | 'groups', objectId: string): Promise<RawGroup[]> {
    const rows = await this.getAll<RawGroup>(`/v1.0/${kind}/${objectId}/transitiveMemberOf?$select=id,displayName,groupTypes,securityEnabled&$top=999`);
    return rows.filter((r) => !r['@odata.type'] || /group$/i.test(r['@odata.type']));
  }

  /** Microsoft 365 groups a person owns (owners are not always members). Best effort. */
  private async ownedGroups(userId: string): Promise<RawGroup[]> {
    try {
      const rows = await this.getAll<RawGroup>(`/v1.0/users/${userId}/ownedObjects?$select=id,displayName,groupTypes,securityEnabled&$top=999`);
      return rows.filter((r) => /group$/i.test(r['@odata.type'] || '') && groupKind(r) === 'm365');
    } catch {
      return [];
    }
  }

  /**
   * Turns the admin's list into principals with every claim that can give them access.
   * expandGroups adds the groups each one belongs to (needs GroupMember.Read.All).
   */
  public async resolvePrincipals(requests: PrincipalRequest[], expandGroups: boolean): Promise<{ principals: Principal[]; unresolved: UnresolvedPrincipal[]; warnings: string[] }> {
    const principals: Principal[] = [];
    const unresolved: UnresolvedPrincipal[] = [];
    const warnings: string[] = [];
    const seen = new Set<string>();
    let expandFailed = false;

    const memberships = async (kind: 'users' | 'groups', id: string): Promise<Identity[]> => {
      if (!expandGroups || expandFailed) {
        return [];
      }
      try {
        const groups = await this.groupsOf(kind, id);
        const out: Identity[] = [];
        groups.forEach((g) => {
          const name = g.displayName || g.id;
          const k = groupKind(g);
          if (k === 'm365') {
            out.push({ login: groupClaims('m365', g.id)[0], via: name });
          } else if (g.securityEnabled !== false) {
            out.push({ login: groupClaims('security', g.id)[0], via: name });
          }
        });
        return out;
      } catch (e) {
        if (e instanceof HttpError && (e.status === 403 || e.status === 401)) {
          expandFailed = true;
          warnings.push('Group memberships could not be read (the app needs the Microsoft Graph application permission GroupMember.Read.All). The report relies on SharePoint\'s own effective-permission check and cannot name the group that gives access.');
          return [];
        }
        throw e;
      }
    };

    for (const r of requests) {
      const input = r.id.trim();
      try {
        if (r.type === 'user') {
          const u = await this.findUser(input);
          if (!u) {
            unresolved.push({ input, type: r.type, reason: 'Not found in the directory (deleted, or a typo?)' });
            continue;
          }
          const key = `user:${u.id.toLowerCase()}`;
          if (seen.has(key)) {
            continue;
          }
          seen.add(key);
          const upn = u.userPrincipalName || u.mail || input;
          const guest = String(u.userType || '').toLowerCase() === 'guest';
          const identities = implicitIdentities(guest).concat(await memberships('users', u.id));
          if (expandGroups && !expandFailed) {
            (await this.ownedGroups(u.id)).forEach((g) => identities.push({ login: groupClaims('m365', g.id)[1], via: `${g.displayName || g.id} (owners)` }));
          }
          const shown = u.mail || upn;
          principals.push({
            key,
            kind: 'user',
            label: `${u.displayName || shown} <${shown}>${guest ? ' (guest)' : ''}${u.accountEnabled === false ? ' (sign-in blocked)' : ''}`,
            input,
            login: toLogin(upn),
            guest,
            identities
          });
        } else {
          const g = await this.findGroup(input);
          if (!g) {
            unresolved.push({ input, type: r.type, reason: 'Not found in the directory (deleted, or a typo?)' });
            continue;
          }
          const key = `group:${g.id.toLowerCase()}`;
          if (seen.has(key)) {
            continue;
          }
          seen.add(key);
          const kind = groupKind(g);
          const own: Identity[] = groupClaims(kind, g.id).map((login, i) => ({ login, via: i === 1 ? 'owners of this group' : '' }));
          principals.push({
            key,
            kind,
            label: `${g.displayName || g.id}${g.mail ? ` <${g.mail}>` : ''} (${kind === 'm365' ? 'Microsoft 365 group' : 'security group'})`,
            input,
            guest: false,
            identities: own.concat(await memberships('groups', g.id))
          });
        }
      } catch (e) {
        if (e instanceof HttpError && (e.status === 403 || e.status === 401)) {
          throw new Error(`Microsoft Graph refused to look up ${r.type}s (${e.message}). The app needs the application permissions User.Read.All and GroupMember.Read.All, with admin consent.`);
        }
        if (e instanceof HttpError) {
          throw e;
        }
        unresolved.push({ input, type: r.type, reason: (e as Error).message });
      }
    }
    return { principals, unresolved, warnings };
  }
}
