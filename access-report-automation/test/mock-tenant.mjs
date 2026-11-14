// A pretend Microsoft 365 tenant on localhost: Entra ID sign-in, Microsoft Graph and SharePoint
// REST, backed by a small permission model. It shares no code with src/, so that what the tool
// reports can be compared with what this model says (see `oracle`).
import { createPublicKey, createVerify, createHash, X509Certificate, randomUUID } from 'node:crypto';
import { createServer } from 'node:http';

export const ROLES = {
  'Full Control': { id: 1073741829, kind: 5, low: 4294967295, high: 2147483647 },
  Design: { id: 1073741828, kind: 4, low: 1012866047, high: 432 },
  Edit: { id: 1073741830, kind: 6, low: 1011030767, high: 432 },
  Contribute: { id: 1073741827, kind: 3, low: 1011028719, high: 432 },
  Read: { id: 1073741826, kind: 2, low: 138612833, high: 176 },
  'View Only': { id: 1073741924, kind: 0, low: 138612801, high: 176 },
  'Limited Access': { id: 1073741825, kind: 1, low: 134287360, high: 0 },
  // A custom level: read plus add (no edit / delete).
  'Drop box': { id: 1073741999, kind: 0, low: 138612835, high: 176 }
};

const TENANT_GUID = '11111111-2222-3333-4444-555555555555';

function bit(low, high, kind) {
  return kind <= 32 ? Math.floor(low / 2 ** (kind - 1)) % 2 === 1 : Math.floor(high / 2 ** (kind - 33)) % 2 === 1;
}

/** Independent of src/permissions.ts on purpose. */
export function levelOf(low, high) {
  if (bit(low, high, 26) && bit(low, high, 31)) return 'Full control';
  const write = bit(low, high, 2) && bit(low, high, 3) && bit(low, high, 4);
  if (write && bit(low, high, 19) && bit(low, high, 21)) return 'Design';
  if (write && bit(low, high, 12)) return 'Edit';
  if (write) return 'Contribute';
  if (bit(low, high, 1)) return bit(low, high, 6) ? 'Read' : 'View only';
  return 'No access';
}

const b64url = (b) => Buffer.from(b).toString('base64url');

export class MockTenant {
  /**
   * @param model { users, groups, sites } - see test/fixtures in the tests for the shape.
   * @param opts  { certPem, clientId, tenantId, graphRoles, spRoles, pageSize }
   */
  constructor(model, opts = {}) {
    this.model = model;
    this.opts = { clientId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee', tenantId: TENANT_GUID, graphRoles: ['Sites.Read.All', 'User.Read.All', 'GroupMember.Read.All'], spRoles: ['Sites.FullControl.All'], pageSize: 3, ...opts };
    this.files = new Map(); // server-relative path -> Buffer
    this.folders = new Set();
    this.lockedFiles = new Set();
    this.uploads = new Map();
    this.requests = [];
    this.faults = []; // (req) => { status, headers?, body?, drop?, delayMs? } | undefined
    this.validTokens = new Set();
    this.tokensIssued = 0;
    this.rejectPrincipalTypeFilter = false;
    this.staleGroupUsers = new Set(); // UPNs whose group memberships SharePoint "does not know"
    this.unknownUsers = new Set(); // UPNs SharePoint cannot evaluate at all (answers 404)
    this.everyoneExcludesGuests = false;
    this.oidcToken = 'oidc-' + randomUUID();
    this.oidcRequestToken = 'runner-' + randomUUID();
    this.server = createServer((req, res) => this.handle(req, res).catch((e) => this.send(res, 500, { error: { message: 'mock crashed: ' + e.stack } })));
  }

  async start() {
    await new Promise((r) => this.server.listen(0, '127.0.0.1', r));
    this.origin = `http://127.0.0.1:${this.server.address().port}`;
    this.index();
    return this;
  }

  async stop() {
    this.server.closeAllConnections();
    await new Promise((r) => this.server.close(r));
  }

  get loginBase() {
    return this.origin + '/login';
  }
  get graphBase() {
    return this.origin + '/graph';
  }

  // ---- the permission model ------------------------------------------------------------------

  claimOf(ref) {
    const [type, ...rest] = ref.split(':');
    const v = rest.join(':');
    switch (type) {
      case 'user':
        return 'i:0#.f|membership|' + v.toLowerCase();
      case 'group': {
        const g = this.model.groups.find((x) => x.id === v);
        if (!g) throw new Error('unknown group ' + v);
        return g.unified ? `c:0o.c|federateddirectoryclaimprovider|${v}` : `c:0t.c|tenant|${v}`;
      }
      case 'owners':
        return `c:0o.c|federateddirectoryclaimprovider|${v}_o`;
      case 'everyone':
        return 'c:0(.s|true';
      case 'eeeu':
        return `c:0-.f|rolemanager|spo-grid-all-users/${this.opts.tenantId}`;
      default:
        throw new Error('bad principal ref ' + ref);
    }
  }

  /** (Re)builds each site's user list and ids. Call after changing the model. */
  index() {
    for (const site of this.model.sites) {
      site.url = this.origin + site.path;
      site.entries = []; // { id, login, title, type, admin }
      site.spIds = new Map();
      let next = 3;
      const entry = (ref) => {
        const login = this.claimOf(ref);
        let e = site.entries.find((x) => x.login === login);
        if (!e) {
          e = { id: next++, login, title: ref, type: ref.startsWith('user:') ? 1 : 4, admin: false };
          site.entries.push(e);
        }
        return e;
      };
      (site.spGroups || []).forEach((g) => {
        site.spIds.set(g.title, next++);
        g.members.forEach(entry);
      });
      (site.admins || []).forEach((r) => (entry(r).admin = true));
      const visit = (web) => {
        (web.assignments || []).forEach((a) => a.to.startsWith('sp:') || entry(a.to));
        (web.lists || []).forEach((l) => (l.assignments || []).forEach((a) => a.to.startsWith('sp:') || entry(a.to)));
        (web.webs || []).forEach(visit);
      };
      visit(site.web);
      // People who merely visited: in the user list without any grant.
      (site.visitors || []).forEach(entry);
    }
  }

  /** Every claim a user holds (their own, the implicit ones, and group claims incl. nested and owned). */
  userClaims(upn, { withGroups = true } = {}) {
    const u = this.model.users.find((x) => x.upn.toLowerCase() === upn.toLowerCase());
    const claims = new Set(['i:0#.f|membership|' + upn.toLowerCase(), 'c:0(.s|true']);
    if (!u) return claims;
    if (u.userType === 'Guest' && this.everyoneExcludesGuests) claims.delete('c:0(.s|true');
    if (u.userType !== 'Guest') claims.add(this.claimOf('eeeu'));
    if (withGroups) {
      this.groupsOf(u.id).forEach((g) => claims.add(this.claimOf('group:' + g.id)));
      this.model.groups.filter((g) => g.unified && (g.owners || []).includes(u.id)).forEach((g) => claims.add(this.claimOf('owners:' + g.id)));
    }
    return claims;
  }

  /** Claims that stand for a group: its own, its owners' (M365), and the groups it is nested in. */
  groupPrincipalClaims(groupId) {
    const g = this.model.groups.find((x) => x.id === groupId);
    const claims = new Set([this.claimOf('group:' + groupId)]);
    if (g.unified) claims.add(this.claimOf('owners:' + groupId));
    this.groupsOf(groupId).forEach((p) => claims.add(this.claimOf('group:' + p.id)));
    return claims;
  }

  /** Groups a directory object is in, nested included. */
  groupsOf(objectId) {
    const out = new Map();
    const walk = (id) => {
      for (const g of this.model.groups) {
        if ((g.members || []).includes(id) && !out.has(g.id)) {
          out.set(g.id, g);
          walk(g.id);
        }
      }
    };
    walk(objectId);
    return [...out.values()];
  }

  /** Union of the role masks a set of claims gets from a list of assignments. */
  bitsFor(site, assignments, claims, { skipLimited = false } = {}) {
    let low = 0;
    let high = 0;
    for (const a of assignments) {
      let hit;
      if (a.to.startsWith('sp:')) {
        const g = (site.spGroups || []).find((x) => x.title === a.to.substring(3));
        hit = g && g.members.some((m) => claims.has(this.claimOf(m)));
      } else {
        hit = claims.has(this.claimOf(a.to));
      }
      if (hit) {
        for (const name of a.roles) {
          if (skipLimited && name === 'Limited Access') continue;
          low = (low | ROLES[name].low) >>> 0;
          high = (high | ROLES[name].high) >>> 0;
        }
      }
    }
    return { low, high };
  }

  isAdmin(site, claims) {
    return (site.admins || []).some((r) => claims.has(this.claimOf(r)));
  }

  /**
   * What the model says each principal can reach: [{ who, object, url, access }].
   * `who` is 'user:<upn>' or 'group:<id>'. Honours the same scope switches as the tool.
   */
  oracle(principals, { includeLibraries = true, includeLists = false, includeSubsites = true } = {}) {
    const out = [];
    for (const site of this.model.sites) {
      if (site.locked || site.personal) continue;
      for (const who of principals) {
        const isUser = who.startsWith('user:');
        const claims = isUser ? this.userClaims(who.substring(5)) : this.groupPrincipalClaims(who.substring(6));
        const admin = this.isAdmin(site, claims);
        const access = (assignments) => {
          if (admin) return 'Full control';
          const b = this.bitsFor(site, assignments, claims, { skipLimited: true });
          return levelOf(b.low, b.high);
        };
        const visit = (web, inheritedAssignments, isRoot) => {
          const assignments = web.assignments || inheritedAssignments;
          const a = access(assignments);
          if (a !== 'No access') out.push({ who, object: isRoot ? 'Site' : 'Subsite', url: this.origin + web.path, access: a });
          for (const l of web.lists || []) {
            if (l.hidden || l.system) continue;
            const isLib = (l.baseType ?? 1) === 1;
            if ((isLib && !includeLibraries) || (!isLib && !includeLists)) continue;
            const la = access(l.assignments || assignments);
            if (la !== 'No access') out.push({ who, object: isLib ? 'Library' : 'List', url: this.origin + encodeURI(l.path), access: la });
          }
          if (includeSubsites) (web.webs || []).forEach((w) => visit(w, assignments, false));
        };
        visit(site.web, [], true);
      }
    }
    return out;
  }

  // ---- HTTP ----------------------------------------------------------------------------------

  send(res, status, body, headers = {}) {
    const data = body === undefined ? '' : Buffer.isBuffer(body) ? body : JSON.stringify(body);
    res.writeHead(status, { 'Content-Type': Buffer.isBuffer(body) ? 'application/octet-stream' : 'application/json', ...headers });
    res.end(data);
  }

  spError(res, status, message, headers) {
    this.send(res, status, { 'odata.error': { code: '-1, Mock', message: { lang: 'en-US', value: message } } }, headers);
  }

  async handle(req, res) {
    const chunks = [];
    for await (const c of req) chunks.push(c);
    const body = Buffer.concat(chunks);
    const url = new URL(req.url, this.origin);
    const record = { method: req.method, path: decodeURIComponent(url.pathname), search: url.search, url: req.url, headers: req.headers, body };
    this.requests.push(record);

    for (let i = 0; i < this.faults.length; i++) {
      const f = this.faults[i](record);
      if (f) {
        if (f.once !== false) this.faults.splice(i, 1);
        if (f.delayMs) await new Promise((r) => setTimeout(r, f.delayMs));
        if (f.drop) {
          req.socket.destroy();
          return;
        }
        if (f.status) {
          this.send(res, f.status, f.body ?? { error: { message: 'injected fault' } }, f.headers || {});
          return;
        }
      }
    }

    if (url.pathname.startsWith('/login/')) return this.login(url, body, res);
    if (url.pathname === '/oidc') return this.oidc(req, url, res);

    const auth = String(req.headers.authorization || '');
    const token = auth.replace(/^Bearer /, '');
    if (!this.validTokens.has(token)) return this.send(res, 401, { error: { code: 'InvalidAuthenticationToken', message: 'Access token is missing, expired or revoked.' } });
    const aud = JSON.parse(Buffer.from(token.split('.')[1], 'base64url').toString()).aud;

    if (url.pathname.startsWith('/graph/')) {
      if (aud !== this.graphBase) return this.send(res, 401, { error: { message: 'Token is for another resource.' } });
      return this.graph(url, req, res);
    }
    if (aud !== this.origin) return this.send(res, 401, { error: { message: 'Token is for another resource.' } });
    return this.sharepoint(req, url, body, res);
  }

  // ---- Entra ID ----

  login(url, body, res) {
    const m = /^\/login\/([^/]+)\/oauth2\/v2\.0\/token$/.exec(url.pathname);
    const form = new URLSearchParams(body.toString());
    const bad = (code, description) => this.send(res, 400, { error: 'invalid_client', error_description: `${code}: ${description}\r\nTrace ID: x` });
    if (!m || decodeURIComponent(m[1]) !== this.opts.tenantId) return bad('AADSTS90002', 'Tenant not found.');
    if (form.get('grant_type') !== 'client_credentials') return bad('AADSTS70003', 'Unsupported grant type.');
    if (form.get('client_id') !== this.opts.clientId) return bad('AADSTS700016', 'Application was not found in the directory.');
    if (form.get('client_secret')) return bad('AADSTS7000215', 'Invalid client secret provided.');
    if (form.get('client_assertion_type') !== 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer') return bad('AADSTS50027', 'Assertion type missing.');
    const assertion = form.get('client_assertion') || '';
    if (assertion === this.oidcToken) {
      if (!this.opts.allowFederated) return bad('AADSTS70021', 'No matching federated identity record found for presented assertion.');
    } else {
      const err = this.verifyAssertion(assertion, this.origin + url.pathname);
      if (err) return bad(err[0], err[1]);
    }
    const scope = form.get('scope') || '';
    const resource = scope.replace(/\/\.default$/, '');
    let roles;
    if (resource === this.graphBase) roles = this.opts.graphRoles;
    else if (resource === this.origin) roles = this.opts.spRoles;
    else return bad('AADSTS500011', `The resource principal named ${resource} was not found in the tenant.`);
    const token = `${b64url('{"alg":"none"}')}.${b64url(JSON.stringify({ aud: resource, roles, n: ++this.tokensIssued }))}.sig`;
    this.validTokens.add(token);
    this.send(res, 200, { token_type: 'Bearer', expires_in: this.opts.expiresIn ?? 3599, access_token: token });
  }

  verifyAssertion(jwt, expectedAud) {
    const parts = jwt.split('.');
    if (parts.length !== 3) return ['AADSTS50027', 'JWT token is invalid or malformed.'];
    let header;
    let payload;
    try {
      header = JSON.parse(Buffer.from(parts[0], 'base64url').toString());
      payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString());
    } catch {
      return ['AADSTS50027', 'JWT token is invalid or malformed.'];
    }
    if (!this.opts.certPem) return ['AADSTS700027', 'No certificate is registered.'];
    const cert = new X509Certificate(this.opts.certPem);
    const x5t = createHash('sha1').update(cert.raw).digest('base64url');
    if (header.alg !== 'RS256' || header.x5t !== x5t) return ['AADSTS700027', 'The certificate with identifier used to sign the client assertion is not registered on application.'];
    const ok = createVerify('RSA-SHA256')
      .update(parts[0] + '.' + parts[1])
      .verify(createPublicKey(this.opts.certPem), Buffer.from(parts[2], 'base64url'));
    if (!ok) return ['AADSTS700027', 'Client assertion contains an invalid signature.'];
    const now = Date.now() / 1000;
    if (payload.aud !== expectedAud) return ['AADSTS700024', 'Client assertion audience is wrong: ' + payload.aud];
    if (payload.iss !== this.opts.clientId || payload.sub !== this.opts.clientId) return ['AADSTS700016', 'Issuer does not match the application.'];
    if (!(payload.nbf <= now && payload.exp > now) || payload.exp - now > 600) return ['AADSTS700024', 'Client assertion is not within its valid time range.'];
    if (!payload.jti) return ['AADSTS50027', 'jti missing'];
    return undefined;
  }

  oidc(req, url, res) {
    if (req.headers.authorization !== `Bearer ${this.oidcRequestToken}`) return this.send(res, 401, { message: 'bad runner token' });
    if (url.searchParams.get('audience') !== 'api://AzureADTokenExchange') return this.send(res, 400, { message: 'bad audience' });
    this.send(res, 200, { value: this.oidcToken });
  }

  // ---- Microsoft Graph ----

  graph(url, req, res) {
    const path = decodeURIComponent(url.pathname.substring('/graph'.length));
    const roles = this.opts.graphRoles;
    const deny = () => this.send(res, 403, { error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges to complete the operation.' } });
    const notFound = () => this.send(res, 404, { error: { code: 'Request_ResourceNotFound', message: 'Resource does not exist.' } });
    const page = (rows, sizeOverride) => {
      const size = sizeOverride ?? this.opts.pageSize;
      const skip = Number(url.searchParams.get('$skiptoken') || 0);
      const slice = rows.slice(skip, skip + size);
      const out = { '@odata.context': 'x', value: slice };
      if (skip + size < rows.length) {
        const next = new URL(url);
        next.searchParams.set('$skiptoken', String(skip + size));
        out['@odata.nextLink'] = next.toString();
      }
      this.send(res, 200, out);
    };
    const userJson = (u) => ({ id: u.id, displayName: u.displayName, mail: u.mail ?? u.upn, userPrincipalName: u.upn, userType: u.userType || 'Member', accountEnabled: u.enabled !== false });
    const groupJson = (g) => ({ '@odata.type': '#microsoft.graph.group', id: g.id, displayName: g.displayName, mail: g.mail ?? null, groupTypes: g.unified ? ['Unified'] : [], securityEnabled: !g.unified });

    if (path === '/v1.0/sites/getAllSites') {
      if (!roles.includes('Sites.Read.All')) return deny();
      return page(this.model.sites.filter((s) => !s.hiddenFromGraph).map((s) => ({ id: s.path, webUrl: this.origin + s.path, displayName: s.title, name: s.path.split('/').pop(), isPersonalSite: !!s.personal })));
    }
    let m;
    if ((m = /^\/v1\.0\/users\/([^/]+)$/.exec(path))) {
      if (!roles.includes('User.Read.All')) return deny();
      const key = m[1].toLowerCase();
      if (/\s/.test(key)) return this.send(res, 400, { error: { code: 'Request_BadRequest', message: 'Invalid object identifier.' } });
      const u = this.model.users.find((x) => x.id === key || x.upn.toLowerCase() === key);
      return u ? this.send(res, 200, userJson(u)) : notFound();
    }
    if (path === '/v1.0/users') {
      if (!roles.includes('User.Read.All')) return deny();
      if (req.headers.consistencylevel !== 'eventual') return this.send(res, 400, { error: { message: 'Advanced query needs ConsistencyLevel: eventual' } });
      const f = /^mail eq '((?:[^']|'')*)'/.exec(url.searchParams.get('$filter') || '');
      const mail = f ? f[1].replace(/''/g, "'").toLowerCase() : '';
      return this.send(res, 200, { value: this.model.users.filter((u) => (u.mail || '').toLowerCase() === mail || (u.otherMails || []).map((x) => x.toLowerCase()).includes(mail)).map(userJson) });
    }
    if ((m = /^\/v1\.0\/groups\/([^/]+)$/.exec(path))) {
      if (!roles.includes('GroupMember.Read.All')) return deny();
      const g = this.model.groups.find((x) => x.id === m[1]);
      return g ? this.send(res, 200, groupJson(g)) : notFound();
    }
    if (path === '/v1.0/groups') {
      if (!roles.includes('GroupMember.Read.All')) return deny();
      const f = /^displayName eq '((?:[^']|'')*)' or mail eq '((?:[^']|'')*)'$/.exec(url.searchParams.get('$filter') || '');
      if (!f) return this.send(res, 400, { error: { message: 'Unsupported filter' } });
      const name = f[1].replace(/''/g, "'").toLowerCase();
      return this.send(res, 200, { value: this.model.groups.filter((g) => g.displayName.toLowerCase() === name || (g.mail || '').toLowerCase() === name).map(groupJson) });
    }
    if ((m = /^\/v1\.0\/(users|groups)\/([^/]+)\/transitiveMemberOf$/.exec(path))) {
      if (!roles.includes('GroupMember.Read.All')) return deny();
      const rows = this.groupsOf(m[2]).map(groupJson);
      // Directory roles also come back from this call; the tool must ignore them.
      if (m[1] === 'users') rows.push({ '@odata.type': '#microsoft.graph.directoryRole', id: 'role-1', displayName: 'Reports Reader' });
      return page(rows, 2);
    }
    if ((m = /^\/v1\.0\/users\/([^/]+)\/ownedObjects$/.exec(path))) {
      if (this.opts.denyOwnedObjects) return deny();
      return page(this.model.groups.filter((g) => (g.owners || []).includes(m[1])).map(groupJson), 2);
    }
    this.send(res, 404, { error: { message: 'mock graph: no route for ' + path } });
  }

  // ---- SharePoint ----

  findWeb(path) {
    // Deepest web whose path is a prefix of `path`.
    let best;
    const lower = path.toLowerCase().replace(/\/+$/, '');
    for (const site of this.model.sites) {
      const visit = (web, inherited, parent) => {
        const assignments = web.assignments || inherited;
        const p = web.path.toLowerCase();
        if ((lower === p || lower.startsWith(p + '/')) && (!best || p.length > best.web.path.length)) best = { site, web, assignments, isRoot: !parent };
        (web.webs || []).forEach((w) => visit(w, assignments, web));
      };
      visit(site.web, [], undefined);
    }
    return best;
  }

  assignmentJson(site, a) {
    const id = a.to.startsWith('sp:') ? site.spIds.get(a.to.substring(3)) : site.entries.find((e) => e.login === this.claimOf(a.to)).id;
    return {
      PrincipalId: id,
      RoleDefinitionBindings: a.roles.map((name) => {
        const r = ROLES[name];
        const out = { Id: r.id, Name: name, Description: '', RoleTypeKind: r.kind, Hidden: false, Order: 1 };
        if (!this.opts.omitBasePermissions) out.BasePermissions = { High: String(r.high), Low: String(r.low) };
        return out;
      })
    };
  }

  effective(site, assignments, login) {
    const upn = login.replace(/^i:0#\.f\|membership\|/, '');
    const claims = this.userClaims(upn, { withGroups: !this.staleGroupUsers.has(upn.toLowerCase()) });
    if (this.isAdmin(site, claims)) return { High: '2147483647', Low: '4294967295' };
    const b = this.bitsFor(site, assignments, claims);
    return { High: String(b.high), Low: String(b.low) };
  }

  aliasValue(url, name) {
    const raw = url.searchParams.get('@' + name);
    const m = /^'((?:[^']|'')*)'$/.exec(raw || '');
    return m ? m[1].replace(/''/g, "'") : undefined;
  }

  sharepoint(req, url, body, res) {
    const full = decodeURIComponent(url.pathname);
    const at = full.toLowerCase().indexOf('/_api/');
    if (at < 0) return this.spError(res, 404, 'Not an API call.');
    const webPath = full.substring(0, at);
    const api = full.substring(at + 6);
    const found = this.findWeb(webPath || '/');
    if (!found || found.web.path.toLowerCase() !== webPath.toLowerCase()) return this.spError(res, 404, 'Site not found: ' + webPath);
    const { site, web, assignments, isRoot } = found;
    if (site.locked) return this.spError(res, 403, 'Access denied. The site is locked.');
    if (!this.opts.spRoles.includes('Sites.FullControl.All')) return this.spError(res, 403, 'Access denied. You do not have permission to perform this action or access this resource.');
    const value = (rows) => this.send(res, 200, { value: rows });
    const lower = api.toLowerCase();
    let m;

    if (lower === 'web') {
      return this.send(res, 200, { Title: web.title, Url: this.origin + web.path, HasUniqueRoleAssignments: isRoot || !!web.assignments });
    }
    if (lower === 'web/siteusers') {
      const filter = url.searchParams.get('$filter') || '';
      let rows = site.entries;
      if (/^PrincipalType eq 4$/.test(filter)) {
        if (this.rejectPrincipalTypeFilter) return this.spError(res, 400, "The query is not valid: field 'PrincipalType' cannot be used in the filter.");
        rows = rows.filter((e) => e.type === 4);
      } else if (filter) {
        const wanted = [];
        const re = /LoginName eq '((?:[^']|'')*)'( or |$)/y;
        let pos = 0;
        let mm;
        while (pos < filter.length) {
          re.lastIndex = pos;
          mm = re.exec(filter);
          if (!mm) return this.spError(res, 400, 'The query is not valid: ' + filter);
          wanted.push(mm[1].replace(/''/g, "'").toLowerCase());
          pos = re.lastIndex;
        }
        rows = rows.filter((e) => wanted.includes(e.login.toLowerCase()));
      }
      return value(rows.map((e) => ({ Id: e.id, LoginName: e.login, Title: e.title, IsSiteAdmin: e.admin, PrincipalType: e.type })));
    }
    if ((m = /^web\/siteusers\/getbyid\((\d+)\)\/groups$/.exec(lower))) {
      const e = site.entries.find((x) => x.id === Number(m[1]));
      if (!e) return this.spError(res, 404, 'User cannot be found.');
      return value((site.spGroups || []).filter((g) => g.members.some((r) => this.claimOf(r) === e.login)).map((g) => ({ Id: site.spIds.get(g.title), Title: g.title })));
    }
    if (lower === 'web/roleassignments') {
      return value(assignments.map((a) => this.assignmentJson(site, a)));
    }
    if (lower === 'web/getusereffectivepermissions(@v)') {
      const login = this.aliasValue(url, 'v');
      if (!login) return this.spError(res, 400, 'Missing @v');
      if (this.unknownUsers.has(login.replace(/^i:0#\.f\|membership\|/, ''))) return this.spError(res, 404, 'The specified user could not be found.');
      return this.send(res, 200, this.effective(site, assignments, login));
    }
    if (lower === 'web/lists') {
      const lists = (web.lists || []).filter((l) => !l.hidden || !/Hidden eq false/.test(url.searchParams.get('$filter') || ''));
      const size = this.opts.listPageSize;
      const skip = Number(url.searchParams.get('$skiptoken') || 0);
      const rows = lists.map((l) => ({ Id: l.id, Title: l.title, BaseType: l.baseType ?? 1, HasUniqueRoleAssignments: !!l.assignments, IsSystemList: !!l.system, RootFolder: { ServerRelativeUrl: l.path } }));
      if (size && skip + size < rows.length) {
        const next = new URL(url);
        next.searchParams.set('$skiptoken', String(skip + size));
        return this.send(res, 200, { value: rows.slice(skip, skip + size), 'odata.nextLink': next.toString() });
      }
      return value(size ? rows.slice(skip) : rows);
    }
    if ((m = /^web\/lists\(guid'([0-9a-f-]+)'\)\/(roleassignments|getusereffectivepermissions\(@v\))$/.exec(lower))) {
      const l = (web.lists || []).find((x) => x.id === m[1]);
      if (!l) return this.spError(res, 404, 'List does not exist.');
      const la = l.assignments || assignments;
      if (m[2] === 'roleassignments') return value(la.map((a) => this.assignmentJson(site, a)));
      if (this.unknownUsers.has(this.aliasValue(url, 'v').replace(/^i:0#\.f\|membership\|/, ''))) return this.spError(res, 404, 'The specified user could not be found.');
      return this.send(res, 200, this.effective(site, la, this.aliasValue(url, 'v')));
    }
    if (lower === 'web/webs') {
      return value((web.webs || []).map((w) => ({ Title: w.title, Url: this.origin + w.path })));
    }

    // ---- files ----
    const p = this.aliasValue(url, 'p');
    const key = (x) => x.toLowerCase();
    if (lower === 'web/getfolderbyserverrelativepath(decodedurl=@p)') {
      return this.folders.has(key(p)) ? this.send(res, 200, { Exists: true }) : this.spError(res, 404, 'File Not Found.');
    }
    if (lower === 'web/folders/addusingpath(decodedurl=@p)' && req.method === 'POST') {
      const parent = p.substring(0, p.lastIndexOf('/'));
      if (!this.folders.has(key(parent))) return this.spError(res, 404, 'Parent folder not found.');
      this.folders.add(key(p));
      return this.send(res, 200, { Exists: true });
    }
    if (lower === 'web/getfolderbyserverrelativepath(decodedurl=@p)/files') {
      if (!this.folders.has(key(p))) return this.spError(res, 404, 'File Not Found.');
      const rows = [...this.files.keys()].filter((f) => key(f.substring(0, f.lastIndexOf('/'))) === key(p)).map((f) => ({ Name: f.substring(f.lastIndexOf('/') + 1), Length: String(this.files.get(f).length), TimeLastModified: '2026-01-01T00:00:00Z' }));
      return value(rows);
    }
    if (lower === 'web/getfolderbyserverrelativepath(decodedurl=@f)/files/addusingpath(decodedurl=@n,overwrite=true)' && req.method === 'POST') {
      const f = this.aliasValue(url, 'f');
      const n = this.aliasValue(url, 'n');
      if (!this.folders.has(key(f))) return this.spError(res, 404, 'File Not Found.');
      const path = this.canonical(f + '/' + n);
      if (this.lockedFiles.has(key(path))) return this.spError(res, 423, 'The file is locked for shared use.');
      this.files.set(path, body);
      return this.send(res, 200, { Name: n });
    }
    if (lower === 'web/getfilebyserverrelativepath(decodedurl=@p)/$value') {
      const path = this.canonical(p);
      return this.files.has(path) ? this.send(res, 200, this.files.get(path)) : this.spError(res, 404, 'File Not Found.');
    }
    if (lower === 'web/getfilebyserverrelativepath(decodedurl=@p)' && req.method === 'POST' && req.headers['x-http-method'] === 'DELETE') {
      const path = this.canonical(p);
      if (!this.files.has(path)) return this.spError(res, 404, 'File Not Found.');
      this.files.delete(path);
      return this.send(res, 200, undefined);
    }
    if ((m = /^web\/getfilebyserverrelativepath\(decodedurl=@p\)\/(startupload|continueupload|finishupload|cancelupload)\(uploadid=guid'([0-9a-f-]+)'(?:,fileoffset=(\d+))?\)$/.exec(lower)) && req.method === 'POST') {
      const path = this.canonical(p);
      const [, op, id, offset] = m;
      if (op === 'startupload') {
        if (!this.files.has(path)) return this.spError(res, 404, 'File Not Found.');
        this.uploads.set(id, { path, data: Buffer.from(body) });
        return this.send(res, 200, { value: String(body.length) });
      }
      const u = this.uploads.get(id);
      if (!u || u.path !== path) return this.spError(res, 400, 'Upload session not found.');
      if (op === 'cancelupload') {
        this.uploads.delete(id);
        return this.send(res, 200, undefined);
      }
      if (Number(offset) !== u.data.length) return this.spError(res, 400, `The file offset ${offset} does not match the expected ${u.data.length}.`);
      u.data = Buffer.concat([u.data, body]);
      if (op === 'finishupload') {
        this.files.set(path, u.data);
        this.uploads.delete(id);
      }
      return this.send(res, 200, { value: String(u.data.length) });
    }
    this.spError(res, 404, 'mock sharepoint: no route for ' + req.method + ' ' + api);
  }

  /** Paths are case-insensitive in SharePoint; keep the first spelling seen. */
  canonical(path) {
    for (const k of this.files.keys()) if (k.toLowerCase() === path.toLowerCase()) return k;
    return path;
  }

  /** Makes a document library (and optional folders) exist. */
  addLibrary(path) {
    this.folders.add(path.toLowerCase());
  }

  file(path) {
    return this.files.get(this.canonical(path));
  }

  count(pred) {
    return this.requests.filter(pred).length;
  }
}
