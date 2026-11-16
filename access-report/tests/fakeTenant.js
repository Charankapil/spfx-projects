// A fake SharePoint tenant that answers the REST calls of the web part (through the
// SPHttpClient stand-in) and of the flow (through the simulated "Send an HTTP request to
// SharePoint" action) from the same data, so the two can be compared.

const ORIGIN = 'https://contoso.sharepoint.com';
const FULL = { Low: '4294967295', High: '2147483647' };
const READ = { Low: '138612833', High: '176' };
const EDIT = { Low: '1011030767', High: '432' };
const DESIGN = { Low: '1012866047', High: '432' };
// Limited Access only: can open the site shell, cannot view items.
const LIMITED = { Low: '134287360', High: '48' };
const NONE = { Low: '0', High: '0' };

const trimSlash = (u) => u.replace(/\/+$/, '');

class FakeTenant {
  constructor({ home = ORIGIN + '/sites/it', sites = [], pageSize = 500 } = {}) {
    this.origin = ORIGIN;
    this.home = home;
    this.sites = sites;
    this.pageSize = pageSize;
    this.homePath = trimSlash(home).replace(ORIGIN, '');
    this.lists = {};
    this.files = {};
    this.folders = new Set();
    this.log = [];
    this.nextId = 1;
    /** Search behaviour: 'ok', or 'reject' (400 on every query). A query with IndexDocId is always refused, as some tenants do. */
    this.search = 'ok';
    this.picker = [];
  }

  /** The lists and library the web part's set-up creates, already in place. */
  provisioned({ subjects = [], includeLibraries = true, siteFilter = '', folder } = {}) {
    const p = this.homePath;
    this.lists[`${p}/Lists/AccessReportSubjects`] = { Title: 'Access Report Subjects', fields: ['Title', 'ARLogin', 'ARActive'], items: subjects.map((s, i) => ({ Id: i + 1, Title: s.label, ARLogin: s.login, ARActive: s.active !== false })) };
    this.lists[`${p}/Lists/AccessReportSettings`] = {
      Title: 'Access Report Settings',
      fields: ['Title', 'ARIncludeLibraries', 'ARSiteFilter', 'ARReportFolder', 'ARLastRun'],
      items: [{ Id: 1, Title: 'Settings', ARIncludeLibraries: includeLibraries, ARSiteFilter: siteFilter, ARReportFolder: folder === undefined ? `${p}/AccessReports` : folder, ARLastRun: null }]
    };
    this.lists[`${p}/AccessReports`] = { Title: 'Access Reports', fields: [], items: [] };
    this.folders.add(`${p}/AccessReports`.toLowerCase());
    return this;
  }

  settingsItem() {
    return this.lists[`${this.homePath}/Lists/AccessReportSettings`].items[0];
  }

  site(url) {
    return this.sites.find((s) => trimSlash(s.url).toLowerCase() === trimSlash(url).toLowerCase());
  }

  // ---- the REST surface -------------------------------------------------------------

  /** -> { status, body } where body is JSON-able, or a string for file content. */
  handle(method, url, body, headers = {}) {
    this.log.push({ method, url });
    // What a browser or the connector would send: nothing after a raw #.
    url = url.split('#')[0];
    const m = /^(https:\/\/[^?#]*?)\/_api\/([^?#]*)(?:\?(.*))?$/i.exec(url);
    if (!m) return this.error(400, `Not a REST url: ${url}`);
    const web = m[1];
    let path;
    const params = {};
    try {
      path = decodeURIComponent(m[2]);
      (m[3] || '').split('&').forEach((kv) => {
        if (!kv) return;
        const eq = kv.indexOf('=');
        params[decodeURIComponent(eq < 0 ? kv : kv.slice(0, eq))] = eq < 0 ? '' : decodeURIComponent(kv.slice(eq + 1));
      });
    } catch (e) {
      return this.error(400, `Malformed encoding in ${url}`);
    }
    const alias = (name) => {
      const v = params[name];
      if (v === undefined) return undefined;
      // A quoted literal: any quote inside it must be doubled.
      const q = /^'((?:[^']|'')*)'$/.exec(v);
      return q ? q[1].replace(/''/g, "'") : undefined;
    };
    const verb = headers['X-HTTP-Method'] || method;

    if (/^search\/query$/i.test(path)) return this.searchQuery(web, params);
    if (/^SP\.UI\.ApplicationPages\.ClientPeoplePickerWebServiceInterface\.clientPeoplePickerSearchUser$/i.test(path)) {
      return { status: 200, body: { value: JSON.stringify(this.picker) } };
    }

    const isHome = trimSlash(web).toLowerCase() === trimSlash(this.home).toLowerCase();
    let r;

    // --- configuration lists and report files (home site) ---
    if ((r = /^web\/GetList\(@u\)(.*)$/i.exec(path))) {
      if (!isHome) return this.error(404, 'List does not exist here.');
      const listPath = alias('@u');
      const list = listPath === undefined ? undefined : this.lists[listPath];
      if (!list) return this.error(404, 'File Not Found.');
      const tail = r[1];
      if (tail === '') {
        if (verb === 'MERGE') {
          Object.assign(list, JSON.parse(body));
          return { status: 204, body: '' };
        }
        return { status: 200, body: { Id: 'list-id', Title: list.Title } };
      }
      if (/^\/fields$/i.test(tail)) return { status: 200, body: { value: list.fields.map((f) => ({ InternalName: f })) } };
      if (/^\/fields\/createfieldasxml$/i.test(tail)) {
        const xml = JSON.parse(body).parameters.SchemaXml;
        list.fields.push(/ Name="([^"]+)"/.exec(xml)[1]);
        return { status: 200, body: {} };
      }
      if (/^\/items$/i.test(tail)) {
        if (method === 'POST') {
          const item = Object.assign({ Id: (list.items.reduce((n, i) => Math.max(n, i.Id), 0) || 0) + 1 }, JSON.parse(body));
          Object.keys(item).forEach((k) => {
            if (k !== 'Id' && list.fields.indexOf(k) < 0) throw new Error(`Unknown column ${k}`);
          });
          list.items.push(item);
          return { status: 201, body: item };
        }
        let items = list.items.slice().sort((a, b) => a.Id - b.Id);
        if (params.$top) items = items.slice(0, Number(params.$top));
        return { status: 200, body: { value: items.map((i) => Object.assign({}, i)) } };
      }
      if ((r = /^\/items\((\d+)\)$/i.exec(tail))) {
        const item = list.items.find((i) => i.Id === Number(r[1]));
        if (!item) return this.error(404, 'Item does not exist.');
        if (verb === 'DELETE') {
          list.items = list.items.filter((i) => i !== item);
          return { status: 200, body: '' };
        }
        if (verb === 'MERGE') {
          if (headers['IF-MATCH'] !== '*') return this.error(412, 'Missing IF-MATCH');
          const patch = JSON.parse(body);
          Object.keys(patch).forEach((k) => {
            if (list.fields.indexOf(k) < 0) throw new Error(`Unknown column ${k}`);
          });
          Object.assign(item, patch);
          return { status: 204, body: '' };
        }
        return { status: 200, body: Object.assign({}, item) };
      }
      return this.error(400, `Unsupported list call ${tail}`);
    }
    if (/^web\/lists$/i.test(path) && method === 'POST') {
      const def = JSON.parse(body);
      const p = `${this.homePath}/${def.BaseTemplate === 101 ? '' : 'Lists/'}${def.Title}`;
      if (this.lists[p]) return this.error(500, 'A list with that name already exists.');
      this.lists[p] = { Title: def.Title, fields: ['Title'], items: [] };
      if (def.BaseTemplate === 101) this.folders.add(p.toLowerCase());
      return { status: 201, body: { Title: def.Title } };
    }
    if ((r = /^web\/GetFolderByServerRelativeUrl\(@f\)(.*)$/i.exec(path))) {
      const folder = alias('@f');
      const exists = folder !== undefined && this.folders.has(trimSlash(folder).toLowerCase());
      if (r[1] === '') return { status: 200, body: { Exists: exists } };
      if (!exists) return this.error(404, 'File Not Found.');
      if (/^\/Files\/add\(url=@n,overwrite=true\)$/i.test(r[1])) {
        const name = alias('@n');
        if (!name || method !== 'POST' || typeof body !== 'string') return this.error(400, 'Bad upload');
        this.files[`${trimSlash(folder)}/${name}`] = body;
        return { status: 200, body: { Name: name, ServerRelativeUrl: `${trimSlash(folder)}/${name}` } };
      }
      if (/^\/Files$/i.test(r[1])) {
        const prefix = trimSlash(folder).toLowerCase() + '/';
        return {
          status: 200,
          body: {
            value: Object.keys(this.files)
              .filter((f) => f.toLowerCase().indexOf(prefix) === 0 && f.slice(prefix.length).indexOf('/') < 0)
              .map((f) => ({ Name: f.slice(prefix.length), ServerRelativeUrl: f, Length: String(this.files[f].length), TimeLastModified: '2026-11-01T03:10:00Z', ModifiedBy: { Title: 'Flow Owner' } }))
          }
        };
      }
      return this.error(400, `Unsupported folder call ${r[1]}`);
    }
    if (/^web\/GetFileByServerRelativeUrl\(@f\)\/\$value$/i.test(path)) {
      const f = this.files[alias('@f')];
      return f === undefined ? this.error(404, 'File Not Found.') : { status: 200, body: f, text: true };
    }

    // --- the sites being scanned ---
    const site = this.site(web);
    if (!site) return this.error(404, 'Site not found.');
    if (site.deny) return this.error(site.deny, site.deny === 403 ? 'Access denied. You do not have permission to perform this action or access this resource.' : 'Site unavailable.');
    if (site.broken) return site.broken;
    const maskOf = (users) => {
      const login = alias('@v');
      if (login === undefined) return this.error(400, 'The expression is not valid.');
      if ((this.unknownUsers || []).indexOf(login.toLowerCase()) >= 0) return this.error(404, 'User cannot be found.');
      const key = Object.keys(users || {}).find((k) => k.toLowerCase() === login.toLowerCase());
      return { status: 200, body: Object.assign({}, key ? users[key] : NONE) };
    };
    const listById = (id) => (site.lists || []).find((l) => l.Id.toLowerCase() === id.toLowerCase());

    if (/^web\/getusereffectivepermissions\(@v\)$/i.test(path)) return maskOf(site.users);
    if (/^web\/lists$/i.test(path)) {
      if (site.listsFail) return this.error(site.listsFail, 'Lists could not be read.');
      if (!/Hidden eq false/.test(params.$filter || '')) return this.error(400, 'Expected the hidden-lists filter.');
      return { status: 200, body: { value: (site.lists || []).filter((l) => !l.Hidden).map(({ users, assignments, fail, Hidden, ...rest }) => rest) } };
    }
    if ((r = /^web\/lists\(guid'([^']+)'\)\/getusereffectivepermissions\(@v\)$/i.exec(path))) {
      const l = listById(r[1]);
      if (!l) return this.error(404, 'List does not exist.');
      if (l.fail) return this.error(l.fail, 'The library could not be read.');
      return maskOf(l.users);
    }
    if (/^web\/siteusers$/i.test(path)) {
      const wanted = [];
      const re = /LoginName eq '((?:[^']|'')*)'/g;
      let q;
      while ((q = re.exec(params.$filter || ''))) wanted.push(q[1].replace(/''/g, "'").toLowerCase());
      if (!wanted.length || !/Groups/.test(params.$expand || '')) return this.error(400, 'Bad site user query.');
      return { status: 200, body: { value: (site.siteUsers || []).filter((u) => wanted.indexOf(u.LoginName.toLowerCase()) >= 0).map((u) => Object.assign({ IsSiteAdmin: false, Groups: [] }, u)) } };
    }
    if (/^web\/roleassignments$/i.test(path)) {
      if (site.assignmentsFail) return this.error(site.assignmentsFail, 'Permissions could not be read.');
      return { status: 200, body: { value: site.assignments || [] } };
    }
    if ((r = /^web\/lists\(guid'([^']+)'\)\/roleassignments$/i.exec(path))) {
      const l = listById(r[1]);
      if (!l) return this.error(404, 'List does not exist.');
      if (l.fail) return this.error(l.fail, 'The library could not be read.');
      return { status: 200, body: { value: l.assignments || [] } };
    }
    return this.error(400, `Unsupported call ${path}`);
  }

  error(status, message) {
    return { status, body: { 'odata.error': { code: '-1, System.Exception', message: { lang: 'en-US', value: message } } }, message };
  }

  searchQuery(web, params) {
    const refuse = () => this.error(400, "We didn't understand your search terms. Make sure they're using proper syntax.");
    const text = (/^'(.*)'$/s.exec(params.querytext || '') || [])[1] || '';
    if (text !== 'contentclass:STS_Site' || this.search === 'reject') return refuse();
    // Like the tenant that reported the bug: any DocId sort is refused; others can be refused on request.
    const sort = (/^'(.*)'$/s.exec(params.sortlist || "''") || [])[1];
    if (sort === undefined || /docid/i.test(sort) || (this.refusedSorts || []).indexOf(sort) >= 0) return refuse();
    const known = { '': null, 'Created:ascending': 'created', 'LastModifiedTime:ascending': 'modified' };
    if (!(sort in known)) return this.error(500, `Unexpected sortlist ${sort}`);
    this.sortsUsed = (this.sortsUsed || []).concat(params.rowlimit === '1' ? [] : [sort]);
    let hits = this.sites
      .concat(this.extraSearchHits || [])
      .map((s, i) => Object.assign({ docId: 1000 + i * 7, created: 5000 - i * 3, modified: (i * 37) % 101 }, s));
    // Unsorted: a "relevance" order that is not the site order, so a mismatch would show.
    hits.sort(known[sort] ? (a, b) => a[known[sort]] - b[known[sort]] || a.docId - b.docId : (a, b) => (a.docId % 13) - (b.docId % 13) || a.docId - b.docId);
    const start = Number(params.startrow || 0);
    hits = hits.slice(start, start + (params.rowlimit === '1' ? 1 : this.pageSize));
    const rows = hits.map((s) => ({
      Cells: [
        { Key: 'Rank', Value: '0', ValueType: 'Edm.Double' },
        { Key: 'DocId', Value: String(s.docId), ValueType: 'Edm.Int64' },
        { Key: 'Title', Value: s.title === undefined ? null : s.title, ValueType: 'Edm.String' },
        { Key: 'Path', Value: s.searchPath || s.url, ValueType: 'Edm.String' },
        { Key: 'WebTemplate', Value: s.template || 'GROUP', ValueType: 'Edm.String' },
        { Key: 'PartitionId', Value: null, ValueType: 'Null' }
      ]
    }));
    return { status: 200, body: { ElapsedTime: 12, PrimaryQueryResult: { RelevantResults: { TotalRows: rows.length, Table: { Rows: rows } } } } };
  }

  // ---- the two clients ----------------------------------------------------------------

  /** For tests/mocks/sp-http.js: (method, url, config, opts) -> Response. */
  spHttpHandler() {
    return async (method, url, config, opts) => {
      const headers = Object.assign({}, opts.headers || {});
      const r = this.handle(method, url, opts.body, headers);
      if (r.transport) throw new TypeError('Failed to fetch');
      const text = r.text || typeof r.body === 'string' ? r.body : JSON.stringify(r.body);
      return new Response(r.status === 204 ? null : text, { status: r.status, headers: { 'content-type': 'application/json' } });
    };
  }

  /** For tests/wdl.js: answers the SharePoint connector's HttpRequest operation. */
  connector() {
    return async (inputs, actionName) => {
      const host = inputs.host || {};
      if (host.apiId !== '/providers/Microsoft.PowerApps/apis/shared_sharepointonline' || host.operationId !== 'HttpRequest' || host.connectionName !== 'shared_sharepointonline') {
        throw new Error(`${actionName}: not the SharePoint HttpRequest operation`);
      }
      const p = inputs.parameters;
      if (typeof p.dataset !== 'string' || !/^https:\/\//.test(p.dataset)) throw new Error(`${actionName}: dataset is not a site address: ${p.dataset}`);
      if (typeof p['parameters/uri'] !== 'string' || !/^_api\//.test(p['parameters/uri'])) throw new Error(`${actionName}: uri must be relative and start with _api/`);
      if (/@\{|@[a-z]+\(/i.test(p['parameters/uri'].replace(/\(@[a-z]\)|[?&,=]@[a-z]\b/gi, ''))) throw new Error(`${actionName}: unevaluated expression left in ${p['parameters/uri']}`);
      const headers = p['parameters/headers'] || {};
      if (headers.Accept !== 'application/json;odata=nometadata') throw new Error(`${actionName}: expected the nometadata Accept header`);
      const url = `${trimSlash(p.dataset)}/${p['parameters/uri']}`;
      const r = this.handle(p['parameters/method'], url, p['parameters/body'], headers);
      if (r.transport) return { status: 'TimedOut', outputs: null };
      if (r.status >= 200 && r.status < 300) return { status: 'Succeeded', outputs: { statusCode: r.status, headers: {}, body: r.body === '' ? undefined : r.body } };
      // The connector's error envelope - or whatever the site sent, when that is not JSON.
      const body = r.raw !== undefined ? r.raw : { status: r.status, message: `${r.message || 'Error'}\r\nclientRequestId: 11111111-2222\r\nserviceRequestId: 33333333-4444`, source: url, errors: ['-1', 'System.Exception'] };
      return { status: 'Failed', outputs: { statusCode: r.status, headers: {}, body } };
    };
  }
}

module.exports = { FakeTenant, ORIGIN, FULL, READ, EDIT, DESIGN, LIMITED, NONE };
