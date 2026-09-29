/**
 * A fake SharePoint Online REST surface for tests: webs, subwebs, document
 * libraries (small ones in memory and a lazily generated large one), paging
 * with odata.nextLink / $skiptoken, $filter on Id, throttling, failures,
 * StorageMetrics and a Site Assets library for saved results.
 */
const MB = 1024 * 1024;
const GB = 1024 * MB;
export const ORIGIN = 'https://contoso.sharepoint.com';
export const SITE = '/sites/finance';

function hash(n) {
  n = (n ^ 61) ^ (n >>> 16);
  n = n + (n << 3);
  n = n ^ (n >>> 4);
  n = Math.imul(n, 0x27d4eb2d);
  n = n ^ (n >>> 15);
  return (n >>> 0) / 4294967296;
}

export function monthsBefore(now, months, extraDays) {
  const d = new Date(now.getTime());
  d.setMonth(d.getMonth() - months);
  d.setDate(d.getDate() - extraDays);
  return d;
}

function makeItems(count, seed, opts, now) {
  const items = [];
  for (let id = 1; id <= count; id++) {
    const r = hash(id * 31 + seed);
    if (id % 25 === 0) {
      // Folders carry their own (creation) date, like real ones.
      const folderAge = opts.age(r, hash(id * 19 + seed));
      items.push({ id, folder: true, name: `Folder ${id}`, modified: monthsBefore(now, folderAge, 0).toISOString() });
      continue;
    }
    const age = opts.age(r, hash(id * 17 + seed));
    const size = Math.floor(opts.size(hash(id * 13 + seed)));
    const ext = opts.ext[id % opts.ext.length];
    const name = ext ? `${opts.prefix}-${id}.${ext}` : `${opts.prefix}-${id}`;
    items.push({ id, folder: false, modified: monthsBefore(now, age, Math.floor(hash(id * 7 + seed) * 20)).toISOString(), size, name });
  }
  return items;
}

/**
 * options:
 *   bigItems        IDs in the large lazy library (0 = no large library)
 *   flaky           a library whose item reads fail until heal() is called
 *   poisonItem      a library with one item that fails every page containing it
 *   throttleRate    fraction of large-library requests answered 429 Retry-After 0
 *   alwaysThrottle  a library that always answers 429 with Retry-After 120
 *   extraLibraries  number of extra small libraries (a mix of active and dormant)
 *   notAcceptable   'orderby': item queries with $orderby answer 406 (empty body);
 *                   'both': also the lean Accept header, so only default headers work
 *   latencyMs       simulated time per request
 *   savedFiles      initial Site Assets files { path: text }
 *   owner           whether the current user manages the web
 */
export function createMockSharePoint(options = {}) {
  const now = new Date();
  const opts = {
    bigItems: 0,
    throttleRate: 0,
    alwaysThrottle: false,
    flaky: false,
    poisonItem: false,
    latencyMs: 3,
    notAcceptable: '',
    savedFiles: {},
    owner: true,
    ...options
  };
  let rng = 42;
  const rand = () => ((rng = (rng * 1103515245 + 12345) % 2147483648) / 2147483648);

  const oldHeavy = (r, r2) => (r < 0.2 ? Math.floor(r2 * 12) : r < 0.35 ? 12 + Math.floor(r2 * 12) : 24 + Math.floor(r2 * 110));
  const recent = (r, r2) => (r < 0.75 ? Math.floor(r2 * 12) : 12 + Math.floor(r2 * 30));
  const dormant = (r, r2) => 40 + Math.floor(r2 * 60);

  const BIG = opts.bigItems;
  const bigItem = (id) => {
    if (id < 1 || id > BIG || id % 17 === 0) return null; // deleted IDs leave gaps
    if (id % 25 === 0) return { id, folder: true, name: `Folder ${id}`, modified: monthsBefore(now, 30, 0).toISOString() };
    const age = hash(id) < 0.3 ? Math.floor(hash(id * 3) * 12) : 12 + Math.floor(hash(id * 5) * 100);
    return {
      id,
      folder: false,
      modified: monthsBefore(now, age, Math.floor(hash(id * 7) * 20)).toISOString(),
      size: Math.floor(hash(id * 11) * 4 * MB),
      name: `Scan-${id}.pdf`,
      hidden: id % 401 === 0 // item-level permissions hide it from this user
    };
  };

  const libs = {
    docs: { id: 'aaaaaaaa-0000-0000-0000-000000000001', title: 'Documents', url: `${SITE}/Shared Documents`, items: makeItems(12400, 1, { age: oldHeavy, size: (r) => (r < 0.02 ? 200 * MB + r * 50 * 2 * GB : r * 30 * MB), prefix: 'Report', ext: ['docx', 'xlsx', 'pdf', 'pptx'] }, now) },
    assets: { id: 'aaaaaaaa-0000-0000-0000-000000000002', title: 'Site Assets', url: `${SITE}/SiteAssets`, items: makeItems(40, 2, { age: recent, size: (r) => r * 2 * MB, prefix: 'asset', ext: ['png', 'json'] }, now), noMetrics: true },
    audit: { id: 'aaaaaaaa-0000-0000-0000-000000000003', title: 'Audit Archive 2016', url: `${SITE}/Audit Archive 2016`, items: makeItems(6200, 3, { age: dormant, size: (r) => r * 60 * MB, prefix: 'Audit #1 100%', ext: ['pdf', 'zip'] }, now), failBigPages: true },
    empty: { id: 'aaaaaaaa-0000-0000-0000-000000000004', title: 'Empty Library', url: `${SITE}/Empty Library`, items: [] },
    pages: { id: 'aaaaaaaa-0000-0000-0000-000000000005', title: 'Site Pages', url: `${SITE}/SitePages`, baseTemplate: 119, items: makeItems(30, 5, { age: recent, size: (r) => r * MB, prefix: 'Page', ext: ['aspx'] }, now) },
    hold: { id: 'aaaaaaaa-0000-0000-0000-000000000006', title: 'Preservation Hold Library', url: `${SITE}/PreservationHoldLibrary`, hidden: true, items: makeItems(300, 6, { age: oldHeavy, size: (r) => r * 20 * MB, prefix: 'Held', ext: ['docx', ''] }, now) },
    catalog: { id: 'aaaaaaaa-0000-0000-0000-000000000007', title: 'Master Page Gallery', url: `${SITE}/_catalogs/masterpage`, hidden: true, isCatalog: true, items: makeItems(20, 7, { age: oldHeavy, size: (r) => r * MB, prefix: 'm', ext: ['master'] }, now) },
    projects: { id: 'bbbbbbbb-0000-0000-0000-000000000001', title: 'Documents', url: `${SITE}/projects/Shared Documents`, throttleOnce: true, items: makeItems(8300, 8, { age: recent, size: (r) => r * 25 * MB, prefix: 'Plan', ext: ['docx', 'mp4', 'xlsx'] }, now) },
    big: BIG ? { id: 'cccccccc-0000-0000-0000-000000000001', title: 'Scanned Records', url: `${SITE}/projects/Scanned Records`, big: true } : null,
    flaky: opts.flaky ? { id: 'eeeeeeee-0000-0000-0000-000000000001', title: 'Flaky Library', url: `${SITE}/Flaky Library`, failing: true, items: makeItems(900, 10, { age: recent, size: (r) => r * 5 * MB, prefix: 'f', ext: ['docx'] }, now) } : null,
    poison: opts.poisonItem ? { id: 'ffffffff-0000-0000-0000-000000000001', title: 'Poisoned Library', url: `${SITE}/Poisoned Library`, poisonId: 1200, items: makeItems(3000, 11, { age: oldHeavy, size: (r) => r * 5 * MB, prefix: 'p', ext: ['pdf'] }, now) } : null,
    stuck: opts.alwaysThrottle ? { id: 'dddddddd-0000-0000-0000-000000000001', title: 'Busy Library', url: `${SITE}/Busy Library`, alwaysThrottle: true, items: makeItems(100, 9, { age: recent, size: () => MB, prefix: 'b', ext: ['txt'] }, now) } : null
  };
  if (libs.big) {
    let count = 0;
    for (let i = 1; i <= BIG; i++) if (bigItem(i)) count++;
    libs.big.itemCount = count;
  }

  const extra = [];
  for (let i = 0; i < (opts.extraLibraries || 0); i++) {
    extra.push({
      id: `99999999-0000-0000-0000-${String(i).padStart(12, '0')}`,
      title: `Project ${i}`,
      url: `${SITE}/projects/Project ${i}`,
      items: makeItems(400 + (i % 7) * 300, 100 + i, { age: i % 3 === 0 ? recent : dormant, size: (r) => r * 8 * MB, prefix: `P${i}`, ext: ['docx', 'pdf', 'xlsx'] }, now)
    });
  }

  const webs = {
    [SITE]: { title: 'Finance', subwebs: [`${SITE}/projects`, `${SITE}/hr`], libraries: [libs.docs, libs.assets, libs.audit, libs.empty, libs.pages, libs.hold, libs.catalog, libs.flaky, libs.poison, libs.stuck].filter(Boolean) },
    [`${SITE}/projects`]: { title: 'Projects', subwebs: [], libraries: [libs.projects, libs.big, ...extra].filter(Boolean) },
    [`${SITE}/hr`]: { title: 'HR (restricted)', subwebs: [], libraries: null }
  };

  const itemCount = (l) => (l.big ? l.itemCount : l.items.length);
  const visibleItems = (l, gt, le, top) => {
    const out = [];
    if (l.big) {
      for (let id = gt + 1; id <= Math.min(le, BIG) && out.length < top; id++) {
        const it = bigItem(id);
        if (it && !it.hidden) out.push(it);
      }
      return out;
    }
    for (const it of l.items) {
      if (it.id > gt && it.id <= le) {
        out.push(it);
        if (out.length >= top) break;
      }
    }
    return out;
  };

  /** Newest Modified of any item (files and folders), hidden ones included: what SharePoint reports for the list. */
  function lastChange(l) {
    if (l.big) return now.toISOString();
    let max = 0;
    for (const i of l.items) max = Math.max(max, new Date(i.modified).getTime());
    return new Date(max || now.getTime() - 400 * 24 * 3600 * 1000).toISOString();
  }

  const files = { ...opts.savedFiles };
  let inFlight = 0;
  let maxInFlight = 0;
  const log = [];
  const acceptHeaders = new Set();

  function resp(status, body, headers) {
    const text = typeof body === 'string' ? body : JSON.stringify(body);
    return Promise.resolve({
      ok: status >= 200 && status < 300,
      status,
      statusText: String(status),
      headers: { get: (h) => (headers || {})[h] || null },
      json: () => Promise.resolve(JSON.parse(text)),
      text: () => Promise.resolve(text)
    });
  }

  let notAcceptableCount = 0;

  function handle(url, method, options) {
    log.push(`${method} ${decodeURIComponent(url)}`);
    const u = new URL(url);
    const apiIdx = u.pathname.indexOf('/_api/');
    const webPath = decodeURIComponent(u.pathname.substring(0, apiIdx));
    const api = decodeURIComponent(u.pathname.substring(apiIdx + 6));
    const q = new URLSearchParams(u.search);
    const web = webs[webPath];
    if (!web) return resp(404, `no web ${webPath}`);

    if (api === 'site') return resp(200, { Usage: { Storage: String(Math.round(3.4 * 1024 * GB)) } });
    if (api === 'web') return resp(200, { Title: web.title });
    if (api.startsWith('web/getsubwebsfilteredforcurrentuser')) {
      return resp(200, { value: web.subwebs.map((s) => ({ Title: webs[s].title, ServerRelativeUrl: s, WebTemplate: 'STS' })) });
    }
    if (api === 'web/lists' && method === 'GET') {
      if ((q.get('$filter') || '').indexOf('BaseTemplate eq 101') >= 0) {
        return resp(200, { value: [{ RootFolder: { ServerRelativeUrl: `${SITE}/SiteAssets` } }] });
      }
      if (!web.libraries) return resp(403, { 'odata.error': { code: '-2147024891', message: { lang: 'en-US', value: 'Access denied.' } } });
      return resp(200, {
        value: web.libraries.map((l) => ({
          Id: l.id,
          Title: l.title,
          ItemCount: itemCount(l),
          Hidden: !!l.hidden,
          IsCatalog: !!l.isCatalog,
          BaseTemplate: l.baseTemplate || 101,
          LastItemUserModifiedDate: lastChange(l),
          RootFolder: { ServerRelativeUrl: l.url }
        }))
      });
    }
    let m = api.match(/^web\/lists\(guid'([^']+)'\)\/items$/);
    if (m) {
      const lib = web.libraries.find((l) => l.id === m[1]);
      const top = parseInt(q.get('$top'), 10);
      if (q.get('$orderby') === 'Id desc') {
        const maxId = lib.big ? BIG - 1 : lib.items.length ? lib.items[lib.items.length - 1].id : 0;
        return resp(200, { value: maxId ? [{ Id: maxId }] : [] });
      }
      const lean = !!(options && options.headers && options.headers.Accept);
      if (opts.notAcceptable && ((q.get('$orderby') && opts.notAcceptable !== 'accept') || (lean && opts.notAcceptable === 'both'))) {
        notAcceptableCount++;
        return resp(406, '');
      }
      if (lib.alwaysThrottle) return resp(429, 'slow down', { 'Retry-After': '120' });
      if (lib.failing && q.get('$orderby') !== 'Id desc') return resp(500, { 'odata.error': { message: { value: 'Server is busy.' } } });
      if (lib.throttleOnce) {
        lib.throttleOnce = false;
        return resp(429, 'slow down', { 'Retry-After': '1' });
      }
      if (lib.big && rand() < opts.throttleRate) return resp(429, 'slow down', { 'Retry-After': '0' });
      if (lib.failBigPages && top > 2500) return resp(500, { 'odata.error': { message: { value: 'The operation has timed out.' } } });
      if (q.get('$expand') !== 'File') return resp(400, 'expected $expand=File');
      if (top > 5000) return resp(400, '$top above 5000');
      let gt = 0;
      let le = Infinity;
      const f = q.get('$filter');
      if (f) {
        const fm = /^Id gt (\d+)(?: and Id le (\d+))?$/.exec(f);
        if (!fm) return resp(400, `bad filter ${f}`);
        gt = +fm[1];
        if (fm[2]) le = +fm[2];
      }
      const skip = q.get('$skiptoken');
      const after = skip ? parseInt(/p_ID=(\d+)/.exec(skip)[1], 10) : 0;
      const pageItems = visibleItems(lib, Math.max(gt, after), le, top);
      if (lib.poisonId && pageItems.some((i) => i.id === lib.poisonId)) {
        return resp(500, { 'odata.error': { message: { value: 'Unexpected error reading item.' } } });
      }
      const body = {
        value: pageItems.map((i) => ({
          Id: i.id,
          FSObjType: i.folder ? 1 : 0,
          Modified: i.folder ? now.toISOString() : i.modified,
          FileRef: `${lib.url}/${i.name}`,
          File: i.folder ? null : { Length: String(i.size) }
        }))
      };
      const last = pageItems[pageItems.length - 1];
      if (last && visibleItems(lib, last.id, le, 1).length) {
        body['odata.nextLink'] =
          `${ORIGIN}${encodeURI(webPath)}/_api/web/lists(guid'${lib.id}')/items?%24skiptoken=Paged%3dTRUE%26p_ID%3d${last.id}` +
          `&%24select=Id%2cFSObjType%2cModified%2cFileRef%2cFile%2fLength&%24expand=File` +
          (f ? `&%24filter=${encodeURIComponent(f)}` : '') +
          `&%24top=${top}`;
      }
      return resp(200, body);
    }
    m = api.match(/^web\/lists\(guid'([^']+)'\)\/RootFolder$/);
    if (m) {
      const lib = web.libraries.find((l) => l.id === m[1]);
      if (lib.noMetrics) return resp(403, 'denied');
      if (lib.big) {
        return resp(200, { StorageMetrics: { TotalSize: String(900 * GB), TotalFileStreamSize: String(700 * GB), TotalFileCount: String(lib.itemCount), LastModified: now.toISOString() } });
      }
      const fileItems = lib.items.filter((i) => !i.folder);
      const cur = fileItems.reduce((s, i) => s + (i.size || 0), 0);
      return resp(200, {
        StorageMetrics: {
          TotalSize: String(Math.round(cur * 1.6)),
          TotalFileStreamSize: String(cur),
          TotalFileCount: String(fileItems.length),
          LastModified: lastChange(lib)
        }
      });
    }
    m = api.match(/^web\/GetFileByServerRelativeUrl\('(.+)'\)\/\$value$/);
    if (m) {
      const f = files[m[1]];
      return f ? resp(200, f) : resp(404, 'not found');
    }
    m = api.match(/^web\/GetFileByServerRelativeUrl\('(.+)'\)$/);
    if (m) {
      return files[m[1]] ? resp(200, { TimeLastModified: now.toISOString(), ModifiedBy: { Title: 'Megan Bowen' } }) : resp(404, 'not found');
    }
    m = api.match(/^web\/GetFolderByServerRelativeUrl\('(.+)'\)\/Files\/add\(url='(.+)',overwrite=true\)$/);
    if (m) {
      files[`${m[1]}/${m[2]}`] = options.body;
      return resp(200, {});
    }
    return resp(404, `unmocked ${api}`);
  }

  /** What a scan with the given options should find: visible files and bytes, and hidden items. */
  function expected(scanOpts = {}) {
    let fileCount = 0;
    let bytes = 0;
    let hidden = 0;
    const excluded = (scanOpts.excludedLibraries || []).map((n) => n.toLowerCase());
    for (const w of Object.values(webs)) {
      for (const l of w.libraries || []) {
        if (l.isCatalog || l.alwaysThrottle || l.poisonId) continue;
        if (l.hidden && !scanOpts.includeHidden) continue;
        if (scanOpts.excludeSystemLibraries && (l.baseTemplate === 119 || /\/(SitePages|SiteAssets)$/.test(l.url))) continue;
        if (excluded.indexOf(l.title.toLowerCase()) >= 0) continue;
        if (l.big) {
          for (let id = 1; id <= BIG; id++) {
            const i = bigItem(id);
            if (!i) continue;
            if (i.hidden) {
              hidden++;
              continue;
            }
            if (!i.folder) {
              fileCount++;
              bytes += i.size;
            }
          }
          continue;
        }
        for (const i of l.items) {
          if (!i.folder) {
            fileCount++;
            bytes += i.size;
          }
        }
      }
    }
    return { files: fileCount, bytes, hidden };
  }

  const context = {
    spHttpClient: {
      get: (url, cfg, options) => {
        acceptHeaders.add(options && options.headers && options.headers.Accept);
        inFlight++;
        maxInFlight = Math.max(maxInFlight, inFlight);
        return new Promise((r) =>
          setTimeout(() => {
            inFlight--;
            r(handle(url, 'GET', options));
          }, opts.latencyMs)
        );
      },
      post: (url, cfg, options) => handle(url, 'POST', options)
    },
    pageContext: {
      site: { absoluteUrl: `${ORIGIN}${SITE}`, serverRelativeUrl: SITE },
      web: { absoluteUrl: `${ORIGIN}${SITE}`, title: 'Finance', permissions: { hasPermission: () => opts.owner } },
      user: { displayName: 'Alex Wilber' },
      legacyPageContext: { isSiteAdmin: opts.owner }
    }
  };

  return {
    context,
    log,
    files,
    acceptHeaders,
    expected,
    now,
    libs,
    heal: () => {
      if (libs.flaky) libs.flaky.failing = false;
    },
    maxInFlight: () => maxInFlight,
    notAcceptableCount: () => notAcceptableCount
  };
}
