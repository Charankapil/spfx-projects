import { readFile } from 'node:fs/promises';
import { parseCsv } from './csv.js';
import { PrincipalRequest } from './graph.js';
import { joinPath, serverRelative, trimSlash } from './sharepoint.js';

/**
 * Settings come from a JSON file (--config), from ACCESS_REPORT_CONFIG (the same JSON as text),
 * and from single environment variables, which win. Secrets only ever come from the environment
 * or from files named in the settings, never from the settings themselves being printed.
 */

export class ConfigError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ConfigError';
  }
}

export interface Config {
  tenantId: string;
  clientId: string;
  auth: {
    mode: 'certificate' | 'federated';
    /** PEM text holding the certificate and its private key. */
    certificate?: string;
    password?: string;
  };
  endpoints: { login: string; graph: string };
  report: {
    /** The site (or subsite) that holds the report library. */
    siteUrl: string;
    /** Server-relative path of the library ("/sites/gov/Shared Documents"). Never created. */
    libraryPath: string;
    /** Folders below the library, created when missing ("Access reports"). May be ''. */
    folder: string;
    /** File names start with this. */
    name: string;
  };
  principals: PrincipalRequest[];
  /** A CSV in the report folder that lists more principals (Type, Identity). '' disables it. */
  principalsFile: string;
  sites: {
    /** Scan exactly these instead of asking Graph for every site. */
    only: string[];
    /** Wildcard patterns (*) a site address must match (any of them). Empty = all. */
    include: string[];
    exclude: string[];
    includeOneDrive: boolean;
  };
  scan: {
    includeLibraries: boolean;
    includeLists: boolean;
    includeSubsites: boolean;
    expandGroups: boolean;
    thorough: boolean;
    concurrency: number;
    /** Stop cleanly after this long and carry on in the next run. 0 = no limit. */
    maxMinutes: number;
    checkpointMinutes: number;
    maxSubsites: number;
  };
}

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function isLocal(url: URL): boolean {
  return url.hostname === 'localhost' || url.hostname === '127.0.0.1';
}

/** An https address (http only for localhost, which the tests use). */
export function checkedUrl(value: string, what: string): string {
  let u: URL;
  try {
    u = new URL(value.trim());
  } catch {
    throw new ConfigError(`${what} is not a valid address: "${value}"`);
  }
  if (u.protocol !== 'https:' && !(u.protocol === 'http:' && isLocal(u))) {
    throw new ConfigError(`${what} must start with https://`);
  }
  return trimSlash(u.origin + u.pathname);
}

/**
 * Splits the address of a folder into site, library and folder.
 * Understands a plain folder address and the address of a library view (…/Forms/AllItems.aspx?id=/sites/…).
 * The site is taken as /sites/<name> or /teams/<name>, else the root site; for a subsite give siteUrl and folder separately.
 */
export function splitFolderUrl(folderUrl: string): { siteUrl: string; libraryPath: string; folder: string } {
  let u: URL;
  try {
    u = new URL(folderUrl.trim());
  } catch {
    throw new ConfigError(`The report folder is not a valid address: "${folderUrl}"`);
  }
  let path = serverRelative(u.origin + u.pathname);
  const id = u.searchParams.get('id') || u.searchParams.get('RootFolder');
  if (id && id.charAt(0) === '/') {
    path = serverRelative(u.origin + encodeURI(id));
  } else {
    path = path.replace(/\/Forms\/[^/]+\.aspx$/i, '');
  }
  if (/\.aspx$/i.test(path) || /\/_layouts\//i.test(path)) {
    throw new ConfigError('The report folder must be the address of a folder in a document library, not a page.');
  }
  const segs = path.split('/').filter((s) => !!s);
  const managed = segs.length >= 2 && /^(sites|teams)$/i.test(segs[0]);
  const siteSegs = managed ? segs.slice(0, 2) : [];
  const rest = segs.slice(siteSegs.length);
  if (!rest.length) {
    throw new ConfigError('The report folder must include the document library, e.g. https://contoso.sharepoint.com/sites/governance/Shared Documents/Access reports');
  }
  const sitePath = siteSegs.length ? '/' + siteSegs.join('/') : '';
  return {
    siteUrl: checkedUrl(u.origin + encodeURI(sitePath), 'The report site'),
    libraryPath: joinPath(sitePath, rest[0]),
    folder: rest.slice(1).join('/')
  };
}

/** "user:alice@contoso.com; group:Finance team" or one per line. A bare e-mail is a user. */
export function parsePrincipalList(text: string): PrincipalRequest[] {
  const out: PrincipalRequest[] = [];
  text
    .split(/[\r\n;]+/)
    .map((s) => s.trim())
    .filter((s) => !!s && s.charAt(0) !== '#')
    .forEach((s) => {
      const m = /^(user|person|group)\s*[:=]\s*(.+)$/i.exec(s);
      if (m) {
        out.push({ type: /^group$/i.test(m[1]) ? 'group' : 'user', id: m[2].trim() });
      } else {
        out.push({ type: s.indexOf('@') > 0 ? 'user' : 'group', id: s });
      }
    });
  return out;
}

/** The principals file: a header row, then Type and Identity columns (any order; extra columns ignored). */
export function parsePrincipalsCsv(text: string): { principals: PrincipalRequest[]; problems: string[] } {
  const rows = parseCsv(text);
  const principals: PrincipalRequest[] = [];
  const problems: string[] = [];
  if (!rows.length) {
    return { principals, problems };
  }
  const head = rows[0].map((h) => h.trim().toLowerCase());
  const typeCol = head.findIndex((h) => h === 'type' || h === 'kind');
  const idCol = head.findIndex((h) => ['identity', 'id', 'e-mail', 'email', 'upn', 'name', 'who', 'user or group'].indexOf(h) >= 0);
  if (idCol < 0) {
    problems.push('The principals file needs a header row with the columns "Type" and "Identity".');
    return { principals, problems };
  }
  rows.slice(1).forEach((r, i) => {
    const id = (r[idCol] || '').trim();
    if (!id || id.charAt(0) === '#' || (r[0] || '').trim().charAt(0) === '#') {
      return;
    }
    const t = (typeCol >= 0 ? r[typeCol] || '' : '').trim().toLowerCase();
    if (t === 'user' || t === 'person' || (!t && id.indexOf('@') > 0)) {
      principals.push({ type: 'user', id });
    } else if (t === 'group' || t === 'security group' || t === 'microsoft 365 group' || t === 'm365' || !t) {
      principals.push({ type: 'group', id });
    } else {
      problems.push(`Row ${i + 2} of the principals file: Type must be "user" or "group", not "${r[typeCol]}".`);
    }
  });
  return { principals, problems };
}

export const PRINCIPALS_TEMPLATE = 'Type,Identity,Notes\r\n';

type Dict = Record<string, unknown>;

function dict(v: unknown, what: string): Dict {
  if (v === undefined || v === null) {
    return {};
  }
  if (typeof v !== 'object' || Array.isArray(v)) {
    throw new ConfigError(`"${what}" must be an object.`);
  }
  return v as Dict;
}

function known(o: Dict, keys: string[], what: string): void {
  Object.keys(o).forEach((k) => {
    if (keys.indexOf(k) < 0 && k.charAt(0) !== '$' && k.charAt(0) !== '_') {
      throw new ConfigError(`Unknown setting "${what ? what + '.' : ''}${k}". Allowed: ${keys.join(', ')}.`);
    }
  });
}

function bool(v: unknown, def: boolean, what: string): boolean {
  if (v === undefined || v === null || v === '') {
    return def;
  }
  if (typeof v === 'boolean') {
    return v;
  }
  if (/^(true|yes|1|on)$/i.test(String(v))) {
    return true;
  }
  if (/^(false|no|0|off)$/i.test(String(v))) {
    return false;
  }
  throw new ConfigError(`"${what}" must be true or false.`);
}

function num(v: unknown, def: number, min: number, max: number, what: string): number {
  if (v === undefined || v === null || v === '') {
    return def;
  }
  const n = Number(v);
  if (!isFinite(n) || n < min || n > max) {
    throw new ConfigError(`"${what}" must be a number between ${min} and ${max}.`);
  }
  return n;
}

function strings(v: unknown, what: string): string[] {
  if (v === undefined || v === null || v === '') {
    return [];
  }
  if (typeof v === 'string') {
    return v
      .split(/[\r\n;,]+/)
      .map((s) => s.trim())
      .filter((s) => !!s);
  }
  if (!Array.isArray(v) || v.some((x) => typeof x !== 'string')) {
    throw new ConfigError(`"${what}" must be a list of text values.`);
  }
  return (v as string[]).map((s) => s.trim()).filter((s) => !!s);
}

function principalsOf(v: unknown): PrincipalRequest[] {
  if (v === undefined || v === null || v === '') {
    return [];
  }
  if (typeof v === 'string') {
    return parsePrincipalList(v);
  }
  if (!Array.isArray(v)) {
    throw new ConfigError('"principals" must be a list.');
  }
  const out: PrincipalRequest[] = [];
  v.forEach((item, i) => {
    if (typeof item === 'string') {
      parsePrincipalList(item).forEach((p) => out.push(p));
      return;
    }
    const o = dict(item, `principals[${i}]`);
    known(o, ['type', 'id', 'note'], `principals[${i}]`);
    const type = String(o.type || '').toLowerCase();
    const id = String(o.id || '').trim();
    if ((type !== 'user' && type !== 'group') || !id) {
      throw new ConfigError(`principals[${i}] needs "type" ("user" or "group") and "id".`);
    }
    out.push({ type: type as 'user' | 'group', id });
  });
  return out;
}

export type Env = Record<string, string | undefined>;

/** Builds the settings. `file` is the parsed JSON file (or {}); `readText` reads files named in the settings. */
export async function buildConfig(file: unknown, env: Env, readText: (path: string) => Promise<string> = (p) => readFile(p, 'utf8')): Promise<Config> {
  const f = dict(file, 'config');
  known(f, ['tenantId', 'clientId', 'auth', 'endpoints', 'report', 'principals', 'principalsFile', 'sites', 'scan'], '');
  const auth = dict(f.auth, 'auth');
  known(auth, ['mode', 'certificatePath', 'privateKeyPath'], 'auth');
  const endpoints = dict(f.endpoints, 'endpoints');
  known(endpoints, ['login', 'graph'], 'endpoints');
  const report = dict(f.report, 'report');
  known(report, ['folderUrl', 'siteUrl', 'folder', 'name'], 'report');
  const sites = dict(f.sites, 'sites');
  known(sites, ['only', 'include', 'exclude', 'includeOneDrive'], 'sites');
  const scan = dict(f.scan, 'scan');
  known(scan, ['includeLibraries', 'includeLists', 'includeSubsites', 'expandGroups', 'thorough', 'concurrency', 'maxMinutes', 'checkpointMinutes', 'maxSubsites'], 'scan');

  const tenantId = String(env.ACCESS_REPORT_TENANT_ID || f.tenantId || '').trim();
  const clientId = String(env.ACCESS_REPORT_CLIENT_ID || f.clientId || '').trim();
  if (!tenantId) {
    throw new ConfigError('The tenant id is missing. Set ACCESS_REPORT_TENANT_ID (or "tenantId" in the config file).');
  }
  if (!GUID.test(tenantId) && !/^[a-z0-9.-]+\.[a-z]{2,}$/i.test(tenantId)) {
    throw new ConfigError('The tenant id must be the directory (tenant) id, a GUID, or the tenant\'s domain name.');
  }
  if (!GUID.test(clientId)) {
    throw new ConfigError('The client id is missing or not a GUID. Set ACCESS_REPORT_CLIENT_ID to the application (client) id of the app registration.');
  }

  const mode = String(env.ACCESS_REPORT_AUTH || auth.mode || 'certificate').toLowerCase();
  if (mode !== 'certificate' && mode !== 'federated') {
    throw new ConfigError('"auth.mode" must be "certificate" or "federated".');
  }
  let certificate: string | undefined;
  if (mode === 'certificate') {
    const parts: string[] = [];
    if (env.ACCESS_REPORT_CERTIFICATE) {
      parts.push(env.ACCESS_REPORT_CERTIFICATE);
    }
    if (env.ACCESS_REPORT_PRIVATE_KEY) {
      parts.push(env.ACCESS_REPORT_PRIVATE_KEY);
    }
    if (!parts.length) {
      for (const key of ['certificatePath', 'privateKeyPath']) {
        const p = auth[key];
        if (typeof p === 'string' && p) {
          try {
            parts.push(await readText(p));
          } catch (e) {
            throw new ConfigError(`"auth.${key}": cannot read ${p}: ${(e as Error).message}`);
          }
        }
      }
    }
    if (!parts.length) {
      throw new ConfigError('The certificate is missing. Set ACCESS_REPORT_CERTIFICATE to the PEM text (certificate and private key), or "auth.certificatePath" in the config file.');
    }
    // Stored secrets sometimes lose their line breaks or arrive base64-encoded: each part is normalised later.
    certificate = parts.join('\n');
  }

  const folderUrl = String(env.ACCESS_REPORT_FOLDER_URL || report.folderUrl || '').trim();
  let siteUrl: string;
  let libraryPath: string;
  let folder: string;
  if (folderUrl) {
    if (report.siteUrl || report.folder) {
      if (!env.ACCESS_REPORT_FOLDER_URL) {
        throw new ConfigError('Give either "report.folderUrl", or "report.siteUrl" with "report.folder", not both.');
      }
    }
    ({ siteUrl, libraryPath, folder } = splitFolderUrl(folderUrl));
  } else {
    if (!report.siteUrl || !report.folder) {
      throw new ConfigError('Where should the report go? Set ACCESS_REPORT_FOLDER_URL to the address of a folder in a document library (or "report.siteUrl" and "report.folder").');
    }
    siteUrl = checkedUrl(String(report.siteUrl), '"report.siteUrl"');
    const segs = String(report.folder)
      .split('/')
      .map((s) => s.trim())
      .filter((s) => !!s);
    if (!segs.length) {
      throw new ConfigError('"report.folder" must start with the document library, e.g. "Shared Documents/Access reports".');
    }
    libraryPath = joinPath(serverRelative(siteUrl), segs[0]);
    folder = segs.slice(1).join('/');
  }
  const name = String(report.name || 'Access report').trim();
  if (!name || /["*:<>?/\\|#%]/.test(name) || name.length > 80) {
    throw new ConfigError('"report.name" must be a short file name without any of " * : < > ? / \\ | # %.');
  }

  const principals = principalsOf(f.principals).concat(env.ACCESS_REPORT_PRINCIPALS ? parsePrincipalList(env.ACCESS_REPORT_PRINCIPALS) : []);
  const principalsFile = f.principalsFile === undefined ? 'principals.csv' : String(f.principalsFile || '').trim();
  if (principalsFile && (/[/\\]/.test(principalsFile) || !/\.csv$/i.test(principalsFile))) {
    throw new ConfigError('"principalsFile" is the name of a .csv file in the report folder (no path).');
  }

  const only = strings(sites.only, 'sites.only').map((u) => checkedUrl(u, 'An address in "sites.only"'));

  return {
    tenantId,
    clientId,
    auth: { mode, certificate, password: env.ACCESS_REPORT_CERTIFICATE_PASSWORD || undefined },
    endpoints: {
      login: checkedUrl(String(env.ACCESS_REPORT_LOGIN_URL || endpoints.login || 'https://login.microsoftonline.com'), '"endpoints.login"'),
      graph: checkedUrl(String(env.ACCESS_REPORT_GRAPH_URL || endpoints.graph || 'https://graph.microsoft.com'), '"endpoints.graph"')
    },
    report: { siteUrl, libraryPath, folder, name },
    principals,
    principalsFile,
    sites: {
      only,
      include: strings(sites.include, 'sites.include'),
      exclude: strings(sites.exclude, 'sites.exclude'),
      includeOneDrive: bool(sites.includeOneDrive, false, 'sites.includeOneDrive')
    },
    scan: {
      includeLibraries: bool(scan.includeLibraries, true, 'scan.includeLibraries'),
      includeLists: bool(scan.includeLists, false, 'scan.includeLists'),
      includeSubsites: bool(scan.includeSubsites, true, 'scan.includeSubsites'),
      expandGroups: bool(scan.expandGroups, true, 'scan.expandGroups'),
      thorough: bool(scan.thorough, true, 'scan.thorough'),
      concurrency: Math.round(num(env.ACCESS_REPORT_CONCURRENCY || scan.concurrency, 2, 1, 8, 'scan.concurrency')),
      maxMinutes: num(env.ACCESS_REPORT_MAX_MINUTES || scan.maxMinutes, 300, 0, 100000, 'scan.maxMinutes'),
      checkpointMinutes: num(scan.checkpointMinutes, 10, 0.001, 600, 'scan.checkpointMinutes'),
      maxSubsites: Math.round(num(scan.maxSubsites, 2000, 1, 100000, 'scan.maxSubsites'))
    }
  };
}

export async function loadConfig(path: string | undefined, env: Env): Promise<Config> {
  let file: unknown = {};
  const inline = env.ACCESS_REPORT_CONFIG;
  if (path) {
    let text: string;
    try {
      text = await readFile(path, 'utf8');
    } catch (e) {
      throw new ConfigError(`Cannot read the config file ${path}: ${(e as Error).message}`);
    }
    try {
      file = JSON.parse(text.charCodeAt(0) === 0xfeff ? text.substring(1) : text);
    } catch (e) {
      throw new ConfigError(`The config file ${path} is not valid JSON: ${(e as Error).message}`);
    }
  } else if (inline && inline.trim()) {
    try {
      file = JSON.parse(inline);
    } catch (e) {
      throw new ConfigError(`ACCESS_REPORT_CONFIG is not valid JSON: ${(e as Error).message}`);
    }
  }
  return buildConfig(file, env);
}

/** "https://contoso.sharepoint.com/sites/hr*" style patterns, case-insensitive. */
export function wildcard(pattern: string): RegExp {
  const esc = trimSlash(pattern.trim())
    .replace(/[.+?^${}()|[\]\\]/g, '\\$&')
    .replace(/\*/g, '.*');
  return new RegExp(`^${esc}$`, 'i');
}

export function siteFilter(include: string[], exclude: string[]): (url: string) => boolean {
  const inc = include.map(wildcard);
  const exc = exclude.map(wildcard);
  return (url) => {
    const u = trimSlash(url);
    return (!inc.length || inc.some((r) => r.test(u))) && !exc.some((r) => r.test(u));
  };
}
