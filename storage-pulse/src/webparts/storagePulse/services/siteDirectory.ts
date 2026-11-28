import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISiteEntry {
  url: string;
  title: string;
}

export type SiteCheck =
  | { ok: true; url: string; title: string }
  | { ok: false; reason: 'invalid' | 'host' | 'access' | 'notfound' | 'error'; detail?: string };

const JSON_OPTIONS = { headers: { Accept: 'application/json;odata.metadata=nometadata' } };
const PAGE = 500;
const MAX_SITES = 5000;

interface ICell {
  Key: string;
  Value: string | null;
}
type Many<T> = T[] | { results?: T[] } | undefined | null;

function many<T>(value: Many<T>): T[] {
  if (!value) {
    return [];
  }
  return Array.isArray(value) ? value : value.results || [];
}

/**
 * Every site collection the signed-in person can open, from SharePoint search
 * (security trimmed, so nobody sees sites they could not scan anyway).
 * OneDrive sites live on another host and are left out.
 */
export async function findSites(context: WebPartContext): Promise<ISiteEntry[]> {
  const webUrl = context.pageContext.web.absoluteUrl;
  const origin = new URL(webUrl).origin;
  const found = new Map<string, ISiteEntry>();
  for (let start = 0; start < MAX_SITES; start += PAGE) {
    const response = await context.spHttpClient.get(
      `${webUrl}/_api/search/query?querytext='contentclass:STS_Site'` +
        `&selectproperties='${encodeURIComponent('Title,SPSiteUrl,Path')}'` +
        `&rowlimit=${PAGE}&startrow=${start}&trimduplicates=false&enablequeryrules=false`,
      SPHttpClient.configurations.v1,
      JSON_OPTIONS
    );
    if (!response.ok) {
      throw new Error(`SharePoint search answered ${response.status}`);
    }
    const json = (await response.json()) as {
      d?: { query?: unknown };
      PrimaryQueryResult?: { RelevantResults?: { TotalRows?: number; Table?: { Rows?: Many<{ Cells?: Many<ICell> }> } } };
    };
    const body = (json.d && json.d.query ? json.d.query : json) as typeof json;
    const relevant = (body.PrimaryQueryResult && body.PrimaryQueryResult.RelevantResults) || {};
    const rows = many(relevant.Table ? relevant.Table.Rows : undefined);
    for (const row of rows) {
      const cells: Record<string, string> = {};
      many(row.Cells).forEach((c) => (cells[c.Key] = c.Value || ''));
      const raw = cells.SPSiteUrl || cells.Path;
      let url: URL;
      try {
        url = new URL(raw);
      } catch {
        continue;
      }
      if (url.protocol !== 'https:' || url.origin !== origin) {
        continue;
      }
      const clean = `${url.origin}${url.pathname}`.replace(/\/+$/, '');
      if (!found.has(clean.toLowerCase())) {
        found.set(clean.toLowerCase(), { url: clean, title: cells.Title || decodeURIComponent(url.pathname) || clean });
      }
    }
    if (rows.length < PAGE || (relevant.TotalRows !== undefined && start + PAGE >= relevant.TotalRows)) {
      break;
    }
  }
  const sites: ISiteEntry[] = [];
  found.forEach((s) => sites.push(s));
  sites.sort((a, b) => a.title.localeCompare(b.title));
  return sites;
}

/**
 * Turns an address someone pasted (a site, a page in it, a library...) into
 * its site collection, and checks the person can open it.
 */
export async function checkSite(context: WebPartContext, address: string): Promise<SiteCheck> {
  const origin = new URL(context.pageContext.web.absoluteUrl).origin;
  let url: URL;
  try {
    url = new URL(address.trim());
  } catch {
    return { ok: false, reason: 'invalid' };
  }
  if (url.protocol !== 'https:') {
    return { ok: false, reason: 'invalid' };
  }
  if (url.origin !== origin) {
    return { ok: false, reason: 'host' };
  }
  const path = decodeURIComponent(url.pathname);
  const managed = /^\/(sites|teams)\/[^/]+/i.exec(path);
  const base = `${origin}${managed ? encodeURI(managed[0]) : ''}`;

  const get = async (u: string): Promise<{ status: number; json?: Record<string, unknown> }> => {
    const response = await context.spHttpClient.get(u, SPHttpClient.configurations.v1, JSON_OPTIONS);
    return { status: response.status, json: response.ok ? ((await response.json()) as Record<string, unknown>) : undefined };
  };
  try {
    const site = await get(`${base}/_api/site?$select=Url`);
    if (site.status === 401 || site.status === 403) {
      return { ok: false, reason: 'access' };
    }
    if (site.status === 404) {
      return { ok: false, reason: 'notfound' };
    }
    if (!site.json || typeof site.json.Url !== 'string') {
      return { ok: false, reason: 'error', detail: String(site.status) };
    }
    const siteUrl = (site.json.Url as string).replace(/\/+$/, '');
    if (new URL(siteUrl).origin !== origin) {
      return { ok: false, reason: 'host' };
    }
    const web = await get(`${siteUrl}/_api/web?$select=Title`);
    if (web.status === 401 || web.status === 403) {
      return { ok: false, reason: 'access' };
    }
    const title = web.json && typeof web.json.Title === 'string' ? (web.json.Title as string) : siteUrl;
    return { ok: true, url: siteUrl, title };
  } catch (err) {
    return { ok: false, reason: 'error', detail: err instanceof Error ? err.message : String(err) };
  }
}
