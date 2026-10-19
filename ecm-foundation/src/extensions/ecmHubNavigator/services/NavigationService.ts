import { SPHttpClient } from '@microsoft/sp-http';

import { normalizeUrl, SpRestClient } from '../../../common/services/SpRestClient';
import { hubColor } from '../../../common/theme/ecmTheme';
import { IHub, INavigationLocation, INavigationModel, INavLink, NavItemType } from '../models/INavigation';

interface INavListItem {
  Id: number;
  Title: string;
  EcmNavType: NavItemType | null;
  EcmParentHub: string | null;
  EcmUrl: { Url: string; Description: string } | string | null;
  EcmDescription: string | null;
  EcmOrder: number | null;
  EcmRestricted: boolean | null;
  EcmColor: string | null;
}

interface ICacheEntry<T> {
  savedAt: number;
  value: T;
}

const CACHE_VERSION = 'v1';
const SHARED_HUB_ID = -1;

function readCache<T>(storage: Storage, key: string, maxAgeMs: number): T | undefined {
  try {
    const raw = storage.getItem(key);
    if (!raw) {
      return undefined;
    }
    const entry = JSON.parse(raw) as ICacheEntry<T>;
    return Date.now() - entry.savedAt <= maxAgeMs ? entry.value : undefined;
  } catch {
    return undefined;
  }
}

function writeCache<T>(storage: Storage, key: string, value: T): void {
  try {
    storage.setItem(key, JSON.stringify({ savedAt: Date.now(), value }));
  } catch {
    // Storage full or blocked (private mode): the navigator still works, it just reloads.
  }
}

/** The site collection a link lives in: search returns sites, not arbitrary pages. */
function siteUrlOf(url: string): string {
  const match = /^(https?:\/\/[^/]+\/(?:sites|teams)\/[^/?#]+)/i.exec(url);
  if (match) {
    return match[1].toLowerCase();
  }
  const origin = /^(https?:\/\/[^/]+)/i.exec(url);
  return origin ? origin[1].toLowerCase() : url.toLowerCase();
}

export interface INavigationServiceOptions {
  /** Absolute URL of the site holding the navigation list (the Corporate ECM Hub). */
  configSiteUrl: string;
  listTitle: string;
  cacheMinutes: number;
  sharedServicesTitle: string;
  userKey: string;
}

/**
 * Loads the hub/department structure from one SharePoint list, so governance
 * edits the menu without redeploying anything. The list is cached in
 * localStorage (it is the same for everyone) and the user's access check in
 * sessionStorage (it is not).
 */
export class NavigationService {
  private rest: SpRestClient;
  private origin: string;

  constructor(httpClient: SPHttpClient, private options: INavigationServiceOptions) {
    this.rest = new SpRestClient(httpClient, 'ECM Hub Navigator');
    this.origin = /^(https?:\/\/[^/]+)/i.exec(options.configSiteUrl)?.[1] || '';
  }

  private get modelCacheKey(): string {
    return `ecm-nav:${CACHE_VERSION}:${this.options.configSiteUrl.toLowerCase()}:${this.options.listTitle}`;
  }

  private get accessCacheKey(): string {
    return `ecm-nav-access:${CACHE_VERSION}:${this.options.userKey}`;
  }

  private get maxAgeMs(): number {
    return Math.max(0, this.options.cacheMinutes) * 60 * 1000;
  }

  /** The cached menu, already trimmed, if both caches are fresh — lets the header paint instantly. */
  public getCached(): INavigationModel | undefined {
    const model = readCache<INavigationModel>(localStorage, this.modelCacheKey, this.maxAgeMs);
    const access = readCache<string[]>(sessionStorage, this.accessCacheKey, this.maxAgeMs);
    return model && access ? this.trim(model, access) : undefined;
  }

  public async load(): Promise<INavigationModel> {
    const model = await this.loadModel();
    writeCache(localStorage, this.modelCacheKey, model);
    const access = await this.loadAccessibleSites(model);
    writeCache(sessionStorage, this.accessCacheKey, access);
    return this.trim(model, access);
  }

  private async loadModel(): Promise<INavigationModel> {
    const select = 'Id,Title,EcmNavType,EcmParentHub,EcmUrl,EcmDescription,EcmOrder,EcmRestricted,EcmColor';
    const url =
      `${this.options.configSiteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.options.listTitle.replace(/'/g, "''"))}')/items` +
      `?$select=${select}&$filter=EcmActive eq 1&$top=2000`;
    const json = await this.rest.getJson<{ value: INavListItem[] }>(url);
    return this.buildModel(json.value || []);
  }

  private toLink(item: INavListItem): INavLink {
    const rawUrl = typeof item.EcmUrl === 'string' ? item.EcmUrl : item.EcmUrl?.Url || '';
    const url = rawUrl.indexOf('/') === 0 ? this.origin + rawUrl : rawUrl;
    return {
      id: item.Id,
      title: item.Title || '',
      url,
      description: item.EcmDescription || '',
      restricted: !!item.EcmRestricted,
      order: typeof item.EcmOrder === 'number' ? item.EcmOrder : 1000
    };
  }

  private buildModel(items: INavListItem[]): INavigationModel {
    const byOrder = (a: INavLink, b: INavLink): number => a.order - b.order || a.title.localeCompare(b.title);
    const ofType = (type: NavItemType): INavListItem[] => items.filter((i) => i.EcmNavType === type);

    const hubItems = ofType('Hub')
      .map((item) => ({ item, link: this.toLink(item) }))
      .sort((a, b) => byOrder(a.link, b.link));

    const hubs: IHub[] = hubItems.map(({ item, link }, index) => ({
      ...link,
      color: hubColor(link.title, item.EcmColor || undefined, index),
      sites: [],
      resources: [],
      isSharedServices: false
    }));
    const hubByTitle: { [title: string]: IHub } = {};
    hubs.forEach((hub) => {
      hubByTitle[hub.title.toLowerCase()] = hub;
    });
    const parentOf = (item: INavListItem): IHub | undefined => hubByTitle[(item.EcmParentHub || '').trim().toLowerCase()];

    ofType('Department site').forEach((item) => parentOf(item)?.sites.push(this.toLink(item)));
    ofType('Hub resource').forEach((item) => parentOf(item)?.resources.push(this.toLink(item)));
    hubs.forEach((hub) => {
      hub.sites.sort(byOrder);
      hub.resources.sort(byOrder);
    });

    const sharedServices = ofType('Shared service').map((i) => this.toLink(i)).sort(byOrder);
    if (sharedServices.length) {
      hubs.push({
        id: SHARED_HUB_ID,
        title: this.options.sharedServicesTitle,
        url: '',
        description: 'Enterprise-wide capabilities used by every department; department-owned content stays in department sites.',
        restricted: false,
        order: 1e9,
        color: hubColor(this.options.sharedServicesTitle, undefined, 6),
        sites: sharedServices,
        resources: [],
        isSharedServices: true
      });
    }

    return { hubs, sharedServices, headerLinks: ofType('Header link').map((i) => this.toLink(i)).sort(byOrder) };
  }

  /**
   * Restricted entries stay visible only when search — which is
   * security-trimmed — returns their site to this user. Only restricted
   * sites are checked, so a typical menu costs one small query. If the query
   * fails the restricted entries are hidden (fail closed).
   */
  private async loadAccessibleSites(model: INavigationModel): Promise<string[]> {
    const restricted: string[] = [];
    const collect = (link: INavLink): void => {
      if (link.restricted && link.url) {
        const site = siteUrlOf(link.url);
        if (restricted.indexOf(site) < 0) {
          restricted.push(site);
        }
      }
    };
    model.hubs.forEach((hub) => {
      collect(hub);
      hub.sites.forEach(collect);
      hub.resources.forEach(collect);
    });
    model.headerLinks.forEach(collect);
    if (!restricted.length) {
      return [];
    }

    const accessible: string[] = [];
    // Keep each KQL query well inside the search URL length limit.
    for (let i = 0; i < restricted.length; i += 20) {
      const chunk = restricted.slice(i, i + 20);
      const kql = `contentclass:STS_Site AND (${chunk.map((u) => `Path:"${u}"`).join(' OR ')})`;
      try {
        const page = await this.rest.search(this.options.configSiteUrl, kql, {
          selectProperties: ['Path', 'SPSiteUrl'],
          rowLimit: chunk.length * 2
        });
        page.rows.forEach((row) => {
          const site = normalizeUrl(row.SPSiteUrl || row.Path || '', this.origin);
          if (site && accessible.indexOf(site) < 0) {
            accessible.push(site);
          }
        });
      } catch {
        // Fail closed for this chunk.
      }
    }
    return accessible;
  }

  private trim(model: INavigationModel, accessibleSites: string[]): INavigationModel {
    const visible = (link: INavLink): boolean => !link.restricted || accessibleSites.indexOf(siteUrlOf(link.url)) >= 0;
    const hubs = model.hubs
      .filter(visible)
      .map((hub) => ({ ...hub, sites: hub.sites.filter(visible), resources: hub.resources.filter(visible) }));
    return {
      hubs,
      sharedServices: model.sharedServices.filter(visible),
      headerLinks: model.headerLinks.filter(visible)
    };
  }

  /** Which hub and department site the current page belongs to, by URL. */
  public locate(model: INavigationModel, currentSiteUrl: string, hubSiteUrl?: string): INavigationLocation {
    const current = normalizeUrl(currentSiteUrl, this.origin);
    const hubUrl = hubSiteUrl ? normalizeUrl(hubSiteUrl, this.origin) : '';
    for (const hub of model.hubs) {
      const site = hub.sites.filter((s) => s.url && normalizeUrl(s.url, this.origin) === current)[0];
      if (site) {
        return { hubId: hub.id, siteId: site.id };
      }
    }
    for (const hub of model.hubs) {
      const url = hub.url ? normalizeUrl(hub.url, this.origin) : '';
      if (url && (url === current || url === hubUrl)) {
        return { hubId: hub.id };
      }
    }
    return {};
  }

  /** The hub site a site is associated with, from SharePoint itself (for sites not listed in the menu). */
  public async getHubSiteUrl(webUrl: string, hubSiteId: string): Promise<string | undefined> {
    try {
      const json = await this.rest.getJson<{ SiteUrl?: string }>(
        `${webUrl}/_api/HubSites/GetById?hubSiteId='${encodeURIComponent(hubSiteId)}'`
      );
      return json.SiteUrl;
    } catch {
      return undefined;
    }
  }
}
