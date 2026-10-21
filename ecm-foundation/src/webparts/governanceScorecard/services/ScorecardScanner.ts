import { SPHttpClient } from '@microsoft/sp-http';

import { SpRestClient } from '../../../common/services/SpRestClient';
import { NavigationService } from '../../../extensions/ecmHubNavigator/services/NavigationService';
import { LibrariesSource } from '../../metadataExplorer/services/LibrariesSource';
import { IScanProgress, IScorecardRun, ISiteMetrics, OwnershipKind } from '../models/IScorecard';
import { measureSite } from './metrics';

export interface IScanOptions {
  /** Site holding the ECM Navigation list. */
  configSiteUrl: string;
  navListTitle: string;
  requiredFields: string[];
  staleYears: number;
  maxDocsPerSite: number;
}

export interface IScanTarget {
  hub: string;
  title: string;
  url: string;
}

interface IOwnerEntry {
  PrincipalType: number;
}

const SITE_ROOT = /^(https?:\/\/[^/]+(?:\/(?:sites|teams)\/[^/?#]+)?)/i;
const PARALLEL_SITES = 2;
/** SharePoint principal type for a security group (Entra ID or Microsoft 365 group claims). */
const SECURITY_GROUP = 4;

/**
 * Scans every site in the ECM Navigation list: its document libraries (via
 * the Metadata Explorer's library reader), its storage and its owners. Runs
 * as the person viewing the page, so it measures what they can open; sites
 * they cannot are reported, not skipped silently.
 */
export class ScorecardScanner {
  private rest: SpRestClient;

  constructor(private httpClient: SPHttpClient, private options: IScanOptions) {
    this.rest = new SpRestClient(httpClient, 'Governance Scorecard');
  }

  /** Hubs, department sites and shared services from the navigation list, one entry per site. */
  public async getTargets(): Promise<IScanTarget[]> {
    const nav = new NavigationService(this.httpClient, {
      configSiteUrl: this.options.configSiteUrl,
      listTitle: this.options.navListTitle,
      cacheMinutes: 0,
      sharedServicesTitle: 'Shared Services',
      userKey: 'scorecard'
    });
    const model = await nav.load();
    const targets: IScanTarget[] = [];
    const seen: { [url: string]: boolean } = {};
    const add = (hub: string, title: string, url: string): void => {
      const root = SITE_ROOT.exec(url || '');
      if (!root) {
        return;
      }
      const key = root[1].toLowerCase();
      if (!seen[key]) {
        seen[key] = true;
        targets.push({ hub, title, url: root[1] });
      }
    };
    model.hubs.forEach((hub) => {
      if (hub.url) {
        add(hub.title, hub.title, hub.url);
      }
      hub.sites.forEach((site) => add(hub.title, site.title, site.url));
    });
    return targets;
  }

  public async scan(onProgress: (p: IScanProgress) => void, onSite: (site: ISiteMetrics) => void): Promise<IScorecardRun> {
    const targets = await this.getTargets();
    const sites: ISiteMetrics[] = [];
    const queue = targets.slice();
    let done = 0;
    onProgress({ done, total: targets.length, current: '' });

    const worker = async (): Promise<void> => {
      for (let target = queue.shift(); target; target = queue.shift()) {
        onProgress({ done, total: targets.length, current: target.title });
        const site = await this.scanSite(target);
        sites.push(site);
        done++;
        onSite(site);
        onProgress({ done, total: targets.length, current: target.title });
      }
    };
    const workers: Promise<void>[] = [];
    for (let i = 0; i < PARALLEL_SITES; i++) {
      workers.push(worker());
    }
    await Promise.all(workers);

    // Keep the navigation order (hub, then site) rather than completion order.
    const order = targets.map((t) => t.url);
    sites.sort((a, b) => order.indexOf(a.url) - order.indexOf(b.url));
    return {
      version: 1,
      runAt: new Date().toISOString(),
      requiredFields: this.options.requiredFields,
      staleYears: this.options.staleYears,
      sites
    };
  }

  public async scanSite(target: IScanTarget): Promise<ISiteMetrics> {
    const base = { hub: target.hub, title: target.title, url: target.url };
    try {
      const source = new LibrariesSource(this.httpClient, {
        webUrl: target.url,
        webTitle: target.title,
        libraryTitles: [],
        excludedFields: [],
        maxItems: this.options.maxDocsPerSite
      });
      const [data, storageBytes, owners] = await Promise.all([source.load(() => undefined), this.getStorage(target.url), this.getOwners(target.url)]);
      const failedLibraries = data.sources.filter((s) => s.error);
      if (data.sources.length && failedLibraries.length === data.sources.length) {
        throw new Error(failedLibraries[0].error);
      }
      const itemCount = data.sources.reduce((sum, s) => sum + (s.itemCount || 0), 0);
      const site = measureSite(
        base,
        data.documents,
        data.fields,
        itemCount,
        data.truncated,
        this.options.requiredFields,
        this.options.staleYears,
        new Date()
      );
      site.storageBytes = storageBytes;
      site.ownership = owners.kind;
      site.owners = owners.count;
      return site;
    } catch (err) {
      const status = (err as { status?: number }).status;
      const message = (err as Error).message || '';
      return {
        ...base,
        documents: 0,
        scanned: 0,
        sampled: false,
        fieldFilled: {},
        missingColumns: [],
        complete: 0,
        labelled: 0,
        stale: 0,
        staleBytes: 0,
        duplicates: 0,
        duplicateBytes: 0,
        ownership: 'unknown',
        error:
          status === 404 || /^404/.test(message)
            ? 'Site not found'
            : status === 401 || status === 403 || /^40[13]/.test(message)
            ? 'No access'
            : message
      };
    }
  }

  /** Site collection storage as SharePoint tracks it; undefined when this person may not read it. */
  private async getStorage(siteUrl: string): Promise<number | undefined> {
    try {
      const json = await this.rest.getJson<{ Usage?: { Storage?: number } }>(`${siteUrl}/_api/site?$select=Usage`);
      return typeof json.Usage?.Storage === 'number' ? json.Usage.Storage : undefined;
    } catch {
      return undefined;
    }
  }

  /**
   * Owners of the site's owners group. A group entry (the ATC-SP-{Dept}-Owners
   * Entra group, or a Microsoft 365 group's owners) is what the permission
   * model asks for; two or more named people are acceptable; one is a risk.
   */
  private async getOwners(siteUrl: string): Promise<{ kind: OwnershipKind; count?: number }> {
    try {
      const json = await this.rest.getJson<{ value: IOwnerEntry[] }>(`${siteUrl}/_api/web/AssociatedOwnerGroup/Users?$select=PrincipalType`);
      const entries = json.value || [];
      if (entries.some((e) => e.PrincipalType === SECURITY_GROUP)) {
        return { kind: 'group', count: entries.length };
      }
      return { kind: entries.length >= 2 ? 'people' : entries.length === 1 ? 'single' : 'none', count: entries.length };
    } catch {
      return { kind: 'unknown' };
    }
  }
}
