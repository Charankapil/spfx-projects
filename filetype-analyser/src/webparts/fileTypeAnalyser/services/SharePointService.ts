import {
  ODataVersion,
  SPHttpClient,
  SPHttpClientConfiguration,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { IFileTypeStat } from '../models/IFileTypeStat';
import { ILibraryNode } from '../models/ILibraryNode';
import { IScanProgress } from '../models/IScanProgress';
import { ISiteCollectionOverview } from '../models/ISiteCollectionOverview';
import { IStorageInfo } from '../models/IStorageInfo';
import { IWebNode } from '../models/IWebNode';

/**
 * All calls go through the ambient SPHttpClient, which reuses the current
 * user's SharePoint session (form digest / cookie auth). No AAD app
 * registration, client id or secret is required, and nothing ever leaves
 * SharePoint.
 */

interface IWebListItem {
  Title: string;
  Url: string;
  ServerRelativeUrl: string;
}

interface IListListItem {
  Id: string;
  Title: string;
  ItemCount: number;
  RootFolder: { ServerRelativeUrl: string };
}

interface ISearchRefinerEntry {
  RefinementName: string;
  RefinementCount: string;
}

/**
 * OData 3 ("verbose") wraps every collection in a { results: [] } envelope
 * while OData 4 / nometadata returns a bare array, and which one comes back
 * depends on the negotiated Accept header. Both shapes are accepted so the
 * parsing does not silently yield zero file types if the format changes.
 */
type ODataCollection<T> = T[] | { results?: T[] } | undefined;

interface ISearchRefiner {
  Name: string;
  Entries: ODataCollection<ISearchRefinerEntry>;
}

interface IPrimaryQueryResult {
  RelevantResults?: { TotalRows?: number };
  RefinementResults?: { Refiners?: ODataCollection<ISearchRefiner> };
}

interface ISearchQueryResponse {
  PrimaryQueryResult?: IPrimaryQueryResult;
  d?: {
    query?: { PrimaryQueryResult?: IPrimaryQueryResult };
    PrimaryQueryResult?: IPrimaryQueryResult;
  };
}

function toArray<T>(collection: ODataCollection<T>): T[] {
  if (!collection) {
    return [];
  }
  if (Array.isArray(collection)) {
    return collection;
  }
  return collection.results || [];
}

const REQUEST_GAP_MS = 120;
const MAX_RETRIES = 3;

/**
 * SPHttpClient.configurations.v1 sends "OData-Version: 4.0", but the search
 * REST endpoint only speaks OData 3.0 and answers a v4 request with HTTP 500
 * as soon as it has actual rows or refiners to serialize (an empty result set
 * happens to survive, which makes this look like a query problem when it is
 * really a protocol-version one). The _api/web and _api/site calls are fine
 * on v4, so only search is downgraded.
 */
const SEARCH_CONFIG: SPHttpClientConfiguration = SPHttpClient.configurations.v1.overrideWith({
  defaultODataVersion: ODataVersion.v3
});

export class ScanCancelledError extends Error {}

export class SharePointService {
  private cancelled = false;

  constructor(private context: WebPartContext) {}

  public cancel(): void {
    this.cancelled = true;
  }

  public resetCancellation(): void {
    this.cancelled = false;
  }

  private throwIfCancelled(): void {
    if (this.cancelled) {
      throw new ScanCancelledError('Scan cancelled by user.');
    }
  }

  private get siteAbsoluteUrl(): string {
    return this.context.pageContext.site.absoluteUrl;
  }

  private async delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  /**
   * SharePoint puts the useful part of a failure in the response body, so read
   * it rather than reporting a bare status code - "Request failed (500)" says
   * nothing, while the body names the actual managed property or syntax it
   * choked on.
   */
  private async describeError(response: SPHttpClientResponse): Promise<string> {
    try {
      const body = await response.text();
      if (!body) {
        return response.statusText || 'no error details returned';
      }
      try {
        const parsed = JSON.parse(body);
        const message = parsed?.error?.message;
        if (typeof message === 'string') {
          return message;
        }
        if (message && typeof message.value === 'string') {
          return message.value;
        }
      } catch {
        // Body was not JSON - fall through and surface the raw text instead.
      }
      return body.substring(0, 300);
    } catch {
      return response.statusText || 'no error details returned';
    }
  }

  private async getJson<T>(
    url: string,
    configuration: SPHttpClientConfiguration = SPHttpClient.configurations.v1,
    attempt = 0
  ): Promise<T> {
    this.throwIfCancelled();
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(url, configuration);

    if (response.status === 429 || response.status === 503) {
      if (attempt >= MAX_RETRIES) {
        throw new Error('SharePoint throttled this request too many times.');
      }
      const retryAfterHeader = response.headers.get('Retry-After');
      const retryAfterSec = retryAfterHeader ? parseInt(retryAfterHeader, 10) : 2 * (attempt + 1);
      await this.delay((isNaN(retryAfterSec) ? 2 : retryAfterSec) * 1000);
      return this.getJson<T>(url, configuration, attempt + 1);
    }

    if (!response.ok) {
      const detail = await this.describeError(response);
      // The URL is long enough to swamp the tree UI, so keep it to the console.
      console.error(`[File Type Analyser] ${response.status} from ${url}: ${detail}`);
      throw new Error(`${response.status}: ${detail}`);
    }

    await this.delay(REQUEST_GAP_MS);
    return (await response.json()) as T;
  }

  /**
   * Overall site collection storage usage, from SharePoint's own tracked
   * metrics - no Azure / tenant admin API involved.
   */
  public async getStorageInfo(): Promise<IStorageInfo> {
    try {
      const url = `${this.siteAbsoluteUrl}/_api/site?$select=Usage`;
      const json = await this.getJson<{ Usage?: { Storage?: number } }>(url);
      const storageBytes = json.Usage && typeof json.Usage.Storage === 'number' ? json.Usage.Storage : 0;
      return { usedBytes: storageBytes, available: true };
    } catch {
      return { usedBytes: 0, available: false };
    }
  }

  private async getSubwebs(webUrl: string): Promise<IWebListItem[]> {
    const url = `${webUrl}/_api/web/webs?$select=Title,Url,ServerRelativeUrl`;
    const json = await this.getJson<{ value: IWebListItem[] }>(url);
    return json.value || [];
  }

  private async getDocumentLibraries(webUrl: string): Promise<IListListItem[]> {
    const url =
      `${webUrl}/_api/web/lists` +
      `?$filter=BaseTemplate eq 101 and Hidden eq false` +
      `&$select=Id,Title,ItemCount,RootFolder/ServerRelativeUrl` +
      `&$expand=RootFolder`;
    const json = await this.getJson<{ value: IListListItem[] }>(url);
    return json.value || [];
  }

  /**
   * Aggregates a file-type breakdown for one library using the search
   * index's refiners, instead of enumerating every file. The count comes
   * back in a single request regardless of how many files the library holds.
   *
   * Scoped by Path with a trailing "/*" rather than ListId: a ListId-scoped
   * query (with or without braces around the GUID) reliably returned
   * HTTP 500 from this tenant's search endpoint, while this exact
   * Path-based shape (only the managed property differed - IsDocument here
   * vs. a bogus contentclass in an earlier version) is confirmed to return
   * HTTP 200. The trailing slash before "*" stops "Documents" from also
   * matching a library like "Documents2".
   */
  private async getFileTypeBreakdown(
    libraryAbsoluteUrl: string
  ): Promise<{ stats: IFileTypeStat[]; totalFiles: number }> {
    const kql = `IsDocument:1 Path:"${libraryAbsoluteUrl}/*"`;
    const queryText = encodeURIComponent(`'${kql}'`);
    const url =
      `${this.siteAbsoluteUrl}/_api/search/query` +
      `?querytext=${queryText}` +
      `&rowlimit=1` +
      `&refiners='FileType'` +
      `&trimduplicates=false` +
      `&clienttype='ContentSearchRegular'`;

    const json = await this.getJson<ISearchQueryResponse>(url, SEARCH_CONFIG);
    const primary =
      json.PrimaryQueryResult || json.d?.query?.PrimaryQueryResult || json.d?.PrimaryQueryResult;
    const totalRowsRaw = primary?.RelevantResults?.TotalRows;
    const totalFiles = typeof totalRowsRaw === 'number' ? totalRowsRaw : parseInt(String(totalRowsRaw), 10) || 0;

    const refiners = toArray(primary?.RefinementResults?.Refiners);
    const fileTypeRefiner = refiners.filter((r) => r.Name === 'FileType')[0];
    const entries = toArray(fileTypeRefiner?.Entries);

    const stats: IFileTypeStat[] = entries.map((entry) => {
      const rawCount = String(entry.RefinementCount ?? '0').replace(/[^\d]/g, '');
      return {
        extension: (entry.RefinementName || '(unknown)').toLowerCase(),
        count: parseInt(rawCount, 10) || 0
      };
    });

    stats.sort((a, b) => b.count - a.count);
    return { stats, totalFiles };
  }

  /**
   * Walks the web/library structure (cheap metadata calls, not file
   * enumeration) and attaches a file-type breakdown to every library via
   * the search index. Reports progress as it goes so the UI can render a
   * tree incrementally rather than waiting for the whole scan.
   */
  public async scanSiteCollection(
    onProgress: (progress: IScanProgress) => void,
    onNodeUpdated?: (rootWeb: IWebNode) => void
  ): Promise<ISiteCollectionOverview> {
    this.resetCancellation();
    const scanStartedAt = new Date();

    const progress: IScanProgress = {
      phase: 'starting',
      currentItem: this.siteAbsoluteUrl,
      websDiscovered: 0,
      librariesDiscovered: 0,
      librariesScanned: 0
    };
    onProgress({ ...progress });

    const storagePromise = this.getStorageInfo();

    progress.phase = 'discovering-structure';
    onProgress({ ...progress });

    const rootWeb = await this.buildWebNode(
      this.siteAbsoluteUrl,
      this.context.pageContext.web.serverRelativeUrl,
      this.context.pageContext.web.title,
      progress,
      onProgress
    );

    progress.phase = 'aggregating-file-types';
    onProgress({ ...progress });

    const allLibraries: ILibraryNode[] = [];
    this.collectLibraries(rootWeb, allLibraries);

    for (const library of allLibraries) {
      this.throwIfCancelled();
      progress.currentItem = library.absoluteUrl;
      onProgress({ ...progress });
      try {
        const { stats, totalFiles } = await this.getFileTypeBreakdown(library.absoluteUrl);
        library.fileTypes = stats;
        library.totalFiles = totalFiles;
        library.scanned = true;
      } catch (err) {
        if (err instanceof ScanCancelledError) {
          throw err;
        }
        library.error = err instanceof Error ? err.message : 'Unknown error';
        library.scanned = true;
      }
      progress.librariesScanned++;
      onProgress({ ...progress });
      if (onNodeUpdated) {
        onNodeUpdated(rootWeb);
      }
    }

    const totalFileTypeStats = this.aggregateFileTypes(allLibraries);
    const totalFiles = totalFileTypeStats.reduce((sum, s) => sum + s.count, 0);
    const storage = await storagePromise;

    progress.phase = 'completed';
    onProgress({ ...progress });

    return {
      siteUrl: this.siteAbsoluteUrl,
      siteTitle: this.context.pageContext.web.title,
      storage,
      rootWeb,
      totalFileTypeStats,
      totalFiles,
      totalLibraries: allLibraries.length,
      totalWebs: progress.websDiscovered,
      scanStartedAt,
      scanCompletedAt: new Date()
    };
  }

  private async buildWebNode(
    webAbsoluteUrl: string,
    serverRelativeUrl: string,
    title: string,
    progress: IScanProgress,
    onProgress: (progress: IScanProgress) => void
  ): Promise<IWebNode> {
    this.throwIfCancelled();
    progress.currentItem = webAbsoluteUrl;
    progress.websDiscovered++;
    onProgress({ ...progress });

    const [libraries, subwebs] = await Promise.all([
      this.getDocumentLibraries(webAbsoluteUrl),
      this.getSubwebs(webAbsoluteUrl)
    ]);

    const libraryNodes: ILibraryNode[] = libraries.map((lib) => ({
      id: lib.Id,
      title: lib.Title,
      serverRelativeUrl: lib.RootFolder.ServerRelativeUrl,
      absoluteUrl: this.toAbsoluteUrl(lib.RootFolder.ServerRelativeUrl),
      itemCount: lib.ItemCount,
      totalFiles: 0,
      fileTypes: [],
      scanned: false
    }));
    progress.librariesDiscovered += libraryNodes.length;
    onProgress({ ...progress });

    const webNode: IWebNode = {
      id: serverRelativeUrl,
      title,
      url: webAbsoluteUrl,
      serverRelativeUrl,
      libraries: libraryNodes,
      webs: []
    };

    for (const sub of subwebs) {
      this.throwIfCancelled();
      const childNode = await this.buildWebNode(
        sub.Url,
        sub.ServerRelativeUrl,
        sub.Title,
        progress,
        onProgress
      );
      webNode.webs.push(childNode);
    }

    return webNode;
  }

  private toAbsoluteUrl(serverRelativeUrl: string): string {
    const origin = new URL(this.siteAbsoluteUrl).origin;
    return `${origin}${serverRelativeUrl}`;
  }

  private collectLibraries(web: IWebNode, out: ILibraryNode[]): void {
    out.push(...web.libraries);
    for (const child of web.webs) {
      this.collectLibraries(child, out);
    }
  }

  private aggregateFileTypes(libraries: ILibraryNode[]): IFileTypeStat[] {
    const map = new Map<string, number>();
    for (const lib of libraries) {
      for (const stat of lib.fileTypes) {
        map.set(stat.extension, (map.get(stat.extension) || 0) + stat.count);
      }
    }
    const stats: IFileTypeStat[] = [];
    map.forEach((count, extension) => stats.push({ extension, count }));
    stats.sort((a, b) => b.count - a.count);
    return stats;
  }
}
