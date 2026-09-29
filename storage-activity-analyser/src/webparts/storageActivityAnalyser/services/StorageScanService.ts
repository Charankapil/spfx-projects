import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
  IAgeHistogram,
  ILargeFile,
  ILibraryResult,
  ILibraryStorageMetrics,
  IScanProgress,
  IScanResult,
  IWebResult,
  ScanScope
} from '../models/IScanResult';
import { ageInMonths, emptyHistogram, MIN_THRESHOLD_MONTHS } from './activity';
import { describeError } from './httpErrors';

/**
 * Every call goes through the ambient SPHttpClient, which reuses the signed-in
 * user's SharePoint session. No Azure AD app registration, API permission or
 * admin consent is involved, and the scan only sees what the user can open.
 */

interface IWebListItem {
  Title: string;
  ServerRelativeUrl: string;
  WebTemplate: string;
}

interface IListListItem {
  Id: string;
  Title: string;
  ItemCount: number;
  RootFolder: { ServerRelativeUrl: string };
}

interface IFileItem {
  FSObjType?: number | string;
  FileSystemObjectType?: number;
  Modified?: string;
  FileRef?: string;
  FileLeafRef?: string;
  File?: { Length?: number | string } | null;
}

interface IItemsPage {
  value?: IFileItem[];
  'odata.nextLink'?: string;
  '@odata.nextLink'?: string;
}

interface IStorageMetricsResponse {
  StorageMetrics?: { TotalSize?: number | string; TotalFileStreamSize?: number | string };
}

const REQUEST_GAP_MS = 150;
const MAX_THROTTLE_RETRIES = 5;
/** The REST items endpoint returns at most 5,000 items per page. */
const MAX_PAGE_SIZE = 5000;
/** A page that fails for a reason other than throttling is retried with fewer items, down to this size. */
const MIN_PAGE_SIZE = 500;
const LARGEST_FILES_KEPT = 200;

export class ScanCancelledError extends Error {}

class HttpError extends Error {
  constructor(public status: number, message: string) {
    super(message);
  }
}

function toNumber(value: number | string | undefined | null): number {
  if (typeof value === 'number') {
    return value;
  }
  const parsed = parseFloat(String(value === undefined || value === null ? '' : value));
  return isNaN(parsed) ? 0 : parsed;
}

function setPageSize(url: string, size: number): string {
  if (/([?&](\$|%24)top=)\d+/i.test(url)) {
    return url.replace(/([?&](\$|%24)top=)\d+/i, `$1${size}`);
  }
  return `${url}${url.indexOf('?') >= 0 ? '&' : '?'}$top=${size}`;
}

export class StorageScanService {
  private cancelled = false;

  constructor(private context: WebPartContext) {}

  public cancel(): void {
    this.cancelled = true;
  }

  private throwIfCancelled(): void {
    if (this.cancelled) {
      throw new ScanCancelledError('Scan cancelled.');
    }
  }

  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  /** GET with SharePoint throttling (429/503) handled by waiting for Retry-After. */
  private async getJson<T>(url: string, attempt = 0): Promise<T> {
    this.throwIfCancelled();
    let response: SPHttpClientResponse;
    try {
      response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    } catch (err) {
      // Network failure or timeout: retry a couple of times before giving up.
      if (attempt < 2) {
        await this.delay(2000 * (attempt + 1));
        return this.getJson<T>(url, attempt + 1);
      }
      throw err;
    }

    if (response.status === 429 || response.status === 503) {
      if (attempt >= MAX_THROTTLE_RETRIES) {
        throw new HttpError(response.status, 'SharePoint throttled this request too many times. Try again later.');
      }
      const header = response.headers.get('Retry-After');
      const seconds = header ? parseInt(header, 10) : NaN;
      await this.delay((isNaN(seconds) ? 5 * (attempt + 1) : seconds) * 1000);
      return this.getJson<T>(url, attempt + 1);
    }

    if (!response.ok) {
      const detail = await describeError(response);
      console.error(`[Storage Activity Analyser] ${response.status} from ${url}: ${detail}`);
      throw new HttpError(response.status, `${response.status}: ${detail}`);
    }

    await this.delay(REQUEST_GAP_MS);
    return (await response.json()) as T;
  }

  private toAbsoluteUrl(serverRelativeUrl: string): string {
    return `${new URL(this.context.pageContext.site.absoluteUrl).origin}${serverRelativeUrl}`;
  }

  /** Site collection storage as SharePoint reports it; readable by anyone who can read the site. */
  private async getSiteStorage(): Promise<number | undefined> {
    try {
      const json = await this.getJson<{ Usage?: { Storage?: number | string } }>(
        `${this.context.pageContext.site.absoluteUrl}/_api/site?$select=Usage`
      );
      return json.Usage && json.Usage.Storage !== undefined ? toNumber(json.Usage.Storage) : undefined;
    } catch (err) {
      if (err instanceof ScanCancelledError) {
        throw err;
      }
      return undefined;
    }
  }

  private async getWebTitle(webUrl: string): Promise<string> {
    if (webUrl === this.context.pageContext.web.absoluteUrl) {
      return this.context.pageContext.web.title;
    }
    try {
      const json = await this.getJson<{ Title?: string }>(`${webUrl}/_api/web?$select=Title`);
      return json.Title || webUrl;
    } catch (err) {
      if (err instanceof ScanCancelledError) {
        throw err;
      }
      return webUrl;
    }
  }

  /**
   * getsubwebsfilteredforcurrentuser only returns subsites the user can open,
   * so broken-inheritance subsites do not turn into a wall of 403s. App webs
   * live on a separate domain and are skipped.
   */
  private async getSubwebs(webUrl: string): Promise<IWebListItem[]> {
    const json = await this.getJson<{ value?: IWebListItem[] }>(
      `${webUrl}/_api/web/getsubwebsfilteredforcurrentuser(nwebtemplatefilter=-1,nconfigurationfilter=0)` +
        `?$select=Title,ServerRelativeUrl,WebTemplate`
    );
    return (json.value || []).filter((w) => w.WebTemplate !== 'APP');
  }

  /**
   * Every visible document library (BaseType 1): Documents, Site Assets,
   * Site Pages, Style Library, custom libraries and so on - they all count
   * towards the site's storage.
   */
  private async getLibraries(webUrl: string): Promise<IListListItem[]> {
    const json = await this.getJson<{ value?: IListListItem[] }>(
      `${webUrl}/_api/web/lists?$filter=BaseType eq 1 and Hidden eq false` +
        `&$select=Id,Title,ItemCount,RootFolder/ServerRelativeUrl&$expand=RootFolder`
    );
    return json.value || [];
  }

  /** Library storage including version history. Not every user can read it, so failure is not an error. */
  private async getStorageMetrics(webUrl: string, listId: string): Promise<ILibraryStorageMetrics | undefined> {
    try {
      const json = await this.getJson<IStorageMetricsResponse>(
        `${webUrl}/_api/web/lists(guid'${listId}')/RootFolder?$select=StorageMetrics&$expand=StorageMetrics`
      );
      if (!json.StorageMetrics) {
        return undefined;
      }
      return {
        totalSize: toNumber(json.StorageMetrics.TotalSize),
        fileStreamSize: toNumber(json.StorageMetrics.TotalFileStreamSize)
      };
    } catch (err) {
      if (err instanceof ScanCancelledError) {
        throw err;
      }
      return undefined;
    }
  }

  private async discoverWebs(
    webUrl: string,
    title: string,
    webs: IWebResult[],
    libraries: ILibraryResult[],
    progress: IScanProgress,
    onProgress: (p: IScanProgress) => void
  ): Promise<void> {
    this.throwIfCancelled();
    progress.currentItem = webUrl;
    progress.websFound++;
    onProgress({ ...progress });

    const web: IWebResult = { title, url: webUrl };
    webs.push(web);
    const errors: string[] = [];

    let lists: IListListItem[] = [];
    try {
      lists = await this.getLibraries(webUrl);
    } catch (err) {
      if (err instanceof ScanCancelledError) {
        throw err;
      }
      errors.push(`Libraries: ${err instanceof Error ? err.message : 'unknown error'}`);
    }
    for (const list of lists) {
      libraries.push({
        id: list.Id,
        title: list.Title,
        webTitle: title,
        webUrl,
        url: this.toAbsoluteUrl(list.RootFolder.ServerRelativeUrl),
        itemCount: list.ItemCount,
        files: 0,
        bytes: 0,
        histogram: emptyHistogram()
      });
      progress.itemsExpected += list.ItemCount;
    }
    progress.librariesFound = libraries.length;
    onProgress({ ...progress });

    let subwebs: IWebListItem[] = [];
    try {
      subwebs = await this.getSubwebs(webUrl);
    } catch (err) {
      if (err instanceof ScanCancelledError) {
        throw err;
      }
      errors.push(`Subsites: ${err instanceof Error ? err.message : 'unknown error'}`);
    }
    if (errors.length > 0) {
      web.error = errors.join(' | ');
    }

    for (const sub of subwebs) {
      await this.discoverWebs(this.toAbsoluteUrl(sub.ServerRelativeUrl), sub.Title, webs, libraries, progress, onProgress);
    }
  }

  /**
   * Reads one page of items, halving the page size on failures other than
   * throttling (very large pages can time out on busy libraries).
   */
  private async getItemsPage(url: string, pageSize: number): Promise<{ page: IItemsPage; pageSize: number }> {
    let size = pageSize;
    for (;;) {
      try {
        const page = await this.getJson<IItemsPage>(setPageSize(url, size));
        return { page, pageSize: size };
      } catch (err) {
        const retryable = err instanceof HttpError ? err.status >= 500 && err.status !== 503 : !(err instanceof ScanCancelledError);
        if (!retryable || size <= MIN_PAGE_SIZE) {
          throw err;
        }
        size = Math.max(MIN_PAGE_SIZE, Math.floor(size / 2));
      }
    }
  }

  /**
   * Pages through every item in the library, reading only the size and last
   * modified date of each file. Folders are skipped. Paging follows
   * SharePoint's own next link (ordered by ID), which works on libraries of
   * any size without hitting the list view threshold.
   */
  private async readLibrary(
    library: ILibraryResult,
    scanStart: Date,
    largest: ILargeFile[],
    progress: IScanProgress,
    onProgress: (p: IScanProgress) => void
  ): Promise<void> {
    const histogram: IAgeHistogram = library.histogram;
    let next: string | undefined =
      `${library.webUrl}/_api/web/lists(guid'${library.id}')/items` +
      `?$select=Id,FSObjType,Modified,FileRef,FileLeafRef,File/Length&$expand=File&$top=${MAX_PAGE_SIZE}`;
    let pageSize = MAX_PAGE_SIZE;

    while (next) {
      this.throwIfCancelled();
      const result: { page: IItemsPage; pageSize: number } = await this.getItemsPage(next, pageSize);
      pageSize = result.pageSize;
      const items = result.page.value || [];

      for (const item of items) {
        const type = item.FSObjType !== undefined ? Number(item.FSObjType) : item.FileSystemObjectType;
        // Folders have FSObjType 1; if the type is missing, a folder is the item without a File.
        if (type === 1 || ((type === undefined || isNaN(type)) && !item.File)) {
          continue;
        }
        const modified = item.Modified ? new Date(item.Modified) : undefined;
        const bytes = item.File ? toNumber(item.File.Length) : 0;
        const age = modified && !isNaN(modified.getTime()) ? ageInMonths(modified, scanStart) : 0;

        histogram.counts[age]++;
        histogram.bytes[age] += bytes;
        library.files++;
        library.bytes += bytes;
        progress.filesRead++;
        progress.bytesRead += bytes;

        if (age >= MIN_THRESHOLD_MONTHS && bytes > 0 && item.FileRef) {
          largest.push({
            name: item.FileLeafRef || item.FileRef.substring(item.FileRef.lastIndexOf('/') + 1),
            serverRelativeUrl: item.FileRef,
            bytes,
            modified: item.Modified || '',
            ageMonths: age,
            libraryTitle: library.title,
            webTitle: library.webTitle
          });
        }
      }

      if (largest.length > LARGEST_FILES_KEPT * 2) {
        largest.sort((a, b) => b.bytes - a.bytes);
        largest.length = LARGEST_FILES_KEPT;
      }

      onProgress({ ...progress });
      next = result.page['odata.nextLink'] || result.page['@odata.nextLink'];
    }
  }

  public async scan(
    scope: ScanScope,
    onProgress: (p: IScanProgress) => void,
    onLibraryDone?: (libraries: ILibraryResult[]) => void
  ): Promise<IScanResult> {
    this.cancelled = false;
    const scanStart = new Date();
    const rootUrl =
      scope === 'currentWeb' ? this.context.pageContext.web.absoluteUrl : this.context.pageContext.site.absoluteUrl;

    const progress: IScanProgress = {
      phase: 'starting',
      currentItem: rootUrl,
      websFound: 0,
      librariesFound: 0,
      librariesDone: 0,
      itemsExpected: 0,
      filesRead: 0,
      bytesRead: 0
    };
    onProgress({ ...progress });

    const storagePromise = this.getSiteStorage();
    const rootTitle = await this.getWebTitle(rootUrl);

    progress.phase = 'discovering';
    const webs: IWebResult[] = [];
    const libraries: ILibraryResult[] = [];
    await this.discoverWebs(rootUrl, rootTitle, webs, libraries, progress, onProgress);

    progress.phase = 'reading-files';
    const largest: ILargeFile[] = [];
    for (const library of libraries) {
      this.throwIfCancelled();
      progress.currentItem = library.url;
      onProgress({ ...progress });
      try {
        await this.readLibrary(library, scanStart, largest, progress, onProgress);
      } catch (err) {
        if (err instanceof ScanCancelledError) {
          throw err;
        }
        library.error = err instanceof Error ? err.message : 'unknown error';
        library.partial = library.files > 0;
      }
      library.metrics = await this.getStorageMetrics(library.webUrl, library.id);
      progress.librariesDone++;
      onProgress({ ...progress });
      if (onLibraryDone) {
        onLibraryDone(libraries.slice());
      }
    }

    largest.sort((a, b) => b.bytes - a.bytes);
    largest.length = Math.min(largest.length, LARGEST_FILES_KEPT);

    const siteStorageBytes = await storagePromise;
    progress.phase = 'completed';
    onProgress({ ...progress });

    return {
      scope,
      rootUrl,
      rootTitle,
      siteStorageBytes,
      webs,
      libraries,
      largestOldFiles: largest,
      scanStartedAt: scanStart.toISOString(),
      scanCompletedAt: new Date().toISOString(),
      scannedBy: this.context.pageContext.user.displayName
    };
  }
}
