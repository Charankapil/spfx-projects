import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import {
  IAgeHistogram,
  IFileTypeStat,
  ILargeFile,
  ILibraryResult,
  ILibraryStorageMetrics,
  IScanProgress,
  IScanResult,
  IWebResult,
  ScanScope
} from '../models/IScanResult';
import { ageInMonths, bandIndex, emptyFileTypeStat, emptyHistogram, extensionOf, MIN_THRESHOLD_MONTHS } from './activity';
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
  Id?: number;
  ID?: number;
  FSObjType?: number | string;
  FileSystemObjectType?: number;
  Modified?: string;
  FileRef?: string;
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
/**
 * Throttling is expected on long scans, so a throttled request waits and
 * tries again for a long time (about half an hour in total) rather than
 * giving up part-way through a big library.
 */
const MAX_THROTTLE_RETRIES = 12;
const MAX_BACKOFF_SECONDS = 300;
/**
 * nometadata drops the per-item odata.type / id / etag / editLink fields
 * (also on the expanded File), which makes each 5,000-item page much
 * smaller to download and parse.
 */
const JSON_OPTIONS = { headers: { Accept: 'application/json;odata.metadata=nometadata' } };
/** Libraries with more items than this are split into ID ranges read in parallel. */
const PARALLEL_ITEM_THRESHOLD = 20000;
const PARALLEL_READERS = 3;
const IDS_PER_RANGE = 50000;
/** The REST items endpoint returns at most 5,000 items per page. */
const MAX_PAGE_SIZE = 5000;
/** A page that fails for a reason other than throttling is retried with fewer items, down to this size. */
const MIN_PAGE_SIZE = 500;
const LARGEST_FILES_KEPT = 200;
/**
 * Distinct extensions tracked per library. Libraries full of odd names
 * (numbered backups like .001, .002 ...) would otherwise grow the saved
 * result without limit; anything past this is counted as "(other)".
 */
const MAX_FILE_TYPES_PER_LIBRARY = 300;

type FileTypeMap = Map<string, IFileTypeStat>;

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
      response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, JSON_OPTIONS);
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
      const backoff = Math.min(MAX_BACKOFF_SECONDS, 5 * Math.pow(2, attempt));
      await this.delay((isNaN(seconds) ? backoff : Math.min(seconds, MAX_BACKOFF_SECONDS)) * 1000);
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
   * Splits a big library into ID ranges so several readers can page through
   * it at once. ID is the list's primary key, so filtering on it is allowed
   * on libraries of any size. Small libraries, or any library whose highest
   * ID cannot be read, are read as one range.
   */
  private async planRanges(library: ILibraryResult): Promise<(string | undefined)[]> {
    if (library.itemCount <= PARALLEL_ITEM_THRESHOLD) {
      return [undefined];
    }
    let maxId = 0;
    try {
      const json = await this.getJson<IItemsPage>(
        `${library.webUrl}/_api/web/lists(guid'${library.id}')/items?$select=Id&$orderby=Id desc&$top=1`
      );
      const top = (json.value || [])[0];
      maxId = top ? Number(top.Id !== undefined ? top.Id : top.ID) || 0 : 0;
    } catch (err) {
      if (err instanceof ScanCancelledError) {
        throw err;
      }
    }
    if (maxId <= IDS_PER_RANGE) {
      return [undefined];
    }
    const ranges: string[] = [];
    for (let from = 0; from < maxId; from += IDS_PER_RANGE) {
      // The last range is open-ended so files added during the scan are not missed.
      ranges.push(from + IDS_PER_RANGE >= maxId ? `Id gt ${from}` : `Id gt ${from} and Id le ${from + IDS_PER_RANGE}`);
    }
    return ranges;
  }

  /**
   * Reads every item in the library, keeping only each file's size and last
   * modified date. Large libraries are read by several readers in parallel,
   * one ID range each. A range that fails is recorded and the others carry
   * on, so one bad page does not lose the rest of a million-file library.
   */
  private async readLibrary(
    library: ILibraryResult,
    scanStart: Date,
    largest: ILargeFile[],
    progress: IScanProgress,
    onProgress: (p: IScanProgress) => void
  ): Promise<void> {
    const ranges = await this.planRanges(library);
    const types: FileTypeMap = new Map<string, IFileTypeStat>();
    const errors: string[] = [];
    let itemsSeen = 0;
    let nextRange = 0;

    const reader = async (): Promise<void> => {
      while (nextRange < ranges.length) {
        const filter = ranges[nextRange++];
        try {
          // Not "itemsSeen += await ...": that reads itemsSeen before the await and loses other readers' counts.
          const seenInRange = await this.readRange(library, filter, scanStart, types, largest, progress, onProgress);
          itemsSeen += seenInRange;
        } catch (err) {
          if (err instanceof ScanCancelledError) {
            throw err;
          }
          const message = err instanceof Error ? err.message : 'unknown error';
          if (errors.indexOf(message) < 0) {
            errors.push(message);
          }
        }
      }
    };
    const readers: Promise<void>[] = [];
    for (let i = 0; i < Math.min(PARALLEL_READERS, ranges.length); i++) {
      readers.push(reader());
    }
    await Promise.all(readers);

    const fileTypes: IFileTypeStat[] = [];
    types.forEach((stat) => fileTypes.push(stat));
    fileTypes.sort((a, b) => b.bytes.reduce((x, y) => x + y, 0) - a.bytes.reduce((x, y) => x + y, 0));
    library.fileTypes = fileTypes;

    if (errors.length > 0) {
      library.error = errors.join(' | ');
      library.partial = library.files > 0;
    }
    // ItemCount counts every item, including ones the user cannot see. A
    // shortfall beyond normal churn during the scan is reported, not hidden.
    const missing = library.itemCount - itemsSeen;
    if (!library.error && missing > Math.max(10, library.itemCount * 0.001)) {
      library.unreadItems = missing;
    }
  }

  /** Pages through one ID range (or the whole library); returns how many items, files and folders, it saw. */
  private async readRange(
    library: ILibraryResult,
    filter: string | undefined,
    scanStart: Date,
    types: FileTypeMap,
    largest: ILargeFile[],
    progress: IScanProgress,
    onProgress: (p: IScanProgress) => void
  ): Promise<number> {
    const histogram: IAgeHistogram = library.histogram;
    const base =
      `${library.webUrl}/_api/web/lists(guid'${library.id}')/items` +
      `?$select=Id,FSObjType,Modified,FileRef,File/Length&$expand=File` +
      (filter ? `&$filter=${encodeURIComponent(filter)}` : '');
    let next: string | undefined = `${base}&$top=${MAX_PAGE_SIZE}`;
    let pageSize = MAX_PAGE_SIZE;
    let lastId = 0;
    let seen = 0;

    while (next) {
      this.throwIfCancelled();
      const result: { page: IItemsPage; pageSize: number } = await this.getItemsPage(next, pageSize);
      pageSize = result.pageSize;
      const items = result.page.value || [];
      seen += items.length;

      for (const item of items) {
        const id = Number(item.Id !== undefined ? item.Id : item.ID) || 0;
        if (id > lastId) {
          lastId = id;
        }
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

        const name = item.FileRef ? item.FileRef.substring(item.FileRef.lastIndexOf('/') + 1) : '';
        let ext = extensionOf(name);
        let typeStat = types.get(ext);
        if (!typeStat) {
          if (types.size >= MAX_FILE_TYPES_PER_LIBRARY) {
            ext = '(other)';
            typeStat = types.get(ext);
          }
          if (!typeStat) {
            typeStat = emptyFileTypeStat(ext);
            types.set(ext, typeStat);
          }
        }
        const band = bandIndex(age);
        typeStat.counts[band]++;
        typeStat.bytes[band] += bytes;

        if (age >= MIN_THRESHOLD_MONTHS && bytes > 0 && item.FileRef) {
          largest.push({
            name,
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
      if (!next && items.length >= pageSize && lastId > 0) {
        // A full page with no next link: continue from the last ID ourselves
        // rather than silently stopping short.
        next = `${base}&$skiptoken=${encodeURIComponent(`Paged=TRUE&p_ID=${lastId}`)}&$top=${pageSize}`;
      }
    }
    return seen;
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
