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
import { encodePath } from './safeData';

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
  Hidden?: boolean;
  IsCatalog?: boolean;
  BaseTemplate?: number;
  /** Last time a person changed anything in the list (not system updates). */
  LastItemUserModifiedDate?: string;
  RootFolder: { ServerRelativeUrl: string };
}

/** Which libraries a scan reads, and how; set by the site owner in the web part properties. */
export interface IScanOptions {
  scope: ScanScope;
  /** Also read hidden libraries, such as the Preservation Hold Library. Catalogs (_catalogs/*) are always skipped. */
  includeHidden: boolean;
  /** Skip Site Pages, Style Library, Form Templates and Site Assets. */
  excludeSystemLibraries: boolean;
  /** Library titles or URL names to skip, compared case-insensitively. */
  excludedLibraries: string[];
  /**
   * Quick scan: libraries nobody has changed for at least this many months are
   * measured as a whole from SharePoint's storage metrics (one request)
   * instead of file by file. Undefined for a detailed scan.
   */
  quickAfterMonths?: number;
}

/** URL names of the libraries SharePoint creates for the site itself rather than for people's files. */
const SYSTEM_LIBRARY_URL_NAMES = ['sitepages', 'style library', 'formservertemplates', 'siteassets'];
const SITE_PAGES_TEMPLATE = 119;

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
}

interface IStorageMetricsResponse {
  StorageMetrics?: {
    TotalSize?: number | string;
    TotalFileStreamSize?: number | string;
    TotalFileCount?: number | string;
    LastModified?: string;
  };
}

const REQUEST_GAP_MS = 100;
/**
 * Throttling is expected on long scans, so a throttled request waits and
 * tries again for a long time (about half an hour in total) rather than
 * giving up part-way through a big library.
 */
const MAX_THROTTLE_RETRIES = 12;
/** Network failures (no HTTP response at all) are retried separately from throttling. */
const MAX_NETWORK_RETRIES = 3;
const MAX_BACKOFF_SECONDS = 300;
/**
 * nometadata drops the per-item odata.type / id / etag / editLink fields
 * (also on the expanded File), which makes each 5,000-item page much
 * smaller to download and parse.
 */
const JSON_OPTIONS = { headers: { Accept: 'application/json;odata.metadata=nometadata' } };
/**
 * Item queries come in three forms, tried in this order when SharePoint
 * answers 406 "Not Acceptable" (it rejects a request's format, and which
 * part it dislikes can differ between libraries and tenants):
 *   0  fast:            $filter Id + $orderby=Id, lean Accept header
 *   1  no ordering:     $filter Id only, lean Accept header
 *   2  default headers: $filter Id only, SharePoint's default Accept header
 * Levels 1 and 2 rely on SharePoint returning items in ID order when no
 * $orderby is given (as it does), and the items read are checked against the
 * library's item count afterwards.
 */
const MAX_PAGING_LEVEL = 2;
/** After this many libraries needed a fallback, later libraries start at that level. */
const LEVEL_LEARN_AFTER = 3;
/** Libraries read at the same time. Most of a scan of many small libraries is per-library overhead. */
const LIBRARY_CONCURRENCY = 4;
/** Readers per large library (one ID range each). */
const PARALLEL_READERS = 3;
/** Requests in flight at once across the whole scan, to stay friendly to SharePoint's throttling. */
const MAX_CONCURRENT_REQUESTS = 6;
/** Libraries with more items than this are split into ID ranges read in parallel. */
const PARALLEL_ITEM_THRESHOLD = 20000;
const IDS_PER_RANGE = 50000;
/** The REST items endpoint returns at most 5,000 items per page. */
const MAX_PAGE_SIZE = 5000;
/** A page that fails for a reason other than throttling is retried with fewer items, down to this size. */
const MIN_PAGE_SIZE = 500;
/** After a batch still fails at the smallest size it is skipped; a range gives up after this many skips. */
const MAX_SKIPPED_BATCHES = 20;
const LARGEST_FILES_KEPT = 200;
/**
 * Distinct extensions tracked per library. Libraries full of odd names
 * (numbered backups like .001, .002 ...) would otherwise grow the saved
 * result without limit; anything past this is counted as "(other)".
 */
const MAX_FILE_TYPES_PER_LIBRARY = 300;

type FileTypeMap = Map<string, IFileTypeStat>;

/*
 * The SPFx build compiles to ES5, where a subclass of Error loses its
 * prototype: `instanceof ScanCancelledError` would always be false and a
 * cancel (or an HTTP status) would be mistaken for some other failure. Both
 * classes therefore restore their prototype explicitly.
 */
function restorePrototype(instance: object, prototype: object): void {
  // Object.setPrototypeOf exists in every browser SharePoint supports; the ES5 typings just do not declare it.
  (Object as unknown as { setPrototypeOf(o: object, p: object): object }).setPrototypeOf(instance, prototype);
}

export class ScanCancelledError extends Error {
  constructor(message: string) {
    super(message);
    restorePrototype(this, ScanCancelledError.prototype);
    this.name = 'ScanCancelledError';
  }
}

export class HttpError extends Error {
  public status: number;

  constructor(status: number, message: string) {
    super(message);
    restorePrototype(this, HttpError.prototype);
    this.name = 'HttpError';
    this.status = status;
  }
}

/**
 * Says why a request failed in terms a site owner can act on. A site
 * collection admin can still see these: throttling, timeouts and SharePoint
 * refusing a query are not permission problems.
 */
export function describeFailure(err: unknown): string {
  const message = err instanceof Error ? err.message : String(err);
  if (err instanceof HttpError) {
    if (err.status === 429 || err.status === 503) {
      return `SharePoint kept throttling the requests (${message})`;
    }
    if (err.status === 401 || err.status === 403) {
      return `Access denied (${message})`;
    }
    if (err.status === 406) {
      return `SharePoint rejected the request format on every form tried (${message})`;
    }
    if (err.status === 404) {
      return `Not found - it may have been deleted during the scan (${message})`;
    }
    if (/threshold/i.test(message)) {
      return `SharePoint refused the query because of the list view threshold (${message})`;
    }
    if (err.status >= 500) {
      return `SharePoint returned a server error or timed out (${message})`;
    }
    return message;
  }
  if (/fetch|network/i.test(message)) {
    return `The network connection dropped (${message})`;
  }
  return message;
}

function toNumber(value: number | string | undefined | null): number {
  if (typeof value === 'number') {
    return value;
  }
  const parsed = parseFloat(String(value === undefined || value === null ? '' : value));
  return isNaN(parsed) ? 0 : parsed;
}

/** One ID window of a library: items with from < Id <= to (to undefined = open-ended). */
interface IIdRange {
  from: number;
  to?: number;
}

export class StorageScanService {
  private cancelled = false;
  /** Pending waits, so Cancel can end a long throttling back-off at once instead of after it. */
  private waiters: (() => void)[] = [];
  private inFlight = 0;
  /** Item query level new libraries start at, raised once several libraries needed a fallback. */
  private pagingFloor = 0;
  private fallbackLibraries = 0;
  private highestLevelUsed = 0;
  private slotQueue: { resolve: () => void; reject: (e: Error) => void }[] = [];

  constructor(private context: WebPartContext) {}

  public cancel(): void {
    this.cancelled = true;
    const waiters = this.waiters;
    this.waiters = [];
    waiters.forEach((wake) => wake());
    const queued = this.slotQueue;
    this.slotQueue = [];
    queued.forEach((q) => q.reject(new ScanCancelledError('Scan cancelled.')));
  }

  private reset(): void {
    this.cancelled = false;
    this.waiters = [];
    this.slotQueue = [];
    this.inFlight = 0;
    this.pagingFloor = 0;
    this.fallbackLibraries = 0;
    this.highestLevelUsed = 0;
  }

  private throwIfCancelled(): void {
    if (this.cancelled) {
      throw new ScanCancelledError('Scan cancelled.');
    }
  }

  /** Waits `ms`, or rejects with ScanCancelledError as soon as the scan is cancelled. */
  private delay(ms: number): Promise<void> {
    this.throwIfCancelled();
    return new Promise<void>((resolve, reject) => {
      const wake = (): void => {
        clearTimeout(timer);
        reject(new ScanCancelledError('Scan cancelled.'));
      };
      const timer = setTimeout(() => {
        this.waiters = this.waiters.filter((w) => w !== wake);
        resolve();
      }, ms);
      this.waiters.push(wake);
    });
  }

  /** Waits for one of the MAX_CONCURRENT_REQUESTS slots. */
  private acquireSlot(): Promise<void> {
    this.throwIfCancelled();
    if (this.inFlight < MAX_CONCURRENT_REQUESTS) {
      this.inFlight++;
      return Promise.resolve();
    }
    return new Promise<void>((resolve, reject) => {
      this.slotQueue.push({
        resolve: () => {
          this.inFlight++;
          resolve();
        },
        reject
      });
    });
  }

  private releaseSlot(): void {
    this.inFlight--;
    const next = this.slotQueue.shift();
    if (next) {
      next.resolve();
    }
  }

  /** GET with SharePoint throttling (429/503) handled by waiting for Retry-After. The slot is not held while waiting. */
  private async getJson<T>(url: string, throttleAttempt = 0, networkAttempt = 0, lean = true): Promise<T> {
    await this.acquireSlot();
    let response: SPHttpClientResponse;
    try {
      response = await this.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        lean ? JSON_OPTIONS : undefined
      );
    } catch (err) {
      this.releaseSlot();
      this.throwIfCancelled();
      // Network failure or timeout (no response at all): retry with its own counter.
      if (networkAttempt < MAX_NETWORK_RETRIES) {
        await this.delay(2000 * (networkAttempt + 1));
        return this.getJson<T>(url, throttleAttempt, networkAttempt + 1, lean);
      }
      throw err;
    }

    let body: T | undefined;
    let failure: HttpError | undefined;
    try {
      if (response.ok) {
        body = (await response.json()) as T;
      } else if (response.status !== 429 && response.status !== 503) {
        const detail = await describeError(response);
        console.error(`[Storage Pulse] ${response.status} from ${url}: ${detail}`);
        failure = new HttpError(response.status, `${response.status}: ${detail}`);
      }
    } finally {
      this.releaseSlot();
    }
    this.throwIfCancelled();

    if (response.status === 429 || response.status === 503) {
      if (throttleAttempt >= MAX_THROTTLE_RETRIES) {
        throw new HttpError(response.status, `${response.status}: throttled ${MAX_THROTTLE_RETRIES} times in a row`);
      }
      const header = response.headers.get('Retry-After');
      const seconds = header ? parseInt(header, 10) : NaN;
      const backoff = Math.min(MAX_BACKOFF_SECONDS, 5 * Math.pow(2, throttleAttempt));
      await this.delay((isNaN(seconds) || seconds < 0 ? backoff : Math.min(seconds, MAX_BACKOFF_SECONDS)) * 1000);
      return this.getJson<T>(url, throttleAttempt + 1, networkAttempt, lean);
    }
    if (failure) {
      throw failure;
    }
    await this.delay(REQUEST_GAP_MS);
    return body as T;
  }

  private get origin(): string {
    return new URL(this.context.pageContext.site.absoluteUrl).origin;
  }

  private toAbsoluteUrl(serverRelativeUrl: string): string {
    // ServerRelativeUrl comes back unencoded; encode it so "#" or "%" in a name cannot break the URL.
    return `${this.origin}${encodePath(serverRelativeUrl)}`;
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
   * Every document library (BaseType 1) the options allow: Documents, Site
   * Assets, Site Pages, custom libraries and so on - they all count towards
   * the site's storage.
   */
  private async getLibraries(webUrl: string, options: IScanOptions): Promise<IListListItem[]> {
    const json = await this.getJson<{ value?: IListListItem[] }>(
      `${webUrl}/_api/web/lists?$filter=BaseType eq 1` +
        `&$select=Id,Title,ItemCount,Hidden,IsCatalog,BaseTemplate,LastItemUserModifiedDate,RootFolder/ServerRelativeUrl` +
        `&$expand=RootFolder`
    );
    const excluded = options.excludedLibraries.map((n) => n.trim().toLowerCase()).filter((n) => n.length > 0);
    return (json.value || []).filter((list) => {
      const urlName = decodeURIComponent(
        list.RootFolder.ServerRelativeUrl.substring(list.RootFolder.ServerRelativeUrl.lastIndexOf('/') + 1)
      ).toLowerCase();
      if (list.IsCatalog || urlName.charAt(0) === '_') {
        return false;
      }
      if (list.Hidden && !options.includeHidden) {
        return false;
      }
      if (
        options.excludeSystemLibraries &&
        (list.BaseTemplate === SITE_PAGES_TEMPLATE || SYSTEM_LIBRARY_URL_NAMES.indexOf(urlName) >= 0)
      ) {
        return false;
      }
      return excluded.indexOf(list.Title.toLowerCase()) < 0 && excluded.indexOf(urlName) < 0;
    });
  }

  /**
   * SharePoint's own storage figures for the whole library: size with and
   * without version history, file count and the most recent change anywhere
   * in it. Not every user can read them, so failure is not an error.
   */
  private async getStorageMetrics(webUrl: string, listId: string): Promise<ILibraryStorageMetrics | undefined> {
    try {
      const json = await this.getJson<IStorageMetricsResponse>(
        `${webUrl}/_api/web/lists(guid'${listId}')/RootFolder?$select=StorageMetrics&$expand=StorageMetrics`
      );
      const m = json.StorageMetrics;
      if (!m) {
        return undefined;
      }
      return {
        totalSize: toNumber(m.TotalSize),
        fileStreamSize: toNumber(m.TotalFileStreamSize),
        fileCount: m.TotalFileCount === undefined ? undefined : toNumber(m.TotalFileCount),
        lastModified: m.LastModified
      };
    } catch (err) {
      if (err instanceof ScanCancelledError) {
        throw err;
      }
      return undefined;
    }
  }

  private newLibrary(list: IListListItem, webUrl: string, webTitle: string): ILibraryResult & { lastUserChange?: string } {
    return {
      id: list.Id,
      title: list.Title,
      webTitle,
      webUrl,
      url: this.toAbsoluteUrl(list.RootFolder.ServerRelativeUrl),
      itemCount: list.ItemCount,
      files: 0,
      bytes: 0,
      histogram: emptyHistogram(),
      lastUserChange: list.LastItemUserModifiedDate
    };
  }

  private async discoverWebs(
    webUrl: string,
    title: string,
    webs: IWebResult[],
    libraries: ILibraryResult[],
    options: IScanOptions,
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
      lists = await this.getLibraries(webUrl, options);
    } catch (err) {
      if (err instanceof ScanCancelledError) {
        throw err;
      }
      errors.push(`Libraries: ${describeFailure(err)}`);
    }
    for (const list of lists) {
      libraries.push(this.newLibrary(list, webUrl, title));
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
      errors.push(`Subsites: ${describeFailure(err)}`);
    }
    if (errors.length > 0) {
      web.error = errors.join(' | ');
    }

    for (const sub of subwebs) {
      await this.discoverWebs(this.toAbsoluteUrl(sub.ServerRelativeUrl), sub.Title, webs, libraries, options, progress, onProgress);
    }
  }

  /** The library's highest item ID, or undefined if it cannot be read. */
  private async getMaxId(library: ILibraryResult): Promise<number | undefined> {
    const url = `${library.webUrl}/_api/web/lists(guid'${library.id}')/items?$select=Id&$orderby=Id desc&$top=1`;
    // Lean header first, then SharePoint's default one; without a highest ID the library is simply read in one range.
    for (const lean of [true, false]) {
      try {
        const json = await this.getJson<IItemsPage>(url, 0, 0, lean);
        const top = (json.value || [])[0];
        return top ? Number(top.Id !== undefined ? top.Id : top.ID) || 0 : 0;
      } catch (err) {
        if (err instanceof ScanCancelledError) {
          throw err;
        }
        if (!(err instanceof HttpError) || err.status !== 406) {
          return undefined;
        }
      }
    }
    return undefined;
  }

  /**
   * Splits a library into ID ranges so several readers can page through it
   * at once. ID is the list's primary key, so filtering on it is allowed on
   * libraries of any size. The last range is open-ended so files added during
   * the scan are not missed.
   */
  private planRanges(library: ILibraryResult, maxId: number | undefined): IIdRange[] {
    if (library.itemCount <= PARALLEL_ITEM_THRESHOLD || !maxId || maxId <= IDS_PER_RANGE) {
      return [{ from: 0 }];
    }
    const ranges: IIdRange[] = [];
    for (let from = 0; from < maxId; from += IDS_PER_RANGE) {
      ranges.push(from + IDS_PER_RANGE >= maxId ? { from } : { from, to: from + IDS_PER_RANGE });
    }
    return ranges;
  }

  /**
   * Quick scan shortcut: a library nobody has changed for at least the quick
   * threshold is measured from its storage metrics. SharePoint updates the
   * list's LastItemUserModifiedDate (and the metrics' LastModified) on any
   * change by a person, so every file in it is at least that old.
   */
  private measureAsWhole(
    library: ILibraryResult & { lastUserChange?: string },
    metrics: ILibraryStorageMetrics | undefined,
    scanStart: Date,
    quickAfterMonths: number | undefined
  ): boolean {
    if (quickAfterMonths === undefined || !metrics || metrics.fileCount === undefined || !library.lastUserChange) {
      return false;
    }
    const dates = [library.lastUserChange, metrics.lastModified]
      .map((d) => (d ? new Date(d) : undefined))
      .filter((d): d is Date => !!d && !isNaN(d.getTime()));
    if (dates.length === 0) {
      return false;
    }
    const lastChange = new Date(Math.max(...dates.map((d) => d.getTime())));
    const dormantMonths = ageInMonths(lastChange, scanStart);
    if (dormantMonths < quickAfterMonths) {
      return false;
    }
    library.files = metrics.fileCount;
    library.bytes = metrics.fileStreamSize;
    library.histogram.counts[dormantMonths] += metrics.fileCount;
    library.histogram.bytes[dormantMonths] += metrics.fileStreamSize;
    library.measuredAsWhole = { lastChange: lastChange.toISOString(), dormantMonths };
    return true;
  }

  /**
   * Reads every item in the library, keeping only each file's size and last
   * modified date. Large libraries are read by several readers in parallel,
   * one ID range each. A batch that keeps failing is skipped and recorded;
   * the rest of the library is still read.
   */
  private async readLibrary(
    library: ILibraryResult,
    scanStart: Date,
    largest: ILargeFile[],
    progress: IScanProgress,
    onProgress: (p: IScanProgress) => void
  ): Promise<void> {
    // The highest ID is only needed to split big libraries, or to skip past a bad batch.
    let maxId: Promise<number | undefined> | undefined =
      library.itemCount > PARALLEL_ITEM_THRESHOLD ? this.getMaxId(library) : undefined;
    const lastId = (): Promise<number | undefined> => {
      if (!maxId) {
        maxId = this.getMaxId(library);
      }
      return maxId;
    };
    const ranges = this.planRanges(library, maxId ? await maxId : undefined);
    const types: FileTypeMap = new Map<string, IFileTypeStat>();
    const errors: string[] = [];
    // Shared by all readers of this library: once one range learns a form SharePoint accepts, the others use it too.
    const paging = { level: this.pagingFloor };
    let itemsSeen = 0;
    let skippedBatches = 0;
    let nextRange = 0;

    const reader = async (): Promise<void> => {
      while (nextRange < ranges.length) {
        const range = ranges[nextRange++];
        // Not "itemsSeen += await ...": that reads itemsSeen before the await and loses other readers' counts.
        const outcome = await this.readRange(library, range, lastId, paging, scanStart, types, largest, progress, onProgress);
        itemsSeen += outcome.seen;
        skippedBatches += outcome.skipped;
        outcome.errors.forEach((e) => {
          if (errors.indexOf(e) < 0) {
            errors.push(e);
          }
        });
      }
    };
    const readers: Promise<void>[] = [];
    for (let i = 0; i < Math.min(PARALLEL_READERS, ranges.length); i++) {
      readers.push(reader());
    }
    await Promise.all(readers);

    if (paging.level > this.pagingFloor) {
      this.fallbackLibraries++;
      this.highestLevelUsed = Math.max(this.highestLevelUsed, paging.level);
      console.info(`[Storage Pulse] SharePoint rejected the fast item query for "${library.title}" (406); read it with compatibility level ${paging.level}.`);
      if (this.fallbackLibraries >= LEVEL_LEARN_AFTER) {
        this.pagingFloor = Math.max(this.pagingFloor, this.highestLevelUsed);
      }
    }

    const fileTypes: IFileTypeStat[] = [];
    types.forEach((stat) => fileTypes.push(stat));
    fileTypes.sort((a, b) => b.bytes.reduce((x, y) => x + y, 0) - a.bytes.reduce((x, y) => x + y, 0));
    library.fileTypes = fileTypes;

    if (errors.length > 0) {
      library.error =
        (skippedBatches > 0 ? `${skippedBatches} batch(es) of up to ${MIN_PAGE_SIZE} items could not be read. ` : '') +
        errors.join(' | ');
      library.partial = library.files > 0;
    }
    // ItemCount counts every item, including ones the user cannot see. A
    // shortfall beyond normal churn during the scan is reported, not hidden.
    const missing = library.itemCount - itemsSeen;
    if (!library.error && missing > Math.max(10, library.itemCount * 0.001)) {
      library.unreadItems = missing;
    }
  }

  /**
   * Pages through one ID range in ID order, using the last ID of each page as
   * the cursor for the next (no reliance on next links). A page that fails is
   * retried with fewer items; if even the smallest batch fails, that batch of
   * IDs is skipped and recorded so one bad item cannot stop the library.
   */
  private async readRange(
    library: ILibraryResult,
    range: IIdRange,
    lastId: () => Promise<number | undefined>,
    paging: { level: number },
    scanStart: Date,
    types: FileTypeMap,
    largest: ILargeFile[],
    progress: IScanProgress,
    onProgress: (p: IScanProgress) => void
  ): Promise<{ seen: number; skipped: number; errors: string[] }> {
    const histogram: IAgeHistogram = library.histogram;
    const itemsUrl = `${library.webUrl}/_api/web/lists(guid'${library.id}')/items`;
    const errors: string[] = [];
    let cursor = range.from;
    let pageSize = MAX_PAGE_SIZE;
    let seen = 0;
    let skipped = 0;
    let successes = 0;

    for (;;) {
      this.throwIfCancelled();
      const filter = `Id gt ${cursor}${range.to !== undefined ? ` and Id le ${range.to}` : ''}`;
      const level = paging.level;
      const url =
        `${itemsUrl}?$select=Id,FSObjType,Modified,FileRef,File/Length&$expand=File` +
        `&$filter=${encodeURIComponent(filter)}${level === 0 ? '&$orderby=Id' : ''}&$top=${pageSize}`;
      let items: IFileItem[];
      try {
        items = (await this.getJson<IItemsPage>(url, 0, 0, level < MAX_PAGING_LEVEL)).value || [];
      } catch (err) {
        if (err instanceof ScanCancelledError) {
          throw err;
        }
        // 406 rejects the request's format, so a smaller page would not help: try the next, more
        // conservative form of the query from the same place. Another reader may already have.
        if (err instanceof HttpError && err.status === 406 && level < MAX_PAGING_LEVEL) {
          paging.level = Math.max(paging.level, level + 1);
          continue;
        }
        // Server errors and timeouts may pass with a smaller page; a 4xx will not.
        const retryable = !(err instanceof HttpError) || (err.status >= 500 && err.status !== 503);
        successes = 0;
        if (retryable && pageSize > MIN_PAGE_SIZE) {
          pageSize = Math.max(MIN_PAGE_SIZE, Math.floor(pageSize / 2));
          continue;
        }
        const reason = describeFailure(err);
        if (errors.indexOf(reason) < 0) {
          errors.push(reason);
        }
        // Skip this batch of IDs and carry on, unless it keeps happening or there is nothing left.
        skipped++;
        const end = range.to !== undefined ? range.to : await lastId();
        if (!retryable || skipped >= MAX_SKIPPED_BATCHES || end === undefined || cursor + MIN_PAGE_SIZE >= end) {
          break;
        }
        cursor += MIN_PAGE_SIZE;
        continue;
      }

      seen += items.length;
      progress.itemsRead += items.length;
      const fullPage = items.length >= pageSize;
      // One slow page should not slow the rest of the library: after a few
      // successes in a row, try bigger pages again.
      if (++successes >= 3 && pageSize < MAX_PAGE_SIZE) {
        pageSize = Math.min(MAX_PAGE_SIZE, pageSize * 2);
        successes = 0;
      }
      for (const item of items) {
        const id = Number(item.Id !== undefined ? item.Id : item.ID) || 0;
        if (id > cursor) {
          cursor = id;
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

      if (!fullPage) {
        break;
      }
    }
    return { seen, skipped, errors };
  }

  /** Reads a set of libraries, LIBRARY_CONCURRENCY at a time. */
  private async readLibraries(
    libraries: ILibraryResult[],
    options: IScanOptions,
    scanStart: Date,
    largest: ILargeFile[],
    progress: IScanProgress,
    onProgress: (p: IScanProgress) => void
  ): Promise<void> {
    let next = 0;
    const worker = async (): Promise<void> => {
      while (next < libraries.length) {
        const library = libraries[next++] as ILibraryResult & { lastUserChange?: string };
        this.throwIfCancelled();
        progress.currentItem = library.url;
        onProgress({ ...progress });
        library.metrics = await this.getStorageMetrics(library.webUrl, library.id);
        if (this.measureAsWhole(library, library.metrics, scanStart, options.quickAfterMonths)) {
          progress.itemsRead += library.itemCount;
          progress.filesRead += library.files;
          progress.bytesRead += library.bytes;
        } else {
          try {
            await this.readLibrary(library, scanStart, largest, progress, onProgress);
          } catch (err) {
            if (err instanceof ScanCancelledError) {
              throw err;
            }
            library.error = describeFailure(err);
            library.partial = library.files > 0;
          }
        }
        delete library.lastUserChange;
        progress.librariesDone++;
        onProgress({ ...progress });
      }
    };
    const workers: Promise<void>[] = [];
    for (let i = 0; i < Math.min(LIBRARY_CONCURRENCY, libraries.length); i++) {
      workers.push(worker());
    }
    await Promise.all(workers);
  }

  private newProgress(currentItem: string): IScanProgress {
    return {
      phase: 'starting',
      currentItem,
      websFound: 0,
      librariesFound: 0,
      librariesDone: 0,
      itemsExpected: 0,
      itemsRead: 0,
      filesRead: 0,
      bytesRead: 0
    };
  }

  private finishLargest(largest: ILargeFile[]): ILargeFile[] {
    largest.sort((a, b) => b.bytes - a.bytes);
    largest.length = Math.min(largest.length, LARGEST_FILES_KEPT);
    return largest;
  }

  public async scan(options: IScanOptions, onProgress: (p: IScanProgress) => void): Promise<IScanResult> {
    this.reset();
    const scope = options.scope;
    const scanStart = new Date();
    const rootUrl =
      scope === 'currentWeb' ? this.context.pageContext.web.absoluteUrl : this.context.pageContext.site.absoluteUrl;

    const progress = this.newProgress(rootUrl);
    onProgress({ ...progress });

    const storagePromise = this.getSiteStorage();
    const rootTitle = await this.getWebTitle(rootUrl);

    progress.phase = 'discovering';
    const webs: IWebResult[] = [];
    const libraries: ILibraryResult[] = [];
    await this.discoverWebs(rootUrl, rootTitle, webs, libraries, options, progress, onProgress);

    progress.phase = 'reading-files';
    const largest: ILargeFile[] = [];
    await this.readLibraries(libraries, options, scanStart, largest, progress, onProgress);

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
      largestOldFiles: this.finishLargest(largest),
      scanStartedAt: scanStart.toISOString(),
      scanCompletedAt: new Date().toISOString(),
      scannedBy: this.context.pageContext.user.displayName,
      options: {
        includeHidden: options.includeHidden,
        excludeSystemLibraries: options.excludeSystemLibraries,
        excludedLibraries: options.excludedLibraries
      },
      quickAfterMonths: options.quickAfterMonths,
      paging: this.pagingSummary()
    };
  }

  /** How many libraries needed a more conservative query form, for the dashboard note. */
  private pagingSummary(previous?: { level: number; libraries: number }): { level: number; libraries: number } | undefined {
    const libraries = this.fallbackLibraries + (previous ? previous.libraries : 0);
    if (libraries === 0) {
      return undefined;
    }
    return { level: Math.max(this.highestLevelUsed, previous ? previous.level : 0), libraries };
  }

  /**
   * Scans again only what failed last time: libraries with an error, and the
   * libraries and subsites of sites whose listing failed. Ages stay measured
   * from the original scan's start, so the merged result is consistent.
   */
  public async retryFailed(previous: IScanResult, options: IScanOptions, onProgress: (p: IScanProgress) => void): Promise<IScanResult> {
    this.reset();
    const scanStart = new Date(previous.scanStartedAt);
    const quick = { ...options, quickAfterMonths: previous.quickAfterMonths };
    const progress = this.newProgress(previous.rootUrl);
    progress.phase = 'discovering';
    onProgress({ ...progress });

    const webs = previous.webs.map((w) => ({ ...w }));
    const knownLibraries = new Set(previous.libraries.map((l) => `${l.webUrl}|${l.id}`));
    const knownWebs = new Set(webs.map((w) => w.url));
    const toRead: ILibraryResult[] = [];

    // Failed libraries start again from scratch.
    const failedKeys = new Set(previous.libraries.filter((l) => l.error).map((l) => `${l.webUrl}|${l.id}`));
    const kept = previous.libraries.filter((l) => !failedKeys.has(`${l.webUrl}|${l.id}`));
    for (const lib of previous.libraries.filter((l) => failedKeys.has(`${l.webUrl}|${l.id}`))) {
      toRead.push({
        id: lib.id,
        title: lib.title,
        webTitle: lib.webTitle,
        webUrl: lib.webUrl,
        url: lib.url,
        itemCount: lib.itemCount,
        files: 0,
        bytes: 0,
        histogram: emptyHistogram()
      });
      progress.itemsExpected += lib.itemCount;
    }

    // Sites whose library list or subsites could not be read are listed again.
    for (const web of webs.filter((w) => w.error)) {
      const errors: string[] = [];
      try {
        for (const list of await this.getLibraries(web.url, options)) {
          if (!knownLibraries.has(`${web.url}|${list.Id}`)) {
            toRead.push(this.newLibrary(list, web.url, web.title));
            progress.itemsExpected += list.ItemCount;
          }
        }
      } catch (err) {
        if (err instanceof ScanCancelledError) {
          throw err;
        }
        errors.push(`Libraries: ${describeFailure(err)}`);
      }
      try {
        for (const sub of await this.getSubwebs(web.url)) {
          const url = this.toAbsoluteUrl(sub.ServerRelativeUrl);
          if (!knownWebs.has(url)) {
            const added: IWebResult[] = [];
            await this.discoverWebs(url, sub.Title, added, toRead, options, progress, onProgress);
            added.forEach((w) => {
              knownWebs.add(w.url);
              webs.push(w);
            });
          }
        }
      } catch (err) {
        if (err instanceof ScanCancelledError) {
          throw err;
        }
        errors.push(`Subsites: ${describeFailure(err)}`);
      }
      web.error = errors.length ? errors.join(' | ') : undefined;
    }
    progress.librariesFound = toRead.length;

    progress.phase = 'reading-files';
    const retriedRoots = toRead.map((l) => decodeURIComponent(new URL(l.url).pathname) + '/');
    const largest = previous.largestOldFiles.filter((f) => !retriedRoots.some((root) => f.serverRelativeUrl.indexOf(root) === 0));
    await this.readLibraries(toRead, quick, scanStart, largest, progress, onProgress);

    progress.phase = 'completed';
    onProgress({ ...progress });
    return {
      ...previous,
      webs,
      libraries: kept.concat(toRead),
      largestOldFiles: this.finishLargest(largest),
      scannedBy: this.context.pageContext.user.displayName,
      paging: this.pagingSummary(previous.paging)
    };
  }
}
