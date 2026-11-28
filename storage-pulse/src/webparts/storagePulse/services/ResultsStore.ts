import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { IScanResult, ScanScope } from '../models/IScanResult';
import { emptyIndex, ICentralIndex, removeSite, sanitizeIndex, summarize, upsert } from './centralIndex';
import { describeError } from './httpErrors';
import { sanitizeResult } from './safeData';

/** v1: 1.0 - 1.1 (Storage Activity Analyser). v2: adds scan options. Both load. */
const SCHEMA_VERSION = 2;
const READABLE_SCHEMAS = [1, 2];

interface ISavedScanFile {
  schemaVersion: number;
  result: IScanResult;
}

interface IListRootFolder {
  RootFolder: { ServerRelativeUrl: string };
}

interface IFileInfo {
  TimeLastModified?: string;
  ModifiedBy?: { Title?: string };
}

export interface ISavedScan {
  result: IScanResult;
  /** Who last wrote the file and when, according to SharePoint itself (not the file's contents). */
  savedBy?: string;
  savedAt?: string;
}

const INDEX_FILE = 'storage-pulse-central-index.json';

/** A file name for another site collection's results: readable, and unique through a hash of the full address. */
export function siteFileStem(siteUrl: string): string {
  const key = siteUrl.replace(/\/+$/, '').toLowerCase();
  let hash = 0x811c9dc5;
  for (let i = 0; i < key.length; i++) {
    hash ^= key.charCodeAt(i);
    // hash * 16777619 (the FNV prime) in 32 bits, without Math.imul (ES5).
    hash = (hash + (hash << 1) + (hash << 4) + (hash << 7) + (hash << 8) + (hash << 24)) >>> 0;
  }
  let path = key;
  try {
    path = decodeURIComponent(new URL(key).pathname);
  } catch {
    // Not a URL: the hash still keeps it unique.
  }
  const slug = path
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .substring(0, 60);
  return `storage-pulse-sc-${slug || 'root'}-${('0000000' + hash.toString(16)).slice(-8)}`;
}

function quoteForUrl(serverRelativeUrl: string): string {
  return encodeURIComponent(serverRelativeUrl.replace(/'/g, "''")).replace(/%2F/g, '/');
}

/**
 * Keeps the latest scan as a JSON file in the Site Assets library of the web
 * hosting the page, so everyone who opens the page sees the last results
 * instead of waiting for a new scan. Saving uses the site owner's normal
 * rights to add a file to Site Assets; nothing extra is needed.
 *
 * Site members can usually edit Site Assets too, so the file is treated as
 * untrusted when read back (see safeData.ts), and "saved by" is taken from
 * SharePoint's own record of who last modified the file.
 */
export class ResultsStore {
  /**
   * `targetSiteUrl`: in central mode, the site collection the results are
   * about. They are still kept in the Site Assets of the page's own site, one
   * file per site collection, and listed in a central index.
   */
  constructor(private context: WebPartContext, private targetSiteUrl?: string) {}

  private get webUrl(): string {
    return this.context.pageContext.web.absoluteUrl;
  }

  private get origin(): string {
    return new URL(this.webUrl).origin;
  }

  private fileName(scope: ScanScope): string {
    if (this.targetSiteUrl) {
      return `${siteFileStem(this.targetSiteUrl)}.json`;
    }
    return `storage-pulse-${scope === 'currentWeb' ? 'site' : 'site-collection'}.json`;
  }

  /** Name used by v1.x (Storage Activity Analyser), read so earlier scans still show after upgrading. */
  private legacyFileName(scope: ScanScope): string {
    return `storage-activity-analyser-${scope === 'currentWeb' ? 'site' : 'site-collection'}.json`;
  }

  /** Site Assets' URL name is "SiteAssets" regardless of the site's language. */
  private async findSiteAssetsFolder(): Promise<string | undefined> {
    const response = await this.context.spHttpClient.get(
      `${this.webUrl}/_api/web/lists?$filter=BaseTemplate eq 101&$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) {
      throw new Error(`${response.status}: ${await describeError(response)}`);
    }
    const json = (await response.json()) as { value?: IListRootFolder[] };
    const match = (json.value || []).filter((l) => /\/siteassets$/i.test(l.RootFolder.ServerRelativeUrl))[0];
    return match ? match.RootFolder.ServerRelativeUrl : undefined;
  }

  private async loadFile(path: string): Promise<ISavedScan | undefined> {
    const fileUrl = `${this.webUrl}/_api/web/GetFileByServerRelativeUrl('${quoteForUrl(path)}')`;
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(`${fileUrl}/$value`, SPHttpClient.configurations.v1);
    if (response.status === 404) {
      return undefined;
    }
    if (!response.ok) {
      throw new Error(`${response.status}: ${await describeError(response)}`);
    }
    let saved: ISavedScanFile;
    try {
      saved = JSON.parse(await response.text()) as ISavedScanFile;
    } catch {
      throw new Error('the saved results file is not valid JSON');
    }
    if (!saved || READABLE_SCHEMAS.indexOf(saved.schemaVersion) < 0) {
      throw new Error('the saved results file is from an unknown version');
    }
    const result = sanitizeResult(saved.result, this.origin);
    if (!result) {
      throw new Error('the saved results file is damaged or was edited by hand');
    }

    const scan: ISavedScan = { result };
    try {
      const info = await this.context.spHttpClient.get(
        `${fileUrl}?$select=TimeLastModified&$expand=ModifiedBy`,
        SPHttpClient.configurations.v1
      );
      if (info.ok) {
        const json = (await info.json()) as IFileInfo;
        scan.savedBy = json.ModifiedBy && json.ModifiedBy.Title ? json.ModifiedBy.Title : undefined;
        scan.savedAt = json.TimeLastModified;
      }
    } catch {
      // The file itself loaded; who saved it is a nice-to-have.
    }
    return scan;
  }

  /** The last saved scan for this scope, undefined if there is none. Throws if one exists but cannot be used. */
  public async load(scope: ScanScope): Promise<ISavedScan | undefined> {
    const folder = await this.findSiteAssetsFolder();
    if (!folder) {
      return undefined;
    }
    const current = await this.loadFile(`${folder}/${this.fileName(scope)}`);
    if (current || this.targetSiteUrl) {
      return current;
    }
    return this.loadFile(`${folder}/${this.legacyFileName(scope)}`);
  }

  private partialFileName(scope: ScanScope): string {
    if (this.targetSiteUrl) {
      return `${siteFileStem(this.targetSiteUrl)}-inprogress.json`;
    }
    return `storage-pulse-${scope === 'currentWeb' ? 'site' : 'site-collection'}-inprogress.json`;
  }

  /**
   * Writes and deletions are queued, so an older checkpoint can never land
   * after (and overwrite) a newer one, and the checkpoint is deleted only
   * after the writes before it have finished.
   */
  private queue: Promise<void> = Promise.resolve();

  private enqueue(operation: () => Promise<void>): Promise<void> {
    const run = this.queue.then(operation);
    this.queue = run.catch(() => undefined);
    return run;
  }

  /** Saves the progress of an unfinished scan, separately from the last complete scan that everyone sees. */
  public savePartial(result: IScanResult): Promise<void> {
    return this.enqueue(() => this.writeFile(this.partialFileName(result.scope), result));
  }

  /**
   * The unfinished scan saved for this scope, if there is one. Anything
   * unusable is ignored: a checkpoint is a convenience, not something to warn about.
   */
  public async loadPartial(scope: ScanScope): Promise<ISavedScan | undefined> {
    try {
      const folder = await this.findSiteAssetsFolder();
      if (!folder) {
        return undefined;
      }
      const saved = await this.loadFile(`${folder}/${this.partialFileName(scope)}`);
      return saved && saved.result.partial ? saved : undefined;
    } catch {
      return undefined;
    }
  }

  public deletePartial(scope: ScanScope): Promise<void> {
    return this.enqueue(() => this.deleteFile(this.partialFileName(scope)));
  }

  private async deleteFile(fileName: string): Promise<void> {
    const folder = await this.findSiteAssetsFolder();
    if (!folder) {
      return;
    }
    const url = `${this.webUrl}/_api/web/GetFileByServerRelativeUrl('${quoteForUrl(`${folder}/${fileName}`)}')`;
    const response = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' }
    });
    if (!response.ok && response.status !== 404) {
      throw new Error(`${response.status}: ${await describeError(response)}`);
    }
  }

  public async save(result: IScanResult): Promise<void> {
    await this.queue;
    await this.writeFile(this.fileName(result.scope), result);
    if (this.targetSiteUrl) {
      // The overview's line for this site; saved after the results themselves, so it never points at nothing.
      await this.updateIndex((index) => upsert(index, summarize(result)));
    }
  }

  /** The central overview: one summary per site collection scanned from this page. Empty when there is none yet. */
  public async loadIndex(): Promise<ICentralIndex> {
    const folder = await this.findSiteAssetsFolder();
    if (!folder) {
      return emptyIndex();
    }
    const response = await this.context.spHttpClient.get(
      `${this.webUrl}/_api/web/GetFileByServerRelativeUrl('${quoteForUrl(`${folder}/${INDEX_FILE}`)}')/$value`,
      SPHttpClient.configurations.v1
    );
    if (response.status === 404) {
      return emptyIndex();
    }
    if (!response.ok) {
      throw new Error(`${response.status}: ${await describeError(response)}`);
    }
    let raw: unknown;
    try {
      raw = JSON.parse(await response.text());
    } catch {
      throw new Error('the central index file is not valid JSON');
    }
    return sanitizeIndex(raw, this.origin);
  }

  private updateIndex(change: (index: ICentralIndex) => ICentralIndex): Promise<void> {
    return this.enqueue(async () => {
      const next = change(await this.loadIndex());
      await this.writeText(INDEX_FILE, JSON.stringify(next));
    });
  }

  /** Takes a site off the central overview and deletes its saved results (the site itself is not touched). */
  public async forgetSite(siteUrl: string): Promise<void> {
    const stem = siteFileStem(siteUrl);
    await this.updateIndex((index) => removeSite(index, siteUrl));
    await this.enqueue(() => this.deleteFile(`${stem}.json`));
    await this.enqueue(() => this.deleteFile(`${stem}-inprogress.json`));
  }

  private writeFile(fileName: string, result: IScanResult): Promise<void> {
    const body: ISavedScanFile = { schemaVersion: SCHEMA_VERSION, result };
    return this.writeText(fileName, JSON.stringify(body));
  }

  private async writeText(fileName: string, text: string): Promise<void> {
    let folder = await this.findSiteAssetsFolder();
    if (!folder) {
      const ensure = await this.context.spHttpClient.post(
        `${this.webUrl}/_api/web/lists/EnsureSiteAssetsLibrary()`,
        SPHttpClient.configurations.v1,
        {}
      );
      if (!ensure.ok) {
        throw new Error(`${ensure.status}: ${await describeError(ensure)}`);
      }
      folder = await this.findSiteAssetsFolder();
      if (!folder) {
        throw new Error('Site Assets library could not be found or created.');
      }
    }

    const response = await this.context.spHttpClient.post(
      `${this.webUrl}/_api/web/GetFolderByServerRelativeUrl('${quoteForUrl(folder)}')` +
        `/Files/add(url='${fileName}',overwrite=true)`,
      SPHttpClient.configurations.v1,
      { body: text }
    );
    if (!response.ok) {
      throw new Error(`${response.status}: ${await describeError(response)}`);
    }
  }
}
