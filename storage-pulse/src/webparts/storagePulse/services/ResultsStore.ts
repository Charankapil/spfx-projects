import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { IScanResult, ScanScope } from '../models/IScanResult';
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
  constructor(private context: WebPartContext) {}

  private get webUrl(): string {
    return this.context.pageContext.web.absoluteUrl;
  }

  private get origin(): string {
    return new URL(this.webUrl).origin;
  }

  private fileName(scope: ScanScope): string {
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
    return current || this.loadFile(`${folder}/${this.legacyFileName(scope)}`);
  }

  private partialFileName(scope: ScanScope): string {
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
    return this.enqueue(async () => {
      const folder = await this.findSiteAssetsFolder();
      if (!folder) {
        return;
      }
      const url = `${this.webUrl}/_api/web/GetFileByServerRelativeUrl('${quoteForUrl(`${folder}/${this.partialFileName(scope)}`)}')`;
      const response = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' }
      });
      if (!response.ok && response.status !== 404) {
        throw new Error(`${response.status}: ${await describeError(response)}`);
      }
    });
  }

  public async save(result: IScanResult): Promise<void> {
    await this.queue;
    await this.writeFile(this.fileName(result.scope), result);
  }

  private async writeFile(fileName: string, result: IScanResult): Promise<void> {
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

    const body: ISavedScanFile = { schemaVersion: SCHEMA_VERSION, result };
    const response = await this.context.spHttpClient.post(
      `${this.webUrl}/_api/web/GetFolderByServerRelativeUrl('${quoteForUrl(folder)}')` +
        `/Files/add(url='${fileName}',overwrite=true)`,
      SPHttpClient.configurations.v1,
      { body: JSON.stringify(body) }
    );
    if (!response.ok) {
      throw new Error(`${response.status}: ${await describeError(response)}`);
    }
  }
}
