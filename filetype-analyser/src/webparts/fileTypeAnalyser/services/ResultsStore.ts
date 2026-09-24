import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { IScanSummary, ISiteCollectionOverview } from '../models/ISiteCollectionOverview';
import { describeError } from './httpErrors';

const FILE_NAME = 'file-type-analyser-scan.json';
const SCHEMA_VERSION = 1;

interface ISavedScanFile {
  schemaVersion: number;
  overview: ISiteCollectionOverview;
}

interface IListRootFolder {
  RootFolder: { ServerRelativeUrl: string };
}

function quoteForUrl(serverRelativeUrl: string): string {
  return encodeURIComponent(serverRelativeUrl.replace(/'/g, "''")).replace(/%2F/g, '/');
}

export function summarize(overview: ISiteCollectionOverview): IScanSummary {
  return {
    scanCompletedAt: overview.scanCompletedAt || overview.scanStartedAt,
    storageBytes: overview.storage.available ? overview.storage.usedBytes : undefined,
    totalFiles: overview.totalFiles,
    totalLibraries: overview.totalLibraries,
    totalWebs: overview.totalWebs,
    distinctTypes: overview.totalFileTypeStats.length
  };
}

/**
 * Persists the latest scan as a JSON file in the Site Assets library of the
 * site collection's root web, so everyone who opens the page sees the last
 * results instead of an empty web part. Anyone who can read Site Assets can
 * read the file, and it reflects what the person who ran the scan could see.
 */
export class ResultsStore {
  constructor(private context: WebPartContext) {}

  private get rootUrl(): string {
    return this.context.pageContext.site.absoluteUrl;
  }

  /** Site Assets' URL name is "SiteAssets" regardless of the site's language. */
  private async findSiteAssetsFolder(): Promise<string | undefined> {
    const url =
      `${this.rootUrl}/_api/web/lists?$filter=BaseTemplate eq 101` +
      `&$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`;
    const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) {
      throw new Error(`${response.status}: ${await describeError(response)}`);
    }
    const json = (await response.json()) as { value?: IListRootFolder[] };
    const match = (json.value || []).filter((l) => /\/siteassets$/i.test(l.RootFolder.ServerRelativeUrl))[0];
    return match ? match.RootFolder.ServerRelativeUrl : undefined;
  }

  /** Returns undefined when no scan has been saved yet. */
  public async load(): Promise<ISiteCollectionOverview | undefined> {
    const folder = await this.findSiteAssetsFolder();
    if (!folder) {
      return undefined;
    }
    const url = `${this.rootUrl}/_api/web/GetFileByServerRelativeUrl('${quoteForUrl(`${folder}/${FILE_NAME}`)}')/$value`;
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (response.status === 404) {
      return undefined;
    }
    if (!response.ok) {
      throw new Error(`${response.status}: ${await describeError(response)}`);
    }

    let saved: ISavedScanFile;
    try {
      saved = JSON.parse(await response.text());
    } catch {
      return undefined;
    }
    // Anything that is not a file this web part wrote is ignored rather than rendered.
    if (!saved || saved.schemaVersion !== SCHEMA_VERSION || !saved.overview || !saved.overview.rootWeb) {
      return undefined;
    }

    // JSON turned the Dates into ISO strings; turn them back.
    const revive = (value: unknown): Date => new Date(value as string);
    const overview = saved.overview;
    overview.scanStartedAt = revive(overview.scanStartedAt);
    overview.scanCompletedAt = overview.scanCompletedAt ? revive(overview.scanCompletedAt) : undefined;
    if (overview.previous) {
      overview.previous.scanCompletedAt = revive(overview.previous.scanCompletedAt);
    }
    return overview;
  }

  public async save(overview: ISiteCollectionOverview): Promise<void> {
    let folder = await this.findSiteAssetsFolder();
    if (!folder) {
      const ensure = await this.context.spHttpClient.post(
        `${this.rootUrl}/_api/web/lists/EnsureSiteAssetsLibrary()`,
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

    const body: ISavedScanFile = { schemaVersion: SCHEMA_VERSION, overview };
    const url =
      `${this.rootUrl}/_api/web/GetFolderByServerRelativeUrl('${quoteForUrl(folder)}')` +
      `/Files/add(url='${FILE_NAME}',overwrite=true)`;
    const response = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      body: JSON.stringify(body)
    });
    if (!response.ok) {
      throw new Error(`${response.status}: ${await describeError(response)}`);
    }
  }
}
