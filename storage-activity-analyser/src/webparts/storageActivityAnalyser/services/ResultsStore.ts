import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { IScanResult, ScanScope } from '../models/IScanResult';
import { describeError } from './httpErrors';

const SCHEMA_VERSION = 1;

interface ISavedScanFile {
  schemaVersion: number;
  result: IScanResult;
}

interface IListRootFolder {
  RootFolder: { ServerRelativeUrl: string };
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
 * The file holds library names, sizes and counts, plus the names and paths of
 * the largest inactive files that the person who ran the scan could see.
 */
export class ResultsStore {
  constructor(private context: WebPartContext) {}

  private get webUrl(): string {
    return this.context.pageContext.web.absoluteUrl;
  }

  private fileName(scope: ScanScope): string {
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

  public async load(scope: ScanScope): Promise<IScanResult | undefined> {
    const folder = await this.findSiteAssetsFolder();
    if (!folder) {
      return undefined;
    }
    const url = `${this.webUrl}/_api/web/GetFileByServerRelativeUrl('${quoteForUrl(`${folder}/${this.fileName(scope)}`)}')/$value`;
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (response.status === 404) {
      return undefined;
    }
    if (!response.ok) {
      throw new Error(`${response.status}: ${await describeError(response)}`);
    }
    try {
      const saved = JSON.parse(await response.text()) as ISavedScanFile;
      if (!saved || saved.schemaVersion !== SCHEMA_VERSION || !saved.result || !Array.isArray(saved.result.libraries)) {
        return undefined;
      }
      return saved.result;
    } catch {
      return undefined;
    }
  }

  public async save(result: IScanResult): Promise<void> {
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
        `/Files/add(url='${this.fileName(result.scope)}',overwrite=true)`,
      SPHttpClient.configurations.v1,
      { body: JSON.stringify(body) }
    );
    if (!response.ok) {
      throw new Error(`${response.status}: ${await describeError(response)}`);
    }
  }
}
