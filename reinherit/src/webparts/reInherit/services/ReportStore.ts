import { WebPartContext } from '@microsoft/sp-webpart-base';

import { IReportSummary, IRunReport } from '../models/IRunReport';
import { quoteForUrl, SpRest } from './SpRest';

const FOLDER_NAME = 'ReInherit';
const INDEX_FILE = 'reinherit-index.json';
const MAX_INDEX_ENTRIES = 200;
const SCHEMA_VERSION = 1;

interface IIndexFile {
  schemaVersion: number;
  reports: IReportSummary[];
}

interface ISavedReport {
  schemaVersion: number;
  report: IRunReport;
}

interface IListRootFolder {
  RootFolder: { ServerRelativeUrl: string };
}

export function summarizeReport(report: IRunReport, fileName: string): IReportSummary {
  return {
    id: report.id,
    fileName,
    mode: report.mode,
    runBy: report.runBy,
    startedAt: report.startedAt,
    completedAt: report.completedAt,
    cancelled: report.cancelled,
    scopes: report.scopes,
    stats: report.stats
  };
}

/**
 * Keeps every run's report as a JSON file in Site Assets/ReInherit on the
 * site collection's root web, plus a small index so the history list opens
 * without downloading every report. The reports name objects that had unique
 * permissions (and who had access, when backup is on), so they live in a
 * folder that inherits Site Assets' permissions - by default the site's
 * members can read it. Restrict the folder if that is too broad.
 */
export class ReportStore {
  private folderUrl: string | undefined;

  constructor(private context: WebPartContext, private rest: SpRest) {}

  private get rootUrl(): string {
    return this.context.pageContext.site.absoluteUrl;
  }

  /** Site Assets' URL name is "SiteAssets" regardless of the site's language. */
  private async findSiteAssets(): Promise<string | undefined> {
    const json = await this.rest.getJson<{ value?: IListRootFolder[] }>(
      `${this.rootUrl}/_api/web/lists?$filter=BaseTemplate eq 101&$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`
    );
    const match = (json.value || []).filter((l) => /\/siteassets$/i.test(l.RootFolder.ServerRelativeUrl))[0];
    return match ? match.RootFolder.ServerRelativeUrl : undefined;
  }

  private async getFolder(create: boolean): Promise<string | undefined> {
    if (this.folderUrl) {
      return this.folderUrl;
    }
    let assets = await this.findSiteAssets();
    if (!assets) {
      if (!create) {
        return undefined;
      }
      await this.rest.post(`${this.rootUrl}/_api/web/lists/EnsureSiteAssetsLibrary()`);
      assets = await this.findSiteAssets();
      if (!assets) {
        throw new Error('Site Assets library could not be found or created.');
      }
    }
    const folder = `${assets}/${FOLDER_NAME}`;
    if (create) {
      const existing = await this.rest.getText(
        `${this.rootUrl}/_api/web/GetFolderByServerRelativeUrl('${quoteForUrl(folder)}')?$select=Name`
      );
      if (existing === undefined) {
        await this.rest.post(
          `${this.rootUrl}/_api/web/GetFolderByServerRelativeUrl('${quoteForUrl(assets)}')/Folders/add('${FOLDER_NAME}')`
        );
      }
      this.folderUrl = folder;
    }
    return folder;
  }

  private async readFile(folder: string, fileName: string): Promise<string | undefined> {
    return this.rest.getText(
      `${this.rootUrl}/_api/web/GetFileByServerRelativeUrl('${quoteForUrl(`${folder}/${fileName}`)}')/$value`
    );
  }

  private async writeFile(folder: string, fileName: string, content: string): Promise<void> {
    await this.rest.post(
      `${this.rootUrl}/_api/web/GetFolderByServerRelativeUrl('${quoteForUrl(folder)}')` +
        `/Files/add(url='${fileName}',overwrite=true)`,
      content
    );
  }

  public async listReports(): Promise<IReportSummary[]> {
    const folder = await this.getFolder(false);
    if (!folder) {
      return [];
    }
    const text = await this.readFile(folder, INDEX_FILE);
    if (!text) {
      return [];
    }
    try {
      const index = JSON.parse(text) as IIndexFile;
      return index.schemaVersion === SCHEMA_VERSION && Array.isArray(index.reports) ? index.reports : [];
    } catch {
      return [];
    }
  }

  public async loadReport(summary: IReportSummary): Promise<IRunReport | undefined> {
    const folder = await this.getFolder(false);
    if (!folder) {
      return undefined;
    }
    const text = await this.readFile(folder, summary.fileName);
    if (!text) {
      return undefined;
    }
    const saved = JSON.parse(text) as ISavedReport;
    return saved && saved.schemaVersion === SCHEMA_VERSION && saved.report ? saved.report : undefined;
  }

  /** Writes (or overwrites) the report and puts it at the top of the index. */
  public async save(report: IRunReport): Promise<IReportSummary> {
    const folder = (await this.getFolder(true)) as string;
    const fileName = `reinherit-report-${report.id}.json`;
    const saved: ISavedReport = { schemaVersion: SCHEMA_VERSION, report };
    await this.writeFile(folder, fileName, JSON.stringify(saved));

    const summary = summarizeReport(report, fileName);
    const existing = (await this.listReports()).filter((r) => r.id !== report.id);
    const index: IIndexFile = {
      schemaVersion: SCHEMA_VERSION,
      reports: [summary, ...existing].slice(0, MAX_INDEX_ENTRIES)
    };
    await this.writeFile(folder, INDEX_FILE, JSON.stringify(index));
    return summary;
  }
}
