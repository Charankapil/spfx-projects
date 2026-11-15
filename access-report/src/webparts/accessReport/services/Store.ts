import { entityOf, HttpError, odataString, rowsOf, SPClient, trimSlash } from '../core/SPClient';
import { ISubject, normaliseLogin } from './permissions';
import { BOM, stampOf } from './reportCsv';

/**
 * Where the report's configuration and output live: two small lists and a document library
 * in the site that hosts the web part. The web part creates them (ordinary site-owner
 * rights, plain SharePoint REST) and the Power Automate flow reads the same lists, so the
 * page is the only place anything is configured.
 *
 *   Lists/AccessReportSubjects   one item per person / group to report on
 *   Lists/AccessReportSettings   a single item: options, and the flow's last-run note
 *   AccessReports                the default folder the report files are saved in
 *
 * The flow addresses the lists by these URLs and columns: keep flow/definition.ts in step.
 */

export const SUBJECTS_LIST = 'Lists/AccessReportSubjects';
export const SETTINGS_LIST = 'Lists/AccessReportSettings';
export const REPORTS_LIBRARY = 'AccessReports';

export const F_LOGIN = 'ARLogin';
export const F_ACTIVE = 'ARActive';
export const F_LIBRARIES = 'ARIncludeLibraries';
export const F_FILTER = 'ARSiteFilter';
export const F_FOLDER = 'ARReportFolder';
export const F_LAST_RUN = 'ARLastRun';

export interface IStoredSubject extends ISubject {
  id: number;
  active: boolean;
}

export interface ISettings {
  id: number;
  includeLibraries: boolean;
  siteFilter: string;
  /** Server-relative folder the report files are saved in. */
  reportFolder: string;
  /** Written by the flow: when it last started / finished, and what it produced. */
  lastRun: string;
}

export interface IReportFile {
  name: string;
  serverRelativeUrl: string;
  size: number;
  modified?: Date;
  modifiedBy: string;
  /** Run time from the file name. */
  at: Date;
  stamp: string;
  unread: boolean;
}

interface IListDef {
  url: string;
  urlName: string;
  title: string;
  template: number;
  description: string;
  fields: string[];
}

const field = (type: string, name: string, display: string, extra = '', inner = ''): string =>
  `<Field Type="${type}" Name="${name}" StaticName="${name}" DisplayName="${display}" ${extra}>${inner}</Field>`;

const LISTS: IListDef[] = [
  {
    url: SUBJECTS_LIST,
    urlName: 'AccessReportSubjects',
    title: 'Access Report Subjects',
    template: 100,
    description: 'People and groups the Access Report covers. Managed by the Access Report web part.',
    fields: [field('Text', F_LOGIN, 'Login name', 'MaxLength="255"'), field('Boolean', F_ACTIVE, 'Include in report', '', '<Default>1</Default>')]
  },
  {
    url: SETTINGS_LIST,
    urlName: 'AccessReportSettings',
    title: 'Access Report Settings',
    template: 100,
    description: 'Options of the Access Report (one item). Managed by the Access Report web part.',
    fields: [
      field('Boolean', F_LIBRARIES, 'Include libraries', '', '<Default>1</Default>'),
      field('Text', F_FILTER, 'Only sites containing', 'MaxLength="255"'),
      field('Text', F_FOLDER, 'Report folder', 'MaxLength="255"'),
      field('Note', F_LAST_RUN, 'Last scheduled run', 'RichText="FALSE" NumLines="4"')
    ]
  },
  { url: REPORTS_LIBRARY, urlName: REPORTS_LIBRARY, title: 'Access Reports', template: 101, description: 'Access Report files, one per run.', fields: [] }
];

const FIELD_NAME = /Name="([^"]+)"/;

/** A server-relative path as a parameter alias value: ...(@p)?@p='<this>'. */
export function aliasValue(path: string): string {
  return `'${encodeURIComponent(odataString(path))}'`;
}

/** Folder paths that the classic file APIs cannot address. */
export function folderProblem(path: string, webPath?: string): string | undefined {
  const p = path.trim();
  if (!p) {
    return undefined;
  }
  if (p.charAt(0) !== '/') {
    return 'Start with / (a server-relative path such as /sites/it/AccessReports).';
  }
  if (webPath !== undefined && (p + '/').toLowerCase().indexOf(webPath.toLowerCase() + '/') !== 0) {
    return `The folder must be in this site: start with ${webPath}/.`;
  }
  if (/[#%*:<>?|"\\]/.test(p)) {
    return 'The path cannot contain # % * : < > ? | " or \\.';
  }
  return undefined;
}

export class Store {
  public readonly web: string;
  public readonly webPath: string;

  constructor(private client: SPClient, homeWebUrl: string) {
    this.web = trimSlash(homeWebUrl);
    this.webPath = this.web.replace(/^https:\/\/[^/]+/i, '');
  }

  public get defaultFolder(): string {
    return `${this.webPath}/${REPORTS_LIBRARY}`;
  }

  private list(relative: string, tail = ''): string {
    const join = tail.indexOf('?') >= 0 ? '&' : '?';
    return `${this.web}/_api/web/GetList(@u)${tail}${join}@u=${aliasValue(`${this.webPath}/${relative}`)}`;
  }

  private folder(path: string, tail = ''): string {
    const join = tail.indexOf('?') >= 0 ? '&' : '?';
    return `${this.web}/_api/web/GetFolderByServerRelativeUrl(@f)${tail}${join}@f=${aliasValue(path)}`;
  }

  // ---- set-up ------------------------------------------------------------------------

  private async exists(relative: string): Promise<boolean> {
    try {
      await this.client.get(this.list(relative, '?$select=Id'), { cacheTtlMs: 0 });
      return true;
    } catch (e) {
      if (e instanceof HttpError && (e.status === 404 || e.status === 400)) {
        return false;
      }
      throw e;
    }
  }

  /** What is still missing (empty = ready). */
  public async missing(): Promise<string[]> {
    const out: string[] = [];
    for (const l of LISTS) {
      if (!(await this.exists(l.url))) {
        out.push(l.title);
      }
    }
    return out;
  }

  /** Creates whatever is missing. Safe to run again: existing lists, columns and items are kept. */
  public async provision(onStep?: (text: string) => void): Promise<void> {
    for (const l of LISTS) {
      if (!(await this.exists(l.url))) {
        if (onStep) {
          onStep(`Creating ${l.title}…`);
        }
        // Created under its URL name so the address has no spaces, then given its display title.
        await this.client.post(`${this.web}/_api/web/lists`, { Title: l.urlName, BaseTemplate: l.template, Description: l.description });
        await this.client.post(this.list(l.url), { Title: l.title }, { method: 'MERGE' });
      }
      if (!l.fields.length) {
        continue;
      }
      const have = rowsOf<{ InternalName?: string }>(await this.client.get(this.list(l.url, '/fields?$select=InternalName&$top=5000'), { cacheTtlMs: 0 })).map((f) => String(f.InternalName));
      for (const xml of l.fields) {
        const name = (FIELD_NAME.exec(xml) || ['', ''])[1];
        if (have.indexOf(name) < 0) {
          // Options 8 = keep the Name attribute as the internal name.
          await this.client.post(this.list(l.url, '/fields/createfieldasxml'), { parameters: { SchemaXml: xml, Options: 8 } });
        }
      }
    }
    await this.settings();
  }

  // ---- settings ----------------------------------------------------------------------

  /** The single settings item, created with defaults the first time. */
  public async settings(): Promise<ISettings> {
    const read = async (): Promise<{ [k: string]: unknown } | undefined> =>
      rowsOf<{ [k: string]: unknown }>(await this.client.get(this.list(SETTINGS_LIST, `/items?$select=Id,${F_LIBRARIES},${F_FILTER},${F_FOLDER},${F_LAST_RUN}&$orderby=Id&$top=1`), { cacheTtlMs: 0 }))[0];
    let item = await read();
    if (!item) {
      await this.client.post(this.list(SETTINGS_LIST, '/items'), { Title: 'Settings', [F_LIBRARIES]: true, [F_FILTER]: '', [F_FOLDER]: this.defaultFolder });
      item = await read();
    }
    if (!item) {
      throw new Error('The settings item could not be created.');
    }
    return {
      id: Number(item.Id),
      includeLibraries: item[F_LIBRARIES] !== false,
      siteFilter: String(item[F_FILTER] || ''),
      reportFolder: String(item[F_FOLDER] || '') || this.defaultFolder,
      lastRun: String(item[F_LAST_RUN] || '')
    };
  }

  public async saveSettings(s: ISettings): Promise<void> {
    await this.client.post(this.list(SETTINGS_LIST, `/items(${s.id})`), { [F_LIBRARIES]: s.includeLibraries, [F_FILTER]: s.siteFilter.trim(), [F_FOLDER]: trimSlash(s.reportFolder.trim()) || this.defaultFolder }, { method: 'MERGE' });
  }

  // ---- subjects ----------------------------------------------------------------------

  public async subjects(): Promise<IStoredSubject[]> {
    return rowsOf<{ [k: string]: unknown }>(await this.client.get(this.list(SUBJECTS_LIST, `/items?$select=Id,Title,${F_LOGIN},${F_ACTIVE}&$orderby=Id&$top=5000`), { cacheTtlMs: 0 })).map((i) => ({
      id: Number(i.Id),
      label: String(i.Title || ''),
      login: String(i[F_LOGIN] || '').trim(),
      active: i[F_ACTIVE] === true
    }));
  }

  public async addSubject(s: ISubject): Promise<void> {
    await this.client.post(this.list(SUBJECTS_LIST, '/items'), { Title: s.label.trim().substring(0, 255), [F_LOGIN]: normaliseLogin(s.login), [F_ACTIVE]: true });
  }

  public async setSubjectActive(id: number, active: boolean): Promise<void> {
    await this.client.post(this.list(SUBJECTS_LIST, `/items(${id})`), { [F_ACTIVE]: active }, { method: 'MERGE' });
  }

  public async removeSubject(id: number): Promise<void> {
    await this.client.post(this.list(SUBJECTS_LIST, `/items(${id})`), undefined, { method: 'DELETE' });
  }

  // ---- report files ------------------------------------------------------------------

  public async folderExists(path: string): Promise<boolean> {
    try {
      const f = entityOf<{ Exists?: boolean }>(await this.client.get(this.folder(path, '?$select=Exists'), { cacheTtlMs: 0 }));
      return f.Exists !== false;
    } catch (e) {
      if (e instanceof HttpError && (e.status === 404 || e.status === 400)) {
        return false;
      }
      throw e;
    }
  }

  /** Saves (or replaces) a text file in the folder; returns its server-relative URL. */
  public async saveFile(folder: string, name: string, text: string): Promise<string> {
    const body = text.charAt(0) === BOM ? text : BOM + text;
    await this.client.post(this.folder(folder, `/Files/add(url=@n,overwrite=true)?@n=${aliasValue(name)}`), body, { raw: true });
    return `${trimSlash(folder)}/${name}`;
  }

  /** The report files in the folder, newest run first. */
  public async reports(folder: string): Promise<IReportFile[]> {
    const files = rowsOf<{ Name?: string; ServerRelativeUrl?: string; Length?: string | number; TimeLastModified?: string; ModifiedBy?: { Title?: string } }>(
      await this.client.get(this.folder(folder, '/Files?$select=Name,ServerRelativeUrl,Length,TimeLastModified,ModifiedBy/Title&$expand=ModifiedBy&$top=5000'), { cacheTtlMs: 0 })
    );
    const out: IReportFile[] = [];
    files.forEach((f) => {
      const s = stampOf(String(f.Name || ''));
      if (!s) {
        return;
      }
      const modified = f.TimeLastModified ? new Date(f.TimeLastModified) : undefined;
      out.push({
        name: String(f.Name),
        serverRelativeUrl: String(f.ServerRelativeUrl || ''),
        size: Number(f.Length) || 0,
        modified: modified && !isNaN(modified.getTime()) ? modified : undefined,
        modifiedBy: (f.ModifiedBy && f.ModifiedBy.Title) || '',
        at: s.at,
        stamp: s.stamp,
        unread: s.unread
      });
    });
    return out.sort((a, b) => (a.stamp < b.stamp ? 1 : a.stamp > b.stamp ? -1 : a.unread === b.unread ? 0 : a.unread ? 1 : -1));
  }

  public readFile(serverRelativeUrl: string): Promise<string> {
    return this.client.getText(`${this.web}/_api/web/GetFileByServerRelativeUrl(@f)/$value?@f=${aliasValue(serverRelativeUrl)}`);
  }
}
