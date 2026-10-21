import { SPHttpClient } from '@microsoft/sp-http';

import { SpRestClient } from '../../../common/services/SpRestClient';
import {
  FieldKind,
  FILE_TYPE_KEY,
  IDocumentRecord,
  IExplorerData,
  ILoadProgress,
  IMetadataField,
  ISourceInfo,
  LOCATION_KEY,
  MODIFIED_YEAR_KEY
} from '../models/IMetadata';
import { IMetadataSource } from './IMetadataSource';
import { encodePath, isOfficeFile, parseDate, toValues, yearOf } from './valueParsing';

interface IListInfo {
  Id: string;
  Title: string;
  ItemCount: number;
  EntityTypeName: string;
  RootFolder: { ServerRelativeUrl: string };
}

interface IFieldInfo {
  InternalName: string;
  Title: string;
  TypeAsString: string;
}

interface IRenderListDataResponse {
  Row?: { [key: string]: unknown }[];
  NextHref?: string;
}

/** Libraries SharePoint creates for its own use; they hold no business documents. */
const SYSTEM_LIBRARIES = ['SiteAssets', 'Style_x0020_Library', 'FormServerTemplates', 'PreservationHoldLibrary', 'IWConvertedForms'];

/**
 * Columns SharePoint adds by itself that are visible but are not business
 * metadata: Image Tags (AI image tagging) and the document ID columns.
 */
const SYSTEM_COLUMNS = ['lcf76f155ced4ddcb4097134ff3c332f', 'MediaServiceImageTags', 'TaxCatchAll', 'TaxCatchAllLabel', '_dlc_DocId', '_dlc_DocIdUrl', '_dlc_DocIdPersistId'];

/** Column types that make sensible categories. Notes, pictures, locations etc. do not. */
const KIND_BY_TYPE: { [type: string]: FieldKind } = {
  Text: 'text',
  Choice: 'choice',
  MultiChoice: 'choice',
  TaxonomyFieldType: 'taxonomy',
  TaxonomyFieldTypeMulti: 'taxonomy',
  User: 'user',
  UserMulti: 'user',
  Lookup: 'lookup',
  LookupMulti: 'lookup',
  DateTime: 'date',
  Boolean: 'boolean',
  Number: 'number',
  Currency: 'number',
  Calculated: 'text'
};

/** Columns SharePoint always has that still carry ECM meaning. */
const BUILT_IN_COLUMNS: IMetadataField[] = [
  { key: 'ContentType', title: 'Content type', kind: 'choice' },
  { key: '_ComplianceTag', title: 'Retention label', kind: 'choice' }
];

const PAGE_SIZE = 2000;

export interface ILibrariesSourceOptions {
  webUrl: string;
  webTitle: string;
  /** Library titles to include; empty means every document library. */
  libraryTitles: string[];
  excludedFields: string[];
  maxItems: number;
}

/**
 * Reads the document libraries of the current site directly. Needs no search
 * schema work: every custom column a library has becomes a category, and
 * values are the labels users see.
 */
export class LibrariesSource implements IMetadataSource {
  private rest: SpRestClient;
  private origin: string;

  constructor(httpClient: SPHttpClient, private options: ILibrariesSourceOptions) {
    this.rest = new SpRestClient(httpClient, 'Metadata Explorer');
    this.origin = /^(https?:\/\/[^/]+)/i.exec(options.webUrl)?.[1] || '';
  }

  public async load(onProgress: (progress: ILoadProgress) => void): Promise<IExplorerData> {
    const libraries = await this.getLibraries();
    const fieldsByKey: { [key: string]: IMetadataField } = {};
    const fieldOrder: string[] = [];
    const documents: IDocumentRecord[] = [];
    const sources: ISourceInfo[] = [];
    const excluded = this.options.excludedFields.map((f) => f.toLowerCase());

    for (const library of libraries) {
      const info: ISourceInfo = {
        title: library.Title,
        url: this.origin + library.RootFolder.ServerRelativeUrl,
        itemCount: library.ItemCount,
        loaded: 0,
        truncated: false
      };
      sources.push(info);
      const remaining = this.options.maxItems - documents.length;
      if (remaining <= 0) {
        info.truncated = library.ItemCount > 0;
        continue;
      }
      onProgress({ message: `Reading ${library.Title}`, loaded: documents.length });
      try {
        const fields = (await this.getFields(library.Id)).filter(
          (f) => excluded.indexOf(f.title.toLowerCase()) < 0 && excluded.indexOf(f.key.toLowerCase()) < 0
        );
        fields.forEach((f) => {
          if (!fieldsByKey[f.key]) {
            fieldsByKey[f.key] = f;
            fieldOrder.push(f.key);
          }
        });
        const docs = await this.getDocuments(library, fields, remaining, (n) =>
          onProgress({ message: `Reading ${library.Title}`, loaded: documents.length + n })
        );
        info.loaded = docs.length;
        info.truncated = docs.length >= remaining && library.ItemCount > docs.length;
        docs.forEach((d) => documents.push(d));
      } catch (err) {
        info.error = (err as Error).message;
      }
    }

    const builtIns: IMetadataField[] = [
      { key: LOCATION_KEY, title: 'Library', kind: 'builtin' },
      { key: FILE_TYPE_KEY, title: 'File type', kind: 'builtin' },
      { key: MODIFIED_YEAR_KEY, title: 'Modified (year)', kind: 'builtin' }
    ];
    return {
      fields: builtIns.concat(fieldOrder.map((k) => fieldsByKey[k])),
      documents,
      sources,
      truncated: sources.some((s) => s.truncated),
      scopeLabel: `Document libraries in ${this.options.webTitle}`
    };
  }

  private async getLibraries(): Promise<IListInfo[]> {
    const url =
      `${this.options.webUrl}/_api/web/lists` +
      `?$filter=BaseTemplate eq 101 and Hidden eq false and IsCatalog eq false` +
      `&$select=Id,Title,ItemCount,EntityTypeName,RootFolder/ServerRelativeUrl&$expand=RootFolder`;
    const json = await this.rest.getJson<{ value: IListInfo[] }>(url);
    const wanted = this.options.libraryTitles.map((t) => t.toLowerCase());
    return (json.value || [])
      .filter((l) => SYSTEM_LIBRARIES.indexOf(l.EntityTypeName) < 0)
      .filter((l) => !wanted.length || wanted.indexOf(l.Title.toLowerCase()) >= 0)
      .sort((a, b) => a.Title.localeCompare(b.Title));
  }

  private async getFields(listId: string): Promise<IMetadataField[]> {
    const builtInFilter = BUILT_IN_COLUMNS.map((c) => `InternalName eq '${c.key}'`).join(' or ');
    const url =
      `${this.options.webUrl}/_api/web/lists(guid'${listId}')/fields` +
      `?$filter=(Hidden eq false and FromBaseType eq false) or ${builtInFilter}` +
      `&$select=InternalName,Title,TypeAsString`;
    const json = await this.rest.getJson<{ value: IFieldInfo[] }>(url);
    const fields: IMetadataField[] = [];
    (json.value || []).forEach((f) => {
      if (SYSTEM_COLUMNS.indexOf(f.InternalName) >= 0) {
        return;
      }
      const builtIn = BUILT_IN_COLUMNS.filter((c) => c.key === f.InternalName)[0];
      if (builtIn) {
        fields.push(builtIn);
        return;
      }
      const kind = KIND_BY_TYPE[f.TypeAsString];
      if (kind) {
        fields.push({ key: f.InternalName, title: f.Title, kind });
      }
    });
    // Built-ins first, in a fixed order, so every library lists them alike.
    return fields.sort((a, b) => {
      const ai = BUILT_IN_COLUMNS.indexOf(a);
      const bi = BUILT_IN_COLUMNS.indexOf(b);
      return (ai < 0 ? 99 : ai) - (bi < 0 ? 99 : bi);
    });
  }

  /**
   * RenderListDataAsStream returns display values (taxonomy labels, user
   * names) and pages through libraries past the 5,000-item view threshold.
   * Folders are skipped client-side: filtering them out on the server would
   * need a query on a non-indexed column, which large libraries refuse.
   */
  private async getDocuments(
    library: IListInfo,
    fields: IMetadataField[],
    limit: number,
    onCount: (n: number) => void
  ): Promise<IDocumentRecord[]> {
    const viewFields = ['FileLeafRef', 'FileRef', 'FSObjType', 'File_x0020_Type', 'File_x0020_Size', 'Modified', 'UniqueId']
      .concat(fields.map((f) => f.key))
      .map((name) => `<FieldRef Name="${name}"/>`)
      .join('');
    const viewXml =
      `<View Scope="RecursiveAll"><ViewFields>${viewFields}</ViewFields>` +
      `<RowLimit Paged="TRUE">${PAGE_SIZE}</RowLimit></View>`;
    const baseUrl = `${this.options.webUrl}/_api/web/lists(guid'${library.Id}')/RenderListDataAsStream`;
    const body = { parameters: { RenderOptions: 2, ViewXml: viewXml, DatesInUtc: true } };

    const documents: IDocumentRecord[] = [];
    let next: string | undefined = '';
    while (next !== undefined && documents.length < limit) {
      const page: IRenderListDataResponse = await this.rest.postJson<IRenderListDataResponse>(baseUrl + next, body);
      (page.Row || []).forEach((row) => {
        if (String(row.FSObjType) !== '1' && documents.length < limit) {
          documents.push(this.toDocument(row, library, fields));
        }
      });
      onCount(documents.length);
      next = page.NextHref ? page.NextHref : undefined;
    }
    return documents;
  }

  private toDocument(row: { [key: string]: unknown }, library: IListInfo, fields: IMetadataField[]): IDocumentRecord {
    const fileRef = String(row.FileRef || '');
    const fileType = String(row.File_x0020_Type || '').toLowerCase();
    const url = this.origin + encodePath(fileRef);
    const modified = parseDate(row.Modified);
    const values: { [key: string]: string[] } = {
      [LOCATION_KEY]: [library.Title],
      [FILE_TYPE_KEY]: fileType ? [fileType] : [],
      [MODIFIED_YEAR_KEY]: yearOf(row.Modified)
    };
    fields.forEach((f) => {
      values[f.key] = f.kind === 'date' ? yearOf(row[f.key]) : toValues(row[f.key]);
    });
    return {
      key: String(row.UniqueId || fileRef),
      name: String(row.FileLeafRef || fileRef.split('/').pop() || ''),
      url,
      openUrl: isOfficeFile(fileType) ? `${url}?web=1` : url,
      fileType,
      location: library.Title,
      locationUrl: this.origin + library.RootFolder.ServerRelativeUrl,
      modified,
      size: parseInt(String(row.File_x0020_Size || ''), 10) || undefined,
      values
    };
  }
}
