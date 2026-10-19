import { SPHttpClient } from '@microsoft/sp-http';

import { SpRestClient } from '../../../common/services/SpRestClient';
import {
  FILE_TYPE_KEY,
  IDocumentRecord,
  IExplorerData,
  ILoadProgress,
  IMetadataField,
  LOCATION_KEY,
  MODIFIED_YEAR_KEY
} from '../models/IMetadata';
import { IMetadataSource } from './IMetadataSource';
import { isOfficeFile, parseDate, toValues, yearOf } from './valueParsing';

export type SearchScope = 'site' | 'hub' | 'tenant';

export interface ISearchSourceOptions {
  siteUrl: string;
  siteTitle: string;
  scope: SearchScope;
  hubSiteId?: string;
  /** "Title|ManagedProperty" per line. */
  fieldMappings: string;
  extraQuery: string;
  maxItems: number;
}

const BASE_PROPERTIES = [
  'Title',
  'Filename',
  'Path',
  'FileType',
  'SiteTitle',
  'SPWebUrl',
  'LastModifiedTime',
  'ServerRedirectedURL',
  'UniqueId'
];
const PAGE_SIZE = 500;
const EMPTY_GUID = '00000000-0000-0000-0000-000000000000';

export function parseFieldMappings(text: string): IMetadataField[] {
  const fields: IMetadataField[] = [];
  (text || '').split(/\r?\n/).forEach((line) => {
    const [title, property] = line.split('|').map((part) => (part || '').trim());
    if (title && property && /^[A-Za-z0-9_]+$/.test(property) && !fields.some((f) => f.key === property)) {
      fields.push({ key: property, title, kind: 'text' });
    }
  });
  return fields;
}

/**
 * Reads documents through the search index, so one web part can cover a
 * whole hub or the tenant. Columns appear as categories once they are mapped
 * to managed properties (RefinableStringNN) in the search schema — see the
 * README for the ECM mapping.
 */
export class SearchSource implements IMetadataSource {
  private rest: SpRestClient;

  constructor(httpClient: SPHttpClient, private options: ISearchSourceOptions) {
    this.rest = new SpRestClient(httpClient, 'Metadata Explorer');
  }

  private scopeQuery(): { kql: string; label: string } {
    const { scope, hubSiteId, siteUrl, siteTitle } = this.options;
    if (scope === 'hub' && hubSiteId && hubSiteId !== EMPTY_GUID) {
      return { kql: `(DepartmentId:{${hubSiteId}} OR DepartmentId:${hubSiteId})`, label: `Documents across the ${siteTitle} hub` };
    }
    if (scope === 'tenant') {
      return { kql: '', label: 'Documents across the organisation' };
    }
    return { kql: `SPSiteUrl:"${siteUrl}"`, label: `Documents in ${siteTitle}` };
  }

  public async load(onProgress: (progress: ILoadProgress) => void): Promise<IExplorerData> {
    const mapped = parseFieldMappings(this.options.fieldMappings);
    const scope = this.scopeQuery();
    const kql = ['IsDocument:1', scope.kql, this.options.extraQuery ? `(${this.options.extraQuery})` : '']
      .filter(Boolean)
      .join(' AND ');
    const selectProperties = BASE_PROPERTIES.concat(mapped.map((f) => f.key).filter((k) => BASE_PROPERTIES.indexOf(k) < 0));

    const documents: IDocumentRecord[] = [];
    let total = 0;
    let startRow = 0;
    for (;;) {
      onProgress({ message: 'Querying the search index', loaded: documents.length });
      const page = await this.rest.search(this.options.siteUrl, kql, {
        selectProperties,
        rowLimit: Math.min(PAGE_SIZE, this.options.maxItems - documents.length),
        startRow,
        sortList: 'LastModifiedTime:descending'
      });
      total = page.totalRows;
      page.rows.forEach((row) => documents.push(this.toDocument(row, mapped)));
      startRow += page.rows.length;
      if (!page.rows.length || startRow >= total || documents.length >= this.options.maxItems) {
        break;
      }
    }

    const builtIns: IMetadataField[] = [
      { key: LOCATION_KEY, title: 'Site', kind: 'builtin' },
      { key: FILE_TYPE_KEY, title: 'File type', kind: 'builtin' },
      { key: MODIFIED_YEAR_KEY, title: 'Modified (year)', kind: 'builtin' }
    ];
    const fields = builtIns.concat(mapped);

    return {
      fields,
      documents,
      sources: [
        {
          title: scope.label,
          url: this.options.siteUrl,
          itemCount: total,
          loaded: documents.length,
          truncated: total > documents.length
        }
      ],
      truncated: total > documents.length,
      scopeLabel: scope.label
    };
  }

  private toDocument(row: { [key: string]: string }, mapped: IMetadataField[]): IDocumentRecord {
    const fileType = (row.FileType || '').toLowerCase();
    const url = row.Path || '';
    const values: { [key: string]: string[] } = {
      [LOCATION_KEY]: row.SiteTitle ? [row.SiteTitle] : [],
      [FILE_TYPE_KEY]: fileType ? [fileType] : [],
      [MODIFIED_YEAR_KEY]: yearOf(row.LastModifiedTime)
    };
    mapped.forEach((f) => {
      values[f.key] = toValues(row[f.key], true);
    });
    return {
      key: row.UniqueId || url,
      name: row.Filename || row.Title || url.split('/').pop() || url,
      url,
      openUrl: isOfficeFile(fileType) && row.ServerRedirectedURL ? row.ServerRedirectedURL : url,
      fileType,
      location: row.SiteTitle || '',
      locationUrl: row.SPWebUrl || '',
      modified: parseDate(row.LastModifiedTime),
      values
    };
  }
}
