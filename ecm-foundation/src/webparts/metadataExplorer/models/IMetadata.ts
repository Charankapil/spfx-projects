export type FieldKind = 'text' | 'choice' | 'taxonomy' | 'user' | 'lookup' | 'date' | 'boolean' | 'number' | 'builtin';

export interface IMetadataField {
  /** Internal name (libraries) or managed property (search). */
  key: string;
  title: string;
  kind: FieldKind;
}

export interface IDocumentRecord {
  key: string;
  name: string;
  /** Where the file lives, for display and the "open file" link. */
  url: string;
  openUrl: string;
  fileType: string;
  /** Library title (libraries source) or site title (search source). */
  location: string;
  locationUrl: string;
  modified?: Date;
  /** File size in bytes, when the source provides it. */
  size?: number;
  /** Every metadata value per field key, already flattened to display strings. */
  values: { [fieldKey: string]: string[] };
}

export interface ISourceInfo {
  title: string;
  url: string;
  itemCount: number;
  loaded: number;
  truncated: boolean;
  error?: string;
}

export interface IExplorerData {
  fields: IMetadataField[];
  documents: IDocumentRecord[];
  sources: ISourceInfo[];
  truncated: boolean;
  /** Short description of where the data came from, shown under the title. */
  scopeLabel: string;
}

export interface ILoadProgress {
  message: string;
  loaded: number;
}

/** Built-in categories every source provides alongside the real columns. */
export const LOCATION_KEY = '__location';
export const FILE_TYPE_KEY = '__fileType';
export const MODIFIED_YEAR_KEY = '__modifiedYear';
export const NO_VALUE = '(No value)';
