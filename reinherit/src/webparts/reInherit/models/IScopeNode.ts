/** A web, list/library or folder shown in the scope picker. */
export type ScopeKind = 'web' | 'library' | 'list' | 'folder';

export interface IScopeNode {
  /** Unique within the tree: web URL, list id or folder path. */
  key: string;
  kind: ScopeKind;
  title: string;
  /** Absolute URL of the web the node lives in (the web itself for a web). */
  webUrl: string;
  serverRelativeUrl: string;
  /** Lists, libraries and folders. */
  listId?: string;
  /** Folders: the folder's list item id, needed to reset the folder itself. */
  itemId?: number;
  /** Lists and libraries: the list's root folder path. */
  listRootUrl?: string;
  itemCount?: number;
  hasUniquePermissions?: boolean;
  isRootWeb?: boolean;
  /** Children are loaded on first expand. */
  children?: IScopeNode[];
  loading?: boolean;
  error?: string;
}
