export type ObjectKind = 'web' | 'list' | 'library' | 'folder' | 'file' | 'item';

export type ObjectStatus = 'found' | 'restored' | 'failed' | 'excluded' | 'skipped';

/** Something with unique permissions that was found, and what happened to it. */
export interface IUniqueObject {
  key: string;
  kind: ObjectKind;
  name: string;
  /** Server-relative path (web URL for a web). */
  path: string;
  /** Absolute URL of the web the object lives in. */
  webUrl: string;
  listId?: string;
  listTitle?: string;
  itemId?: number;
  /** Folder levels below the list root (0 for webs and lists). */
  depth: number;
  status: ObjectStatus;
  /** Failure reason, or why it was skipped. */
  message?: string;
  /** "Member: Role, Role; Member: Role" as it was before the reset. */
  previousPermissions?: string;
  /** ISO timestamp of the reset. */
  restoredAt?: string;
}
