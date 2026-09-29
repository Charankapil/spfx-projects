export interface IRestoreOptions {
  /** Walk everything below each selected object, not just the object itself. */
  recursive: boolean;
  /** Reset the selected sites, lists and folders themselves too. */
  includeSelected: boolean;
  /** When recursive: which kinds of child object to look at. */
  includeSubsites: boolean;
  includeLists: boolean;
  includeFolders: boolean;
  includeFiles: boolean;
  /**
   * When recursive: how many folder levels below the selected object to go.
   * 0 means no limit. A file counts as being at its parent folder's depth + 1,
   * so a depth of 1 covers the top-level folders and files of a library.
   */
  maxDepth: number;
  /** Read and store each object's current role assignments before resetting it. */
  backupPermissions: boolean;
}

export const DEFAULT_OPTIONS: IRestoreOptions = {
  recursive: true,
  includeSelected: true,
  includeSubsites: true,
  includeLists: true,
  includeFolders: true,
  includeFiles: true,
  maxDepth: 0,
  backupPermissions: true
};
