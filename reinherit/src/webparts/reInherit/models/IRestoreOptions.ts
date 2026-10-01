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
  /**
   * gentle: one request at a time, at least a second apart - for very large
   * libraries and busy tenants. standard: two at a time, 150 ms apart.
   * Optional so reports saved by older versions still load.
   */
  speed?: ScanSpeed;
}

export type ScanSpeed = 'gentle' | 'standard';

export const DEFAULT_OPTIONS: IRestoreOptions = {
  recursive: true,
  includeSelected: true,
  includeSubsites: true,
  includeLists: true,
  includeFolders: true,
  includeFiles: true,
  maxDepth: 0,
  backupPermissions: true,
  speed: 'gentle'
};
