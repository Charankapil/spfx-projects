import { mkdir, readdir, readFile, rm, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { HttpError } from './http.js';
import { joinPath, SharePoint } from './sharepoint.js';

/**
 * Where reports and run state live: a SharePoint folder for scheduled runs, a local folder
 * for trial runs (--no-upload). Names are relative to the report folder ("csv/x.csv").
 */
export interface Store {
  readonly description: string;
  read(name: string): Promise<Uint8Array | undefined>;
  /** Returns where it ended up (the name may change when the existing file is open in Excel). */
  write(name: string, data: Uint8Array, opts?: { renameIfLocked?: boolean }): Promise<string>;
  remove(name: string): Promise<void>;
  /** File names directly in a sub-folder ('' = the report folder). */
  list(folder: string): Promise<string[]>;
}

export class LocalStore implements Store {
  public readonly description: string;
  constructor(private dir: string) {
    this.description = dir;
  }

  public async read(name: string): Promise<Uint8Array | undefined> {
    try {
      return await readFile(join(this.dir, name));
    } catch (e) {
      if ((e as NodeJS.ErrnoException).code === 'ENOENT') {
        return undefined;
      }
      throw e;
    }
  }

  public async write(name: string, data: Uint8Array): Promise<string> {
    const p = join(this.dir, name);
    await mkdir(dirname(p), { recursive: true });
    await writeFile(p, data);
    return p;
  }

  public async remove(name: string): Promise<void> {
    await rm(join(this.dir, name), { force: true });
  }

  public async list(folder: string): Promise<string[]> {
    try {
      return (await readdir(join(this.dir, folder), { withFileTypes: true })).filter((d) => d.isFile()).map((d) => d.name);
    } catch (e) {
      if ((e as NodeJS.ErrnoException).code === 'ENOENT') {
        return [];
      }
      throw e;
    }
  }
}

export class SharePointStore implements Store {
  public readonly description: string;
  private ensured = new Set<string>();
  private root: string | undefined;

  /**
   * @param libraryPath  server-relative path of the document library (must exist)
   * @param folder       folders below it, created when missing
   */
  constructor(private sp: SharePoint, private siteUrl: string, private libraryPath: string, private folder: string) {
    this.description = `${new URL(siteUrl).origin}${encodeURI(joinPath(libraryPath, folder))}`;
  }

  /** Checks the library and makes the report folder. Clear errors here save a long scan from failing at the end. */
  public async prepare(): Promise<void> {
    let exists: boolean;
    try {
      exists = await this.sp.folderExists(this.siteUrl, this.libraryPath);
    } catch (e) {
      if (e instanceof HttpError && (e.status === 401 || e.status === 403)) {
        throw new Error(`The app is not allowed to open the report site ${this.siteUrl} (${e.message}). It needs the SharePoint application permission Sites.FullControl.All with admin consent.`);
      }
      if (e instanceof HttpError && e.status === 404) {
        throw new Error(`The report site ${this.siteUrl} was not found. Check the address of the report folder.`);
      }
      throw e;
    }
    if (!exists) {
      throw new Error(`The document library ${this.libraryPath} does not exist on ${this.siteUrl}. Create it, or correct the address (use the library's address name, e.g. "Shared Documents").`);
    }
    this.root = await this.sp.ensureFolder(this.siteUrl, this.libraryPath, this.folder);
    this.ensured.add('');
  }

  private async folderFor(sub: string): Promise<string> {
    if (!this.root) {
      await this.prepare();
    }
    const root = this.root as string;
    if (!this.ensured.has(sub)) {
      await this.sp.ensureFolder(this.siteUrl, root, sub);
      this.ensured.add(sub);
    }
    return sub ? joinPath(root, sub) : root;
  }

  private split(name: string): { sub: string; leaf: string } {
    const i = name.lastIndexOf('/');
    return i < 0 ? { sub: '', leaf: name } : { sub: name.substring(0, i), leaf: name.substring(i + 1) };
  }

  public async read(name: string): Promise<Uint8Array | undefined> {
    if (!this.root) {
      await this.prepare();
    }
    return this.sp.download(this.siteUrl, joinPath(this.root as string, name));
  }

  public async write(name: string, data: Uint8Array, opts: { renameIfLocked?: boolean } = {}): Promise<string> {
    const { sub, leaf } = this.split(name);
    const folder = await this.folderFor(sub);
    try {
      return await this.sp.upload(this.siteUrl, folder, leaf, data);
    } catch (e) {
      // Someone has last month's re-run open in Excel: keep their copy, save ours beside it.
      if (opts.renameIfLocked && e instanceof HttpError && (e.status === 423 || /locked|checked out/i.test(e.message))) {
        const dot = leaf.lastIndexOf('.');
        const stamp = new Date().toISOString().substring(0, 16).replace(/[-:]/g, '').replace('T', '-');
        const alt = `${leaf.substring(0, dot)} (${stamp})${leaf.substring(dot)}`;
        return this.sp.upload(this.siteUrl, folder, alt, data);
      }
      throw e;
    }
  }

  public async remove(name: string): Promise<void> {
    if (!this.root) {
      await this.prepare();
    }
    await this.sp.deleteFile(this.siteUrl, joinPath(this.root as string, name));
  }

  public async list(folder: string): Promise<string[]> {
    if (!this.root) {
      await this.prepare();
    }
    return (await this.sp.listFiles(this.siteUrl, folder ? joinPath(this.root as string, folder) : (this.root as string))).map((f) => f.name);
  }
}
