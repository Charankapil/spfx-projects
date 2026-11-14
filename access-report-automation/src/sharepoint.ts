import { randomUUID } from 'node:crypto';
import { TokenProvider } from './auth.js';
import { Http, HttpError, HttpRequest } from './http.js';

/** SharePoint REST with app-only tokens: reading (with paging) and the few file operations the report needs. */

export function trimSlash(u: string): string {
  return u.replace(/\/+$/, '');
}

export function originOf(url: string): string {
  const m = /^(https?:\/\/[^/?#]+)/i.exec(url);
  if (!m) {
    throw new Error(`Not an absolute address: ${url}`);
  }
  return m[1].toLowerCase();
}

/** A value inside an OData string literal. */
export function odataString(v: string): string {
  return v.replace(/'/g, "''");
}

/** A parameter alias, e.g. alias('p', '/sites/a b') -> @p='%2Fsites%2Fa%20b'. */
export function alias(name: string, value: string): string {
  return `@${name}='${encodeURIComponent(odataString(value))}'`;
}

/** Rows of a collection response, whatever OData flavour came back. */
export function rowsOf<T>(j: unknown): T[] {
  if (Array.isArray(j)) {
    return j as T[];
  }
  const o = j as { value?: unknown; d?: { results?: unknown }; results?: unknown } | undefined;
  if (!o) {
    return [];
  }
  if (Array.isArray(o.value)) {
    return o.value as T[];
  }
  if (o.d && Array.isArray(o.d.results)) {
    return o.d.results as T[];
  }
  if (Array.isArray(o.results)) {
    return o.results as T[];
  }
  return [];
}

export function entityOf<T>(j: unknown): T {
  const o = j as { d?: unknown } | undefined;
  return (o && o.d ? o.d : o) as T;
}

function nextLink(j: unknown): string | undefined {
  const o = j as Record<string, unknown> | undefined;
  if (!o) {
    return undefined;
  }
  const d = o.d as Record<string, unknown> | undefined;
  const v = o['odata.nextLink'] || o['@odata.nextLink'] || (d && d.__next);
  return typeof v === 'string' && v ? v : undefined;
}

/** Decoded server-relative path of an absolute address ("/sites/a b"; "/" for the root). */
export function serverRelative(absUrl: string): string {
  const u = new URL(absUrl);
  let p: string;
  try {
    p = decodeURIComponent(u.pathname);
  } catch {
    p = u.pathname;
  }
  p = p.replace(/\/{2,}/g, '/');
  return p.length > 1 ? p.replace(/\/+$/, '') : '/';
}

export function joinPath(...parts: string[]): string {
  const joined = parts
    .map((p) => p.replace(/^\/+|\/+$/g, ''))
    .filter((p) => !!p)
    .join('/');
  return '/' + joined;
}

export interface FileInfo {
  name: string;
  length: number;
  modified: string;
}

const CHUNK_THRESHOLD = 60 * 1024 * 1024;
const CHUNK_SIZE = 10 * 1024 * 1024;

export class SharePoint {
  constructor(private http: Http, private tokens: TokenProvider, private opts: { chunkThreshold?: number; chunkSize?: number } = {}) {}

  private async send<T>(url: string, req: HttpRequest): Promise<T> {
    const origin = originOf(url);
    const call = async (force: boolean): Promise<T> => {
      const token = await this.tokens.getToken(origin, force);
      const r = await this.http.request<T>(url, {
        ...req,
        headers: { Accept: 'application/json;odata=nometadata', Authorization: `Bearer ${token}`, ...(req.headers ?? {}) }
      });
      return r.data;
    };
    try {
      return await call(false);
    } catch (e) {
      // A token can expire between being issued and being used on a long run: sign in again, once.
      if (e instanceof HttpError && e.status === 401) {
        return call(true);
      }
      throw e;
    }
  }

  public get<T = unknown>(url: string): Promise<T> {
    return this.send<T>(url, {});
  }

  /** Every row of a collection, following SharePoint's paging links (same site only). */
  public async getAll<T>(url: string, max = 200000): Promise<T[]> {
    const origin = originOf(url);
    const out: T[] = [];
    let next: string | undefined = url;
    for (let page = 0; next && page < 500; page++) {
      const j: unknown = await this.get(next);
      rowsOf<T>(j).forEach((r) => out.push(r));
      if (out.length >= max) {
        break;
      }
      const link = nextLink(j);
      next = link && originOf(link) === origin ? link : undefined;
    }
    return out;
  }

  public post<T = unknown>(url: string, body?: string | Uint8Array, headers: Record<string, string> = {}, as: 'json' | 'none' = 'json'): Promise<T> {
    return this.send<T>(url, { method: 'POST', body, headers, as });
  }

  // ---- files ---------------------------------------------------------------------------------

  private api(siteUrl: string, path: string): string {
    return `${trimSlash(siteUrl)}/_api/${path}`;
  }

  public async folderExists(siteUrl: string, serverRelativePath: string): Promise<boolean> {
    try {
      const j = entityOf<{ Exists?: boolean }>(await this.get(this.api(siteUrl, `web/GetFolderByServerRelativePath(decodedurl=@p)?${alias('p', serverRelativePath)}&$select=Exists`)));
      return !!j && j.Exists !== false;
    } catch (e) {
      if (e instanceof HttpError && e.status === 404) {
        return false;
      }
      throw e;
    }
  }

  /**
   * Makes sure a folder exists, creating the missing levels below `existingRoot`
   * (the library itself is never created: a typo should fail, not make a new library).
   */
  public async ensureFolder(siteUrl: string, existingRoot: string, relative: string): Promise<string> {
    let current = existingRoot;
    for (const seg of relative.split('/').filter((s) => !!s)) {
      current = joinPath(current, seg);
      if (!(await this.folderExists(siteUrl, current))) {
        try {
          await this.post(this.api(siteUrl, `web/Folders/AddUsingPath(decodedurl=@p)?${alias('p', current)}`));
        } catch (e) {
          // Two runs creating the same folder: fine as long as it is there now.
          if (!(await this.folderExists(siteUrl, current))) {
            throw e;
          }
        }
      }
    }
    return current;
  }

  public async listFiles(siteUrl: string, folderPath: string): Promise<FileInfo[]> {
    try {
      const rows = await this.getAll<{ Name?: string; Length?: string | number; TimeLastModified?: string }>(
        this.api(siteUrl, `web/GetFolderByServerRelativePath(decodedurl=@p)/Files?${alias('p', folderPath)}&$select=Name,Length,TimeLastModified&$top=5000`)
      );
      return rows.map((r) => ({ name: String(r.Name || ''), length: Number(r.Length) || 0, modified: String(r.TimeLastModified || '') }));
    } catch (e) {
      if (e instanceof HttpError && e.status === 404) {
        return [];
      }
      throw e;
    }
  }

  /** undefined when the file does not exist. */
  public async download(siteUrl: string, filePath: string): Promise<Uint8Array | undefined> {
    try {
      return await this.send<Uint8Array>(this.api(siteUrl, `web/GetFileByServerRelativePath(decodedurl=@p)/$value?${alias('p', filePath)}`), { as: 'bytes' });
    } catch (e) {
      if (e instanceof HttpError && e.status === 404) {
        return undefined;
      }
      throw e;
    }
  }

  public async deleteFile(siteUrl: string, filePath: string): Promise<void> {
    try {
      await this.post(this.api(siteUrl, `web/GetFileByServerRelativePath(decodedurl=@p)?${alias('p', filePath)}`), undefined, { 'X-HTTP-Method': 'DELETE', 'If-Match': '*' }, 'none');
    } catch (e) {
      if (!(e instanceof HttpError && e.status === 404)) {
        throw e;
      }
    }
  }

  /** Uploads (overwriting). Large files go up in pieces. Returns the server-relative path. */
  public async upload(siteUrl: string, folderPath: string, name: string, content: Uint8Array): Promise<string> {
    const threshold = this.opts.chunkThreshold ?? CHUNK_THRESHOLD;
    const add = (body: Uint8Array): Promise<unknown> =>
      this.post(
        this.api(siteUrl, `web/GetFolderByServerRelativePath(decodedurl=@f)/Files/AddUsingPath(decodedurl=@n,overwrite=true)?${alias('f', folderPath)}&${alias('n', name)}`),
        body,
        { 'Content-Type': 'application/octet-stream' },
        'none'
      );
    const filePath = joinPath(folderPath, name);
    if (content.byteLength <= threshold) {
      await add(content);
      return filePath;
    }
    const size = this.opts.chunkSize ?? CHUNK_SIZE;
    let lastError: unknown;
    // A piece that fails half-way leaves the session in an unknown state: start the file again.
    for (let attempt = 0; attempt < 3; attempt++) {
      const uploadId = randomUUID();
      const file = this.api(siteUrl, `web/GetFileByServerRelativePath(decodedurl=@p)`);
      const q = `?${alias('p', filePath)}`;
      try {
        await add(new Uint8Array(0));
        for (let offset = 0; offset < content.byteLength; offset += size) {
          const piece = content.subarray(offset, Math.min(content.byteLength, offset + size));
          const last = offset + size >= content.byteLength;
          const method = offset === 0 ? `StartUpload(uploadId=guid'${uploadId}')` : last ? `FinishUpload(uploadId=guid'${uploadId}',fileOffset=${offset})` : `ContinueUpload(uploadId=guid'${uploadId}',fileOffset=${offset})`;
          await this.send(`${file}/${method}${q}`, { method: 'POST', body: piece, headers: { 'Content-Type': 'application/octet-stream' }, as: 'none', retry: 'throttled' });
          if (offset === 0 && last) {
            // A single piece still has to be closed.
            await this.send(`${file}/FinishUpload(uploadId=guid'${uploadId}',fileOffset=${piece.byteLength})${q}`, { method: 'POST', as: 'none', retry: 'throttled' });
          }
        }
        return filePath;
      } catch (e) {
        lastError = e;
        await this.send(`${file}/CancelUpload(uploadId=guid'${uploadId}')${q}`, { method: 'POST', as: 'none', retry: 'throttled' }).catch(() => undefined);
      }
    }
    throw lastError;
  }
}
