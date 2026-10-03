import { HttpError, odataString, SPClient, trimSlash } from '../core/SPClient';
import { AdminApi } from './AdminApi';
import { addPoint, emptyDoc, IGrowthDoc, IGrowthSettings, MAX_TRACKED, MIN_GAP_SEC, normalizeDoc } from './GrowthEngine';

export const HISTORY_FILE = 'admin-center-storage-history.json';
const CONCURRENCY = 2;
const FLUSH_EVERY = 100;
const ABORT_AFTER_FAILURES = 10;

export interface ICaptureOptions {
  /** Re-capture sites that already have a snapshot from the last few hours. */
  force?: boolean;
  onProgress?: (done: number, total: number, failed: number) => void;
  shouldCancel?: () => boolean;
}

export interface ICaptureResult {
  captured: number;
  skipped: number;
  failed: Array<{ url: string; message: string }>;
  cancelled: boolean;
  aborted: boolean;
}

/**
 * The history lives in ONE small JSON file in Site Assets on the site hosting
 * the web part, so reading it is one request and saving it is one request, no
 * matter how many sites are tracked. Capturing costs one light request per
 * tracked site (a site's storage is only exposed per site), paced by SPClient.
 * Every write re-reads the file first and merges, so two admins capturing at
 * the same time do not overwrite each other's points.
 */
export class GrowthStore {
  constructor(private client: SPClient, private api: AdminApi, private homeWebUrl: string) {}

  private serverRelativeWeb(): string {
    const m = /^https:\/\/[^/]+(\/[^?#]*)?/i.exec(trimSlash(this.homeWebUrl));
    return m && m[1] ? decodeURIComponent(trimSlash(m[1])) : '';
  }

  private get filePath(): string {
    return `${this.serverRelativeWeb()}/SiteAssets/${HISTORY_FILE}`;
  }

  public async load(fresh = false): Promise<IGrowthDoc> {
    if (fresh) {
      this.client.clearCache();
    }
    const url = `${trimSlash(this.homeWebUrl)}/_api/web/GetFileByServerRelativeUrl('${encodeURI(odataString(this.filePath))}')/$value`;
    try {
      return normalizeDoc(await this.client.get<unknown>(url, { cacheTtlMs: fresh ? 0 : 60000 }));
    } catch (e) {
      if (e instanceof HttpError && e.status === 404) {
        return emptyDoc();
      }
      if (e instanceof SyntaxError) {
        throw new Error('The storage history file could not be read (it is not valid JSON).');
      }
      throw e;
    }
  }

  private async save(doc: IGrowthDoc): Promise<void> {
    const folder = encodeURI(odataString(`${this.serverRelativeWeb()}/SiteAssets`));
    const url = `${trimSlash(this.homeWebUrl)}/_api/web/GetFolderByServerRelativeUrl('${folder}')/Files/add(url='${HISTORY_FILE}',overwrite=true)`;
    try {
      await this.client.post(url, JSON.stringify(doc), { raw: true }, trimSlash(this.homeWebUrl));
    } catch (e) {
      if (e instanceof HttpError && e.status === 404) {
        throw new Error('The Site Assets library was not found on this site, so the history cannot be saved. Add the web part to a site that has Site Assets.');
      }
      throw e;
    }
  }

  /** Re-read, change, write: the unit of every modification. */
  public async update(mutator: (doc: IGrowthDoc) => void): Promise<IGrowthDoc> {
    const doc = await this.load(true);
    mutator(doc);
    await this.save(doc);
    return doc;
  }

  public async track(sites: Array<{ url: string; title: string }>): Promise<{ added: number; capped: boolean }> {
    let added = 0;
    let capped = false;
    await this.update((doc) => {
      sites.forEach((s) => {
        const url = trimSlash(s.url);
        if (doc.sites[url]) {
          return;
        }
        if (Object.keys(doc.sites).length >= MAX_TRACKED) {
          capped = true;
          return;
        }
        doc.sites[url] = { title: s.title || url, s: [] };
        added++;
      });
    });
    return { added, capped };
  }

  public async untrack(urls: string[]): Promise<void> {
    await this.update((doc) => urls.forEach((u) => delete doc.sites[u]));
  }

  public async saveSettings(settings: IGrowthSettings): Promise<void> {
    await this.update((doc) => {
      doc.settings = settings;
    });
  }

  /** Takes a snapshot of every tracked site (one request each) and merges it into the file. */
  public async capture(opts: ICaptureOptions = {}): Promise<ICaptureResult> {
    const doc = await this.load(true);
    const nowSec = Math.floor(Date.now() / 1000);
    const urls = Object.keys(doc.sites);
    const todo = urls.filter((u) => {
      const pts = doc.sites[u].s;
      const last = pts[pts.length - 1];
      return opts.force || !last || nowSec - last[0] >= MIN_GAP_SEC;
    });
    const result: ICaptureResult = { captured: 0, skipped: urls.length - todo.length, failed: [], cancelled: false, aborted: false };
    let pending: Array<{ url: string; point: [number, number, number] }> = [];
    let next = 0;
    let consecutive = 0;

    const flush = async (): Promise<void> => {
      if (!pending.length) {
        return;
      }
      const batch = pending;
      pending = [];
      await this.update((d) => {
        batch.forEach((b) => {
          const site = d.sites[b.url];
          if (site) {
            site.s = addPoint(site.s, b.point);
          }
        });
        d.lastCapture = nowSec;
      });
    };

    const worker = async (): Promise<void> => {
      while (next < todo.length && !result.aborted) {
        if (opts.shouldCancel && opts.shouldCancel()) {
          result.cancelled = true;
          return;
        }
        const url = todo[next++];
        try {
          const u = await this.api.getSiteUsage(url);
          pending.push({ url, point: [Math.floor(Date.now() / 1000), u.bytes, Math.round(u.fraction * 1000)] });
          result.captured++;
          consecutive = 0;
        } catch (e) {
          result.failed.push({ url, message: (e as Error).message });
          if (++consecutive >= ABORT_AFTER_FAILURES) {
            // Something systemic (permissions, outage, heavy throttling): stop rather than hammer.
            result.aborted = true;
          }
        }
        if (opts.onProgress) {
          opts.onProgress(result.captured + result.failed.length, todo.length, result.failed.length);
        }
        if (pending.length >= FLUSH_EVERY) {
          await flush();
        }
      }
    };

    await Promise.all(Array.from({ length: CONCURRENCY }, () => worker()));
    await flush();
    return result;
  }
}
