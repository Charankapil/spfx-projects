import { entityOf, HttpError, odataString, rowsOf, SPClient, trimSlash } from '../core/SPClient';
import { IGrowthSettings } from './GrowthEngine';
import { parseCsv } from './csv';
import { buildSnapshot, IBuildStats, IMapping, ISnapshotData } from './StorageImport';
import { analyzeTenant, ISnapshotMeta, ITenantOverview, ITenantSummary, MAX_SNAPSHOTS_KEPT, selectSnapshots, summarizeInventory } from './TenantGrowth';
import { siteType } from './siteTypes';
import { MIN_GAP_SEC } from './GrowthEngine';
import { toDate } from './SearchApi';

export const INDEX_FILE = 'admin-center-storage-index.json';
const SNAP_PREFIX = 'admin-center-snap-';

export interface ITenantIndex {
  v: 1;
  snapshots: ISnapshotMeta[];
  summary?: ITenantSummary;
  mapping?: IMapping;
  /** Library folder (server-relative) the last import came from, for one-click re-import. */
  folder?: string;
  lastFile?: { name: string; modified: number };
  /** Inventory figures for the latest snapshot (the tenant dashboard). */
  overview?: ITenantOverview;
  /** Tenant SharePoint storage capacity in TB, entered by an admin (shared). */
  capacityTB?: number;
  /** Count deleted sites (still in the recycle bin) in "storage used". Default true. */
  countDeleted?: boolean;
}

export interface ICsvFile {
  name: string;
  serverRelativeUrl: string;
  modified?: Date;
  size: number;
}

export interface IImportResult {
  stats: IBuildStats;
  summary: ITenantSummary;
  replaced: boolean;
}

function emptyIndex(): ITenantIndex {
  return { v: 1, snapshots: [] };
}

function normalizeIndex(raw: unknown): ITenantIndex {
  const r = raw as Partial<ITenantIndex> | null;
  if (!r || typeof r !== 'object' || !Array.isArray(r.snapshots)) {
    return emptyIndex();
  }
  return {
    v: 1,
    snapshots: r.snapshots.filter((s) => s && typeof s.file === 'string' && isFinite(Number(s.t))).map((s) => ({ ...s, t: Number(s.t), sites: Number(s.sites) || 0, bytes: Number(s.bytes) || 0 })),
    summary: r.summary,
    mapping: r.mapping,
    folder: r.folder,
    lastFile: r.lastFile,
    overview: r.overview,
    capacityTB: isFinite(Number(r.capacityTB)) && Number(r.capacityTB) > 0 ? Number(r.capacityTB) : undefined,
    countDeleted: r.countDeleted
  };
}

/**
 * Imported-snapshot storage for tenants too large to poll site by site.
 * Each import writes one immutable snapshot file (~0.5 MB for 17,000 sites)
 * and updates a small index that carries the precomputed analysis, so the
 * Dashboard reads one tiny file. An import costs roughly a dozen requests in
 * total (list, read CSV, write snapshot, read up to 7 earlier snapshots, write
 * index), however many sites the CSV holds.
 */
export class TenantStore {
  private origin: string;

  constructor(private client: SPClient, private homeWebUrl: string) {
    this.origin = (/^(https:\/\/[^/]+)/i.exec(homeWebUrl) || [''])[1];
  }

  private get web(): string {
    return trimSlash(this.homeWebUrl);
  }

  private serverRelativeWeb(): string {
    const m = /^https:\/\/[^/]+(\/[^?#]*)?/i.exec(this.web);
    return m && m[1] ? decodeURIComponent(trimSlash(m[1])) : '';
  }

  private assetsFolder(): string {
    return `${this.serverRelativeWeb()}/SiteAssets`;
  }

  private fileUrl(serverRelative: string): string {
    return `${this.web}/_api/web/GetFileByServerRelativeUrl('${encodeURI(odataString(serverRelative))}')`;
  }

  private folderUrl(serverRelative: string): string {
    return `${this.web}/_api/web/GetFolderByServerRelativeUrl('${encodeURI(odataString(serverRelative))}')`;
  }

  // ---- index ----------------------------------------------------------------

  public async loadIndex(fresh = false): Promise<ITenantIndex> {
    if (fresh) {
      this.client.clearCache();
    }
    try {
      return normalizeIndex(await this.client.get<unknown>(`${this.fileUrl(`${this.assetsFolder()}/${INDEX_FILE}`)}/$value`, { cacheTtlMs: fresh ? 0 : 60000 }));
    } catch (e) {
      if (e instanceof HttpError && e.status === 404) {
        return emptyIndex();
      }
      if (e instanceof SyntaxError) {
        throw new Error('The tenant storage index could not be read (it is not valid JSON).');
      }
      throw e;
    }
  }

  private async writeFile(name: string, content: string): Promise<void> {
    const url = `${this.folderUrl(this.assetsFolder())}/Files/add(url='${odataString(name)}',overwrite=true)`;
    try {
      await this.client.post(url, content, { raw: true }, this.web);
    } catch (e) {
      if (e instanceof HttpError && e.status === 404) {
        throw new Error('The Site Assets library was not found on this site, so imported snapshots cannot be saved. Add the web part to a site that has Site Assets.');
      }
      throw e;
    }
  }

  private async deleteFile(name: string): Promise<void> {
    try {
      await this.client.post(this.fileUrl(`${this.assetsFolder()}/${name}`), undefined, { method: 'DELETE' }, this.web);
    } catch {
      /* best effort: a leftover snapshot file is harmless */
    }
  }

  private async loadSnapshot(meta: ISnapshotMeta): Promise<ISnapshotData> {
    const d = await this.client.get<ISnapshotData>(`${this.fileUrl(`${this.assetsFolder()}/${meta.file}`)}/$value`, { cacheTtlMs: 300000 });
    if (!d || !Array.isArray(d.u) || !Array.isArray(d.b)) {
      throw new Error(`Snapshot ${meta.file} is damaged.`);
    }
    const titles = (d as unknown as { titles?: { [k: string]: string } }).titles;
    const out: ISnapshotData & { titles?: { [k: string]: string } } = { t: meta.t, u: d.u, b: d.b, q: Array.isArray(d.q) ? d.q : [], s: Array.isArray(d.s) ? d.s : undefined, tp: d.tp, tpl: d.tpl, tm: d.tm, la: d.la, titles };
    return out;
  }

  // ---- finding the CSV -------------------------------------------------------

  public async listLibraries(): Promise<Array<{ title: string; rootFolder: string }>> {
    const j = await this.client.get(`${this.web}/_api/web/lists?$select=Title,Hidden,BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=BaseType eq 1`.replace(/ /g, '%20'));
    return rowsOf<{ Title?: string; Hidden?: boolean; RootFolder?: { ServerRelativeUrl?: string } }>(j)
      .filter((l) => !l.Hidden && l.RootFolder && l.RootFolder.ServerRelativeUrl)
      .map((l) => ({ title: String(l.Title), rootFolder: String((l.RootFolder as { ServerRelativeUrl: string }).ServerRelativeUrl) }));
  }

  public async listFolders(folder: string): Promise<Array<{ name: string; path: string }>> {
    const j = await this.client.get(`${this.folderUrl(folder)}/Folders?$select=Name,ServerRelativeUrl&$top=200`, { cacheTtlMs: 0 });
    return rowsOf<{ Name?: string; ServerRelativeUrl?: string }>(j)
      .filter((f) => f.Name && f.Name !== 'Forms' && f.Name.charAt(0) !== '_')
      .map((f) => ({ name: String(f.Name), path: String(f.ServerRelativeUrl) }));
  }

  public async listCsvFiles(folder: string): Promise<ICsvFile[]> {
    const j = await this.client.get(`${this.folderUrl(folder)}/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Length&$top=500`, { cacheTtlMs: 0 });
    return rowsOf<{ Name?: string; ServerRelativeUrl?: string; TimeLastModified?: string; Length?: string | number }>(j)
      .filter((f) => /\.csv$/i.test(String(f.Name)))
      .map((f) => ({ name: String(f.Name), serverRelativeUrl: String(f.ServerRelativeUrl), modified: toDate(f.TimeLastModified), size: Number(f.Length) || 0 }))
      .sort((a, b) => (b.modified ? b.modified.getTime() : 0) - (a.modified ? a.modified.getTime() : 0));
  }

  public readCsv(serverRelativeUrl: string): Promise<string> {
    return this.client.getText(`${this.fileUrl(serverRelativeUrl)}/$value`);
  }

  public async getFileInfo(serverRelativeUrl: string): Promise<ICsvFile> {
    const j = await this.client.get(`${this.fileUrl(serverRelativeUrl)}?$select=Name,ServerRelativeUrl,TimeLastModified,Length`, { cacheTtlMs: 0 });
    const f = entityOf<{ Name?: string; ServerRelativeUrl?: string; TimeLastModified?: string; Length?: string | number }>(j);
    return { name: String(f.Name), serverRelativeUrl: String(f.ServerRelativeUrl), modified: toDate(f.TimeLastModified), size: Number(f.Length) || 0 };
  }

  // ---- import & analysis -----------------------------------------------------

  public async importCsv(
    text: string,
    mapping: IMapping,
    when: { t: number; fileName: string; folder?: string; modified?: number },
    settings: IGrowthSettings,
    progress?: (step: string) => void
  ): Promise<IImportResult> {
    const say = (s: string): void => {
      if (progress) {
        progress(s);
      }
    };
    say('Reading the CSV…');
    const rows = parseCsv(text);
    const built = buildSnapshot(rows, mapping, this.origin, when.t);
    if (built.data.u.length === 0) {
      throw new Error('No usable site rows were found with this column mapping. Check the URL and storage columns.');
    }

    const index = await this.loadIndex(true);
    // A snapshot within a few hours of an existing one replaces it instead of piling up.
    const near = index.snapshots.filter((s) => Math.abs(s.t - when.t) < MIN_GAP_SEC);
    const file = `${SNAP_PREFIX}${when.t}.json`;
    say('Saving the snapshot…');
    await this.writeFile(file, JSON.stringify({ v: 1, u: built.data.u, b: built.data.b, q: built.data.q, s: built.data.s, tp: built.data.tp, tpl: built.data.tpl, tm: built.data.tm, la: built.data.la, titles: built.titles }));

    const meta: ISnapshotMeta = { t: when.t, file, sites: built.data.u.length, bytes: built.data.b.reduce((s, x) => s + x, 0), source: when.fileName };
    const kept = index.snapshots.filter((s) => near.indexOf(s) < 0);
    const all = kept.concat(meta).sort((a, b) => a.t - b.t);

    say('Comparing with earlier snapshots…');
    const wanted = selectSnapshots(all, 8);
    const loaded: ISnapshotData[] = [];
    for (const m of wanted) {
      loaded.push(m.file === file ? built.data : await this.loadSnapshot(m));
    }
    const nowSec = Math.floor(Date.now() / 1000);
    const summary = analyzeTenant(loaded, settings, built.titles, this.origin, nowSec);
    const isNewest = all[all.length - 1].file === file;
    const overview = isNewest ? summarizeInventory(built.data, built.titles, this.origin, nowSec, (t) => siteType(t, false)) : undefined;

    say('Saving the results…');
    // Re-read before writing so a concurrent import by another admin is not lost.
    const latest = await this.loadIndex(true);
    const merged = latest.snapshots.filter((s) => near.every((n) => n.file !== s.file) && s.file !== file).concat(meta).sort((a, b) => a.t - b.t);
    const dropped = merged.length > MAX_SNAPSHOTS_KEPT ? merged.slice(0, merged.length - MAX_SNAPSHOTS_KEPT) : [];
    const next: ITenantIndex = {
      v: 1,
      snapshots: merged.slice(dropped.length),
      summary: merged[merged.length - 1].file === file ? summary : latest.summary,
      overview: merged[merged.length - 1].file === file && overview ? overview : latest.overview,
      capacityTB: latest.capacityTB,
      countDeleted: latest.countDeleted,
      mapping,
      folder: when.folder || latest.folder,
      lastFile: when.modified ? { name: when.fileName, modified: when.modified } : latest.lastFile
    };
    await this.writeFile(INDEX_FILE, JSON.stringify(next));

    for (const old of near.map((n) => n.file).filter((f) => f !== file).concat(dropped.map((d) => d.file))) {
      await this.deleteFile(old);
    }
    return { stats: built.stats, summary, replaced: near.length > 0 };
  }

  /** Tenant capacity and whether deleted sites count as used; shared by every admin. */
  public async saveCapacity(capacityTB: number | undefined, countDeleted: boolean): Promise<void> {
    const latest = await this.loadIndex(true);
    latest.capacityTB = capacityTB && capacityTB > 0 ? capacityTB : undefined;
    latest.countDeleted = countDeleted;
    await this.writeFile(INDEX_FILE, JSON.stringify(latest));
  }

  /** The newest snapshot in full (one request, ~0.5–1 MB for 17,000 sites) for the all-sites table. */
  public async loadLatestSnapshot(): Promise<(ISnapshotData & { titles?: { [k: string]: string } }) | undefined> {
    const index = await this.loadIndex();
    const meta = index.snapshots[index.snapshots.length - 1];
    return meta ? this.loadSnapshot(meta) : undefined;
  }

  /** Re-runs the analysis over the stored snapshots with the current thresholds (no CSV needed). */
  public async recalculate(settings: IGrowthSettings): Promise<ITenantSummary | undefined> {
    const index = await this.loadIndex(true);
    if (index.snapshots.length === 0) {
      return undefined;
    }
    const loaded: ISnapshotData[] = [];
    for (const m of selectSnapshots(index.snapshots, 8)) {
      loaded.push(await this.loadSnapshot(m));
    }
    const summary = analyzeTenant(loaded, settings, {}, this.origin, Math.floor(Date.now() / 1000));
    // Keep titles known from the previous summary.
    const titles: { [url: string]: string } = {};
    ((index.summary && index.summary.anomalies.concat(index.summary.largest, index.summary.growers)) || []).forEach((r) => (titles[r.url] = r.title));
    const fix = (r: { url: string; title: string }): void => {
      r.title = titles[r.url] || r.title;
    };
    summary.anomalies.concat(summary.largest, summary.growers).forEach(fix);
    const newest = loaded[loaded.length - 1];
    const snapTitles = ((newest as unknown) as { titles?: { [k: string]: string } }).titles || {};
    const latest = await this.loadIndex(true);
    latest.summary = summary;
    latest.overview = summarizeInventory(newest, snapTitles, this.origin, Math.floor(Date.now() / 1000), (t) => siteType(t, false));
    await this.writeFile(INDEX_FILE, JSON.stringify(latest));
    return summary;
  }
}
