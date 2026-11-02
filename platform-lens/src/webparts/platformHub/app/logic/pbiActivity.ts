import { IPbiActivity, IPbiEvent, IPbiItemUsage, IPbiUserUsage } from '../models';

/** Raw activity event (subset of the Power BI admin activity log schema). */
export interface IRawPbiEvent {
  CreationTime?: string;
  UserId?: string;
  Activity?: string;
  Operation?: string;
  ItemName?: string;
  WorkSpaceName?: string;
  ReportId?: string;
  ReportName?: string;
  DashboardId?: string;
  DashboardName?: string;
  DatasetName?: string;
  ConsumptionMethod?: string;
  ClientIP?: string;
}

const VIEW = /^(ViewReport|ViewDashboard)$/i;
const EXPORT = /^(Export|AnalyzeInExcel|AnalyzedByExternalApplication|DownloadReport|ExportArtifact|ExportTile)/i;
const SHARE = /^(Share|AddLinkToArtifact|CreateArtifactShareLink|UpdateArtifactShareLink)/i;
const PUBLISH_WEB = /PublishToWeb|CreateEmbedCode/i;

export const SENSITIVE_CAP = 5000;
export const RECENT_CAP = 2000;

export function isExport(activity: string): boolean {
  return EXPORT.test(activity);
}

export function isShare(activity: string): boolean {
  return SHARE.test(activity);
}

export function isPublishToWeb(activity: string): boolean {
  return PUBLISH_WEB.test(activity);
}

/**
 * Aggregates activity events as they stream in, so tenants with millions of
 * events per month never hold the full log in browser memory.
 */
export class PbiActivityAggregator {
  private _total = 0;
  private _views = 0;
  private _exports = 0;
  private _shares = 0;
  private _publish = 0;
  private _users: Record<string, IPbiUserUsage> = {};
  private _daily: Record<string, { users: Record<string, boolean>; views: number; events: number }> = {};
  private _activity: Record<string, number> = {};
  private _consumption: Record<string, number> = {};
  private _reports: Record<string, IPbiItemUsage & { v: Record<string, boolean> }> = {};
  private _sensitive: IPbiEvent[] = [];
  private _recent: IPbiEvent[] = [];

  public add(e: IRawPbiEvent): void {
    const activity = e.Activity || e.Operation || 'Unknown';
    const user = (e.UserId || 'unknown').toLowerCase();
    const time = e.CreationTime || '';
    const day = time.slice(0, 10);
    this._total++;
    const d = this._daily[day] || (this._daily[day] = { users: {}, views: 0, events: 0 });
    d.events++;
    d.users[user] = true;
    this._activity[activity] = (this._activity[activity] || 0) + 1;
    const u = this._users[user] || (this._users[user] = { user, views: 0, events: 0, exports: 0 });
    u.events++;
    if (!u.lastActive || time > u.lastActive) u.lastActive = time;
    const event: IPbiEvent = {
      time, user, activity,
      item: e.ItemName || e.ReportName || e.DashboardName || e.DatasetName || '',
      workspace: e.WorkSpaceName || '',
      consumption: e.ConsumptionMethod,
      clientIp: e.ClientIP
    };
    if (VIEW.test(activity)) {
      this._views++;
      d.views++;
      u.views++;
      const method = e.ConsumptionMethod || 'Unknown';
      this._consumption[method] = (this._consumption[method] || 0) + 1;
      const id = e.ReportId || e.DashboardId;
      if (id) {
        const r = this._reports[id] || (this._reports[id] = { id, name: e.ReportName || e.DashboardName || e.ItemName || id, workspace: e.WorkSpaceName || '', views: 0, viewers: 0, v: {} });
        r.views++;
        r.v[user] = true;
        if (!r.lastViewed || time > r.lastViewed) r.lastViewed = time;
      }
    }
    const exp = isExport(activity);
    const shr = isShare(activity);
    const pub = isPublishToWeb(activity);
    if (exp) {
      this._exports++;
      u.exports++;
    }
    if (shr) this._shares++;
    if (pub) this._publish++;
    if ((exp || shr || pub) && this._sensitive.length < SENSITIVE_CAP) this._sensitive.push(event);
    this._recent.push(event);
    if (this._recent.length > RECENT_CAP * 2) this._trimRecent();
  }

  private _trimRecent(): void {
    this._recent.sort((a, b) => (a.time < b.time ? 1 : a.time > b.time ? -1 : 0));
    this._recent.length = Math.min(this._recent.length, RECENT_CAP);
  }

  public result(days: number, from: string, to: string, failedDays: string[]): IPbiActivity {
    this._trimRecent();
    const toCounts = (m: Record<string, number>): { label: string; value: number }[] =>
      Object.keys(m).map((label) => ({ label, value: m[label] })).sort((a, b) => b.value - a.value);
    return {
      days, from, to, failedDays,
      totalEvents: this._total,
      activeUsers: Object.keys(this._users).length,
      views: this._views,
      exports: this._exports,
      shares: this._shares,
      publishToWeb: this._publish,
      daily: Object.keys(this._daily).filter(Boolean).sort().map((date) => ({
        date, activeUsers: Object.keys(this._daily[date].users).length, views: this._daily[date].views, events: this._daily[date].events
      })),
      byActivity: toCounts(this._activity),
      byConsumption: toCounts(this._consumption),
      reports: Object.keys(this._reports).map((k) => {
        const { v, ...r } = this._reports[k];
        return { ...r, viewers: Object.keys(v).length };
      }).sort((a, b) => b.views - a.views),
      users: Object.keys(this._users).map((k) => this._users[k]).sort((a, b) => b.events - a.events),
      sensitive: this._sensitive.slice().sort((a, b) => (a.time < b.time ? 1 : -1)),
      recent: this._recent.slice()
    };
  }
}

/** UTC day windows for the last `days` days, ending now (the activity API accepts one UTC day per call). */
export function activityWindows(days: number, now: number = Date.now()): { start: string; end: string; day: string }[] {
  const out: { start: string; end: string; day: string }[] = [];
  const end = new Date(now - 5 * 60000);
  for (let i = days - 1; i >= 0; i--) {
    const d = new Date(Date.UTC(end.getUTCFullYear(), end.getUTCMonth(), end.getUTCDate() - i));
    const day = d.toISOString().slice(0, 10);
    const last = i === 0 ? end.toISOString() : `${day}T23:59:59.999Z`;
    out.push({ day, start: `${day}T00:00:00.000Z`, end: last });
  }
  return out;
}
