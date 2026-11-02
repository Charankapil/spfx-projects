import { IPbiActivity, IPowerBiData } from '../models';
import { IPbiReportRow } from './columns';

/** Joins activity usage onto reports. Usage is undefined when the activity log is unavailable. */
export function pbiReportRows(data: IPowerBiData, activity: IPbiActivity | undefined): IPbiReportRow[] {
  const ds: Record<string, string> = {};
  data.datasets.forEach((d) => {
    ds[d.id] = d.name;
  });
  const usage: Record<string, { views: number; viewers: number; lastViewed?: string }> = {};
  if (activity) activity.reports.forEach((r) => (usage[r.id] = r));
  return data.reports.map((r) => {
    const u = usage[r.id];
    return {
      ...r,
      datasetName: r.datasetId ? ds[r.datasetId] : undefined,
      views: activity ? (u ? u.views : 0) : undefined,
      viewers: activity ? (u ? u.viewers : 0) : undefined,
      lastViewed: u ? u.lastViewed : undefined
    };
  });
}

/** Reports in shared workspaces with no views in the activity period. */
export function unusedReports(rows: IPbiReportRow[], data: IPowerBiData): IPbiReportRow[] {
  const personal: Record<string, boolean> = {};
  data.workspaces.forEach((w) => {
    if (w.type === 'Personal') personal[w.id] = true;
  });
  return rows.filter((r) => r.views === 0 && !personal[r.workspaceId]);
}
