import { cacheGet, cacheKey, cacheRemove, cacheSet } from './cache';
import { parseDate } from './format';
import { field, IListDefinition, isMissingList, ListService } from './lists';
import { SpClient } from './SpClient';

export const ALERT_BANNER_COMPONENT_ID = 'b2be5f7c-564a-4ad3-ba1a-16ab1b12a70e';

export const SEVERITIES = ['Info', 'Warning', 'Critical'];

export const ALERTS: IListDefinition = {
  urlName: 'ToolkitAlerts',
  title: 'Site Alerts',
  description: 'Alerts shown at the top of every page by the Site Toolkit alert banner. The Title column is the headline.',
  fields: [
    field.note('Message', 'Message'),
    field.choice('Severity', 'Severity', SEVERITIES),
    field.dateTime('StartDate', 'Show From'),
    field.dateTime('EndDate', 'Show Until'),
    field.text('LinkUrl', 'Link')
  ]
};

export interface IAlert {
  Id: number;
  Title: string;
  Message?: string;
  Severity?: string;
  StartDate?: string;
  EndDate?: string;
  LinkUrl?: string;
  Modified?: string;
}

/**
 * The banner runs on every page for every visitor, so it asks SharePoint at
 * most once per person per 15 minutes; all other page loads read the
 * browser's copy. A missing list is remembered the same way.
 */
export const ALERT_CACHE_MINUTES = 15;

interface ICachedAlerts {
  /** False when the alerts list does not exist on this site collection. */
  listExists: boolean;
  alerts: IAlert[];
}

function alertsKey(sp: SpClient): string {
  return cacheKey(sp.siteUrl, 'alerts');
}

export function clearAlertCache(sp: SpClient): void {
  cacheRemove('local', alertsKey(sp));
}

export type AlertStatus = 'active' | 'scheduled' | 'expired';

export function alertStatus(alert: IAlert, now: Date = new Date()): AlertStatus {
  const start = parseDate(alert.StartDate);
  const end = parseDate(alert.EndDate);
  if (start && start.getTime() > now.getTime()) {
    return 'scheduled';
  }
  if (end && end.getTime() < now.getTime()) {
    return 'expired';
  }
  return 'active';
}

/** Requests already on their way, so callers that ask at the same moment share one request. */
const inFlight: Record<string, Promise<IAlert[] | undefined> | undefined> = {};

/** All alerts on the site collection, newest first; undefined when the list does not exist. */
export function loadAlerts(sp: SpClient, force = false): Promise<IAlert[] | undefined> {
  const key = alertsKey(sp);
  const hit = force ? undefined : cacheGet<ICachedAlerts>('local', key);
  if (hit) {
    return Promise.resolve(hit.listExists ? hit.alerts : undefined);
  }
  const running = inFlight[key];
  if (running) {
    return running;
  }
  const request = fetchAlerts(sp, key);
  inFlight[key] = request;
  const forget = (): void => {
    inFlight[key] = undefined;
  };
  request.then(forget, forget);
  return request;
}

async function fetchAlerts(sp: SpClient, key: string): Promise<IAlert[] | undefined> {
  const lists = ListService.forSiteRoot(sp);
  try {
    const result = await lists.getItems<IAlert>(
      ALERTS.urlName,
      '$select=Id,Title,Message,Severity,StartDate,EndDate,LinkUrl,Modified&$orderby=Id desc&$top=50',
      50
    );
    cacheSet<ICachedAlerts>('local', key, { listExists: true, alerts: result.items }, ALERT_CACHE_MINUTES);
    return result.items;
  } catch (err) {
    if (isMissingList(err)) {
      cacheSet<ICachedAlerts>('local', key, { listExists: false, alerts: [] }, ALERT_CACHE_MINUTES);
      return undefined;
    }
    throw err;
  }
}

function dismissedKey(sp: SpClient, alert: IAlert): string {
  // Editing an alert changes Modified, so an updated alert shows again.
  return cacheKey(sp.siteUrl, `alert-dismissed:${alert.Id}:${alert.Modified || ''}`);
}

export function isDismissed(sp: SpClient, alert: IAlert): boolean {
  return cacheGet<boolean>('local', dismissedKey(sp, alert)) === true;
}

export function dismiss(sp: SpClient, alert: IAlert): void {
  cacheSet('local', dismissedKey(sp, alert), true, 60 * 24 * 60);
}
