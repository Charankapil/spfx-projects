import { entityOf, HttpError, IBatchOp, odataString, rowsOf, SPClient, trimSlash } from '../core/SPClient';
import { ISiteRow } from '../models';
import { IGroupSummary } from './GraphApi';
import { ISiteAccess, peopleText, PEOPLE_SEPARATOR, splitList } from './siteGroupsExport';
import { siteType } from './siteTypes';

/**
 * Saves the "sites with owners and members" report into a SharePoint list: one item per
 * site. The list and its columns are created when missing; items are added with SharePoint
 * $batch requests (up to 50 items or about 1 MB per request, one request at a time), so
 * 17,000 sites take a few hundred requests instead of 17,000.
 */

export interface IListColumn {
  name: string;
  display: string;
  type: 'Text' | 'Note' | 'Number' | 'DateTime';
}

export const LIST_COLUMNS: IListColumn[] = [
  { name: 'SiteUrl', display: 'Site URL', type: 'Text' },
  { name: 'SiteType', display: 'Site type', type: 'Text' },
  { name: 'MembershipFrom', display: 'Membership from', type: 'Text' },
  { name: 'M365Group', display: 'Microsoft 365 group', type: 'Text' },
  { name: 'GroupName', display: 'Group name', type: 'Text' },
  { name: 'GroupEmail', display: 'Group e-mail', type: 'Text' },
  { name: 'Privacy', display: 'Privacy', type: 'Text' },
  { name: 'Teams', display: 'Teams', type: 'Text' },
  { name: 'OwnersCount', display: 'Owners (count)', type: 'Number' },
  { name: 'Owners', display: 'Owners', type: 'Note' },
  { name: 'MembersCount', display: 'Members (count)', type: 'Number' },
  { name: 'Members', display: 'Members', type: 'Note' },
  { name: 'VisitorsCount', display: 'Visitors (count)', type: 'Number' },
  { name: 'Visitors', display: 'Visitors', type: 'Note' },
  { name: 'SiteAdmins', display: 'Site admins', type: 'Note' },
  { name: 'ExportNote', display: 'Note', type: 'Note' },
  { name: 'ExportedOn', display: 'Exported on', type: 'DateTime' }
];

/** Multiple lines of text hold up to 63,999 characters in a SharePoint list. */
export const NOTE_LIMIT = 63000;
const TEXT_LIMIT = 255;
const BATCH_ITEMS = 50;
const BATCH_BYTES = 1000000;

function xmlAttr(v: string): string {
  return v.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

export function fieldXml(c: IListColumn): string {
  const common = `DisplayName="${xmlAttr(c.display)}" Name="${c.name}" StaticName="${c.name}"`;
  switch (c.type) {
    case 'Note':
      return `<Field Type="Note" ${common} NumLines="6" RichText="FALSE" UnlimitedLengthInDocumentLibrary="TRUE" />`;
    case 'Number':
      return `<Field Type="Number" ${common} Decimals="0" />`;
    case 'DateTime':
      return `<Field Type="DateTime" ${common} Format="DateTime" />`;
    default:
      return `<Field Type="Text" ${common} MaxLength="255" />`;
  }
}

/** Fits a people list in a list column, cut between addresses with a count of what was left out. */
export function noteValue(text: string, limit = NOTE_LIMIT): string {
  if (text.length <= limit) {
    return text;
  }
  const first = splitList(text, limit - 80)[0];
  const kept = first.split(PEOPLE_SEPARATOR).length;
  const total = text.split(PEOPLE_SEPARATOR).length;
  return `${first}${PEOPLE_SEPARATOR}… (${(total - kept).toLocaleString()} more; too many for one list column - use the CSV export)`;
}

function textValue(v: string): string {
  return v.length > TEXT_LIMIT ? v.substring(0, TEXT_LIMIT - 1) + '…' : v;
}

/** One list item (field values by internal name) per site. Pure. */
export function listItemsFor(
  sites: ISiteRow[],
  groups: { [id: string]: IGroupSummary },
  accessOf: (s: ISiteRow) => ISiteAccess | undefined,
  exportedOn: Date
): Array<{ [field: string]: unknown }> {
  return sites.map((s) => {
    const g = s.groupId ? groups[s.groupId] : undefined;
    const a = accessOf(s);
    const ok = !!a && !a.error;
    const notes: string[] = [];
    if (!s.groupId && s.relatedGroupId) {
      notes.push('Teams channel site: membership is the channel’s');
    }
    if (!a) {
      notes.push('Not read (export stopped early)');
    } else if (a.error) {
      notes.push(a.error);
    } else if (a.truncated) {
      notes.push('Very large group: list truncated');
    }
    const item: { [field: string]: unknown } = {
      Title: textValue(s.title || s.url),
      SiteUrl: textValue(s.url),
      SiteType: textValue(siteType(s.template, s.groupConnected)),
      MembershipFrom: a ? (a.source === 'm365' ? 'Microsoft 365 group' : 'SharePoint groups') : '',
      M365Group: s.groupId ? 'Yes' : 'No',
      GroupName: textValue(g ? g.displayName : ''),
      GroupEmail: textValue(g ? g.mail : ''),
      Privacy: g ? g.visibility : '',
      Teams: g ? (g.isTeam ? 'Yes' : 'No') : '',
      Owners: a ? noteValue(peopleText(a.owners)) : '',
      Members: a ? noteValue(peopleText(a.members)) : '',
      Visitors: a ? noteValue(peopleText(a.visitors)) : '',
      SiteAdmins: a && a.admins ? noteValue(peopleText(a.admins)) : '',
      ExportNote: notes.join('. '),
      ExportedOn: exportedOn.toISOString()
    };
    if (ok && a) {
      item.OwnersCount = a.owners.length;
      item.MembersCount = a.members.length;
      if (a.source === 'sharepoint') {
        item.VisitorsCount = a.visitors.length;
      }
    }
    return item;
  });
}

export class ListExport {
  constructor(private client: SPClient) {}

  /** Finds the list by title or creates it, then adds any missing column. */
  public async ensureList(webUrl: string, title: string): Promise<{ id: string; url: string; created: boolean }> {
    const web = trimSlash(webUrl);
    let created = false;
    let list: { Id?: string; RootFolder?: { ServerRelativeUrl?: string } };
    try {
      list = entityOf(await this.client.get(`${web}/_api/web/lists/getbytitle('${encodeURIComponent(odataString(title))}')?$select=Id,RootFolder/ServerRelativeUrl&$expand=RootFolder`, { cacheTtlMs: 0 }));
    } catch (e) {
      if (!(e instanceof HttpError) || e.status !== 404) {
        throw e;
      }
      list = entityOf(
        await this.client.post(`${web}/_api/web/lists`, { Title: title, BaseTemplate: 100, Description: 'Site collections with their owners and members, exported by SharePoint Admin Center.' }, {}, web)
      );
      created = true;
      list = entityOf(await this.client.get(`${web}/_api/web/lists(guid'${String(list.Id)}')?$select=Id,RootFolder/ServerRelativeUrl&$expand=RootFolder`, { cacheTtlMs: 0 }));
    }
    const id = String(list.Id);
    const existing = rowsOf<{ InternalName?: string }>(await this.client.get(`${web}/_api/web/lists(guid'${id}')/fields?$select=InternalName`, { cacheTtlMs: 0 })).map((f) =>
      String(f.InternalName)
    );
    for (const c of LIST_COLUMNS) {
      if (existing.indexOf(c.name) < 0) {
        // 25 = add to the default content type and the default view, keep this internal name.
        await this.client.post(`${web}/_api/web/lists(guid'${id}')/fields/CreateFieldAsXml`, { parameters: { SchemaXml: fieldXml(c), Options: 25 } }, {}, web);
      }
    }
    const origin = (/^(https:\/\/[^/]+)/i.exec(web) || [''])[1];
    return { id, url: origin + (list.RootFolder && list.RootFolder.ServerRelativeUrl ? list.RootFolder.ServerRelativeUrl : ''), created };
  }

  /**
   * Adds items in $batch requests. Throttled operations (429 / 503) are retried once after
   * a pause; anything else that fails is reported by site URL.
   */
  public async addItems(
    webUrl: string,
    listId: string,
    items: Array<{ [field: string]: unknown }>,
    onProgress?: (done: number) => void,
    shouldStop?: () => boolean
  ): Promise<{ added: number; failed: string[] }> {
    const web = trimSlash(webUrl);
    const target = `${web}/_api/web/lists(guid'${listId}')/items`;
    let added = 0;
    const failed: string[] = [];
    let i = 0;
    while (i < items.length) {
      if (shouldStop && shouldStop()) {
        break;
      }
      // Up to 50 items, and about 1 MB, per request.
      const chunk: Array<{ [field: string]: unknown }> = [];
      let bytes = 0;
      while (i < items.length && chunk.length < BATCH_ITEMS) {
        const size = JSON.stringify(items[i]).length;
        if (chunk.length && bytes + size > BATCH_BYTES) {
          break;
        }
        chunk.push(items[i]);
        bytes += size;
        i++;
      }
      let pending = chunk;
      for (let round = 0; round < 2 && pending.length; round++) {
        const ops: IBatchOp[] = pending.map((body) => ({ method: 'POST', url: target, body }));
        const statuses = await this.client.batch(web, ops);
        if (statuses.length !== ops.length) {
          throw new Error(`SharePoint answered ${statuses.length} of ${ops.length} operations in a batch; stopped so no item is added twice. ${added.toLocaleString()} items were added.`);
        }
        const retry: Array<{ [field: string]: unknown }> = [];
        statuses.forEach((st, k) => {
          if (st >= 200 && st < 300) {
            added++;
          } else if ((st === 429 || st === 503) && round === 0) {
            retry.push(pending[k]);
          } else {
            failed.push(`${String(pending[k].SiteUrl)} (${st})`);
          }
        });
        pending = retry;
        if (pending.length) {
          await new Promise((r) => setTimeout(r, 5000));
        }
      }
      if (onProgress) {
        onProgress(added + failed.length);
      }
    }
    return { added, failed };
  }
}
