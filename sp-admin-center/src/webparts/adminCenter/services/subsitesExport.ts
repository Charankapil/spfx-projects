import { ISubsiteNode } from '../models';
import { ISiteAccess, peopleText } from './siteGroupsExport';

/** What was read for one web of the tree. */
export interface ISubsiteDetail {
  web: ISubsiteNode;
  size?: { bytes: number; files: number };
  /** Owners and members (only read for webs with their own permissions). */
  access?: ISiteAccess;
}

export const SUBSITE_HEADERS = [
  'Level',
  'Title',
  'URL',
  'Parent URL',
  'Template',
  'Created',
  'Last modified',
  'Size (bytes)',
  'Size (GB)',
  'Files',
  'Permissions',
  'Owners group',
  'Owners',
  'Members group',
  'Members',
  'Visitors group',
  'Visitors',
  'Note'
];

function day(d?: Date): string {
  return d ? d.toISOString().substring(0, 10) : '';
}

/** One CSV row per web; inherited webs say so instead of repeating the parent's people. */
export function subsiteRows(details: ISubsiteDetail[]): unknown[][] {
  return details.map((d) => {
    const w = d.web;
    const a = d.access;
    const unique = w.hasUniquePermissions;
    const note = w.depth === 0 ? 'Top-level site (size excludes subsites)' : !d.size ? 'Size not available' : a && a.error ? a.error : '';
    return [
      w.depth,
      w.title,
      w.url,
      w.parentUrl || '',
      w.template,
      day(w.created),
      day(w.lastModified),
      d.size ? d.size.bytes : '',
      d.size ? Math.round((d.size.bytes / 1073741824) * 100) / 100 : '',
      d.size ? d.size.files : '',
      w.depth === 0 ? 'Top-level site' : unique ? 'Unique' : 'Inherits from parent',
      unique && a && !a.error ? a.ownerGroup || '' : '',
      unique && a && !a.error ? peopleText(a.owners) : '',
      unique && a && !a.error ? a.memberGroup || '' : '',
      unique && a && !a.error ? peopleText(a.members) : '',
      unique && a && !a.error ? a.visitorGroup || '' : '',
      unique && a && !a.error ? peopleText(a.visitors) : '',
      note
    ];
  });
}
