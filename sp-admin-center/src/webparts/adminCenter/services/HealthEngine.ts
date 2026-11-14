import { IFinding, IGroupInfo, IListInfo, ISiteInfo, ISubWeb, IUserInfo, IWebInfo } from '../models';

export interface IHealthInput {
  site?: ISiteInfo;
  web?: IWebInfo;
  users?: IUserInfo[];
  groups?: IGroupInfo[];
  lists?: IListInfo[];
  subwebs?: ISubWeb[];
  recycleCount?: number;
  recycleCapped?: boolean;
  now?: Date;
}

export interface IHealthResult {
  score: number;
  findings: IFinding[];
}

export const LIST_VIEW_THRESHOLD = 5000;
const DAY = 86400000;

/**
 * Pure rules over data the other views already loaded (so the health check
 * costs no extra requests when it is opened after them, and only the cached,
 * cheap calls when opened first). Each rule returns either a finding with a
 * score penalty or a "good" confirmation.
 */
export function evaluateHealth(input: IHealthInput): IHealthResult {
  const f: IFinding[] = [];
  const now = (input.now || new Date()).getTime();

  const { site, web, users, lists } = input;

  if (site) {
    const pct = Math.round(site.storageFraction * 100);
    if (site.storageFraction >= 0.9) {
      f.push({ id: 'storage', severity: 'critical', title: `Storage is ${pct}% full`, detail: 'The site collection is close to its quota. Review large files and the recycle bin, or ask for more storage.', view: 'storage', penalty: 20 });
    } else if (site.storageFraction >= 0.75) {
      f.push({ id: 'storage', severity: 'warning', title: `Storage is ${pct}% full`, detail: 'Plan clean-up or a quota increase before it fills up.', view: 'storage', penalty: 8 });
    } else {
      f.push({ id: 'storage', severity: 'good', title: `Storage healthy (${pct}% used)`, detail: '', penalty: 0 });
    }
    if (site.readOnly) {
      f.push({ id: 'readonly', severity: 'info', title: 'Site collection is read-only (locked)', detail: 'Users cannot add or change content.', penalty: 0 });
    }
  }

  if (users) {
    const admins = users.filter((u) => u.isSiteAdmin && u.kind === 'Member');
    if (admins.length === 0) {
      f.push({ id: 'admins', severity: 'warning', title: 'No individual site collection admins found', detail: 'Admin access may rely on a group only. Make sure someone accountable owns this site.', view: 'people', penalty: 8 });
    } else if (admins.length === 1) {
      f.push({ id: 'admins', severity: 'warning', title: 'Only one site collection admin', detail: `${admins[0].title} is the only admin. Add a second so the site is never orphaned.`, view: 'people', penalty: 6 });
    } else if (admins.length > 5) {
      f.push({ id: 'admins', severity: 'warning', title: `${admins.length} site collection admins`, detail: 'Admin access is broad. Remove anyone who no longer needs it.', view: 'people', penalty: 4 });
    } else {
      f.push({ id: 'admins', severity: 'good', title: `${admins.length} site collection admins`, detail: '', penalty: 0 });
    }

    const guests = users.filter((u) => u.kind === 'Guest');
    if (guests.length > 100) {
      f.push({ id: 'guests', severity: 'warning', title: `${guests.length} external (guest) users`, detail: 'Review whether all of them still need access.', view: 'people', penalty: 4 });
    } else if (guests.length > 0) {
      f.push({ id: 'guests', severity: 'info', title: `${guests.length} external (guest) user${guests.length === 1 ? '' : 's'}`, detail: 'Worth a periodic review.', view: 'people', penalty: 0 });
    } else {
      f.push({ id: 'guests', severity: 'good', title: 'No external users on this site', detail: '', penalty: 0 });
    }

    const orgWide = users.filter((u) => u.kind === 'OrgWide');
    if (orgWide.length > 0) {
      f.push({ id: 'orgwide', severity: 'warning', title: 'Organisation-wide group is on this site', detail: `"${orgWide.map((u) => u.title).join('", "')}" has been added here. If it was granted access, everyone in the organisation can open this content.`, view: 'people', penalty: 8 });
    } else {
      f.push({ id: 'orgwide', severity: 'good', title: 'No organisation-wide sharing groups', detail: '', penalty: 0 });
    }
  }

  if (lists) {
    const visible = lists.filter((l) => !l.hidden && !l.isSystem);
    const huge = visible.filter((l) => l.itemCount >= LIST_VIEW_THRESHOLD);
    const near = visible.filter((l) => l.itemCount >= 4000 && l.itemCount < LIST_VIEW_THRESHOLD);
    if (huge.length > 0) {
      f.push({ id: 'threshold', severity: 'warning', title: `${huge.length} list${huge.length === 1 ? '' : 's'} over the ${LIST_VIEW_THRESHOLD.toLocaleString()}-item view threshold`, detail: 'Unindexed views and operations can fail on these: ' + names(huge) + '. Add indexes or filter the views.', view: 'content', penalty: Math.min(8, 3 * huge.length) });
    } else if (near.length > 0) {
      f.push({ id: 'threshold', severity: 'info', title: `${near.length} list${near.length === 1 ? '' : 's'} approaching the view threshold`, detail: names(near), view: 'content', penalty: 0 });
    } else {
      f.push({ id: 'threshold', severity: 'good', title: 'No lists near the view threshold', detail: '', penalty: 0 });
    }

    const noVersions = visible.filter((l) => l.kind === 'Library' && !l.versioning);
    if (noVersions.length > 0) {
      f.push({ id: 'versioning', severity: 'warning', title: `${noVersions.length} librar${noVersions.length === 1 ? 'y has' : 'ies have'} version history off`, detail: 'Edits and deletions cannot be rolled back: ' + names(noVersions) + '.', view: 'content', penalty: Math.min(8, 2 + noVersions.length) });
    } else if (visible.some((l) => l.kind === 'Library')) {
      f.push({ id: 'versioning', severity: 'good', title: 'All libraries keep version history', detail: '', penalty: 0 });
    }

    const noCrawl = visible.filter((l) => l.noCrawl);
    if (noCrawl.length > 0) {
      f.push({ id: 'nocrawl', severity: 'info', title: `${noCrawl.length} list${noCrawl.length === 1 ? ' is' : 's are'} hidden from search`, detail: names(noCrawl) + '. Their content will not appear in search results or in the Storage insights.', view: 'content', penalty: 0 });
    }

    const unique = visible.filter((l) => l.uniquePermissions);
    if (unique.length > 20) {
      f.push({ id: 'unique', severity: 'warning', title: `${unique.length} lists with their own permissions`, detail: 'Many broken-inheritance lists are hard to audit. Prefer fewer, group-based permissions.', view: 'people', penalty: 4 });
    } else if (unique.length > 0) {
      f.push({ id: 'unique', severity: 'info', title: `${unique.length} list${unique.length === 1 ? '' : 's'} with unique permissions`, detail: names(unique), view: 'content', penalty: 0 });
    }
  }

  if (input.subwebs) {
    const unique = input.subwebs.filter((s) => s.hasUniquePermissions);
    if (unique.length > 0) {
      f.push({ id: 'subwebs', severity: 'info', title: `${unique.length} subsite${unique.length === 1 ? '' : 's'} with unique permissions`, detail: unique.map((s) => s.title).slice(0, 6).join(', '), view: 'overview', penalty: 0 });
    }
  }

  if (input.recycleCount !== undefined && input.recycleCount > 0) {
    const n = input.recycleCount;
    const label = input.recycleCapped ? `${n}+` : String(n);
    if (n >= 1000) {
      f.push({ id: 'recycle', severity: 'info', title: `${label} items in the recycle bin`, detail: 'Deleted content still counts toward storage until the bin is emptied.', view: 'recycle', penalty: 0 });
    }
  }

  if (web && web.lastModified) {
    const age = (now - web.lastModified.getTime()) / DAY;
    if (age > 365) {
      f.push({ id: 'stale', severity: 'warning', title: `No content changes for ${Math.floor(age / 30)} months`, detail: 'This site looks inactive. Confirm the owner still needs it, or archive it.', view: 'overview', penalty: 6 });
    } else if (age > 180) {
      f.push({ id: 'stale', severity: 'info', title: `Last content change ${Math.floor(age / 30)} months ago`, detail: '', view: 'overview', penalty: 0 });
    } else {
      f.push({ id: 'stale', severity: 'good', title: 'Site is actively used', detail: '', penalty: 0 });
    }
  }

  const penalty = f.reduce((s, x) => s + x.penalty, 0);
  const order = { critical: 0, warning: 1, info: 2, good: 3 };
  f.sort((a, b) => order[a.severity] - order[b.severity]);
  return { score: Math.max(0, 100 - penalty), findings: f };
}

function names(items: Array<{ title: string }>): string {
  const shown = items.slice(0, 5).map((i) => i.title);
  return shown.join(', ') + (items.length > 5 ? ` and ${items.length - 5} more` : '');
}
