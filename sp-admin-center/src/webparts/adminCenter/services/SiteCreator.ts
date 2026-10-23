import { entityOf, HttpError, odataString, rowsOf, SPClient, trimSlash } from '../core/SPClient';

/**
 * Site collection creation as the signed-in user, through SharePoint's own
 * site-creation endpoints (the same ones the "+ Create site" panel uses):
 *  - SPSiteManager/create: communication sites (Topic, Showcase, Blank) and team
 *    sites without a Microsoft 365 group, plus your organisation's site templates;
 *  - GroupSiteManager/CreateGroupEx: team sites with a Microsoft 365 group.
 * Classic templates (STS#0, BLOG#0, ...) are only created by the tenant admin API,
 * which a web part cannot reach, so for those (and as an option for every
 * template) the tool gives the exact PowerShell command to run.
 * One request creates the site; status is checked a few times, spaced out.
 */

export type CreationRoute = 'spsite' | 'group' | 'powershell';

export interface ISiteTemplate {
  key: string;
  title: string;
  description: string;
  group: 'Modern' | 'Your organisation' | 'Classic';
  webTemplate: string;
  route: CreationRoute;
  /** Built-in communication site design (Topic / Showcase / Blank). */
  siteDesignId?: string;
  /** A custom site template (site design) from your organisation. */
  customDesignId?: string;
  icon: string;
}

const EMPTY = '00000000-0000-0000-0000-000000000000';

export const BUILT_IN_TEMPLATES: ISiteTemplate[] = [
  { key: 'comm-topic', title: 'Communication site: Topic', description: 'Share news, events and content with a wide audience.', group: 'Modern', webTemplate: 'SITEPAGEPUBLISHING#0', route: 'spsite', siteDesignId: '96c933ac-3698-44c7-9f4a-5fd17d71af9e', icon: 'Megaphone' },
  { key: 'comm-showcase', title: 'Communication site: Showcase', description: 'Showcase a product, team or event with photos and images.', group: 'Modern', webTemplate: 'SITEPAGEPUBLISHING#0', route: 'spsite', siteDesignId: '6142d2a0-63a5-4ba0-aede-d9fefca2c767', icon: 'PictureLibrary' },
  { key: 'comm-blank', title: 'Communication site: Blank', description: 'Start from a blank communication site.', group: 'Modern', webTemplate: 'SITEPAGEPUBLISHING#0', route: 'spsite', siteDesignId: 'f6cc5403-0d63-442e-96c0-285923709ffc', icon: 'Page' },
  { key: 'team-group', title: 'Team site (with Microsoft 365 group)', description: 'Collaborate with a team: group mailbox, calendar, Planner and optional Teams.', group: 'Modern', webTemplate: 'GROUP#0', route: 'group', icon: 'Group' },
  { key: 'team-nogroup', title: 'Team site (no Microsoft 365 group)', description: 'A modern team site without a group, managed with SharePoint permissions.', group: 'Modern', webTemplate: 'STS#3', route: 'spsite', icon: 'People' },
  { key: 'cls-sts0', title: 'Classic team site', description: 'STS#0. Classic experience team site.', group: 'Classic', webTemplate: 'STS#0', route: 'powershell', icon: 'SharepointLogo' },
  { key: 'cls-project', title: 'Project site', description: 'PROJECTSITE#0. Classic project management site.', group: 'Classic', webTemplate: 'PROJECTSITE#0', route: 'powershell', icon: 'ProjectCollection' },
  { key: 'cls-blog', title: 'Blog', description: 'BLOG#0. Classic blog.', group: 'Classic', webTemplate: 'BLOG#0', route: 'powershell', icon: 'Blog' },
  { key: 'cls-doccenter', title: 'Document Center', description: 'BDR#0. Manage documents centrally.', group: 'Classic', webTemplate: 'BDR#0', route: 'powershell', icon: 'Documentation' },
  { key: 'cls-wiki', title: 'Enterprise Wiki', description: 'ENTERWIKI#0. Publishing wiki for knowledge sharing.', group: 'Classic', webTemplate: 'ENTERWIKI#0', route: 'powershell', icon: 'BookAnswers' },
  { key: 'cls-publishing', title: 'Publishing portal', description: 'BLANKINTERNETCONTAINER#0. Classic publishing portal.', group: 'Classic', webTemplate: 'BLANKINTERNETCONTAINER#0', route: 'powershell', icon: 'Globe' },
  { key: 'cls-community', title: 'Community site', description: 'COMMUNITY#0. Discussions and community membership.', group: 'Classic', webTemplate: 'COMMUNITY#0', route: 'powershell', icon: 'ChatInviteFriend' },
  { key: 'cls-records', title: 'Records Center', description: 'OFFILE#1. Records management.', group: 'Classic', webTemplate: 'OFFILE#1', route: 'powershell', icon: 'Archive' },
  { key: 'cls-search', title: 'Basic Search Center', description: 'SRCHCENTERLITE#0. Search results site.', group: 'Classic', webTemplate: 'SRCHCENTERLITE#0', route: 'powershell', icon: 'Search' },
  { key: 'cls-dev', title: 'Developer site', description: 'DEV#0. For app development and testing.', group: 'Classic', webTemplate: 'DEV#0', route: 'powershell', icon: 'Code' }
];

export interface ICreateRequest {
  template: ISiteTemplate;
  title: string;
  /** Managed path: "sites" or "teams". */
  managedPath: string;
  /** URL name (the part after /sites/), also the group e-mail alias for group sites. */
  alias: string;
  description: string;
  lcid: number;
  /** Optional primary owner (e-mail / UPN). Empty = you. */
  owner: string;
  /** Group sites: public or private. */
  isPublic: boolean;
  /** Optional hub to join after creation. */
  hubId?: string;
}

export interface ICreateResult {
  url: string;
  /** 0 not found, 1 provisioning, 2 ready, 3 error. */
  status: number;
  groupId?: string;
  error?: string;
  followUps: string[];
}

export const STATUS_TEXT = ['Not created', 'Being created', 'Ready', 'Failed'];

/** A URL name SharePoint accepts: letters, digits, - and _ (no spaces), max 128. */
export function urlNameFromTitle(title: string): string {
  return title
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Za-z0-9_-]+/g, '')
    .replace(/^[-_]+/, '')
    .substring(0, 64);
}

export function validAlias(alias: string): boolean {
  return /^[A-Za-z0-9][A-Za-z0-9_-]{0,63}$/.test(alias);
}

function psQuote(v: string): string {
  return "'" + v.replace(/'/g, "''") + "'";
}

/** The PowerShell equivalent (SharePoint Online Management Shell or PnP), for any template. */
export function powerShellFor(r: ICreateRequest, origin: string, currentUser: string): string {
  const url = `${origin}/${r.managedPath}/${r.alias}`;
  const owner = r.owner.trim() || currentUser || '<owner@yourtenant.com>';
  if (r.template.route === 'group') {
    return [
      '# PnP PowerShell (Connect-PnPOnline -Url ' + origin + ' -Interactive)',
      `New-PnPSite -Type TeamSite -Title ${psQuote(r.title)} -Alias ${psQuote(r.alias)}${r.description ? ' -Description ' + psQuote(r.description) : ''} -IsPublic:$${r.isPublic ? 'true' : 'false'} -Lcid ${r.lcid} -Owners ${psQuote(owner)}`
    ].join('\n');
  }
  const lines = [
    '# SharePoint Online Management Shell (Connect-SPOService -Url ' + origin.replace('.sharepoint.com', '-admin.sharepoint.com') + ')',
    `New-SPOSite -Url ${psQuote(url)} -Title ${psQuote(r.title)} -Owner ${psQuote(owner)} -Template ${psQuote(r.template.webTemplate)} -LocaleId ${r.lcid} -StorageQuota 1024`
  ];
  if (r.template.customDesignId) {
    lines.push(`Invoke-SPOSiteDesign -Identity ${r.template.customDesignId} -WebUrl ${psQuote(url)}`);
  }
  if (r.hubId) {
    lines.push(`Add-SPOHubSiteAssociation -Site ${psQuote(url)} -HubSite ${r.hubId}`);
  }
  return lines.join('\n');
}

export class SiteCreator {
  private origin: string;

  constructor(private client: SPClient, private homeWebUrl: string) {
    this.origin = (/^(https:\/\/[^/]+)/i.exec(homeWebUrl) || [''])[1];
  }

  private get home(): string {
    return trimSlash(this.homeWebUrl);
  }

  /** Built-in templates plus your organisation's site templates (site designs). */
  public async templates(): Promise<ISiteTemplate[]> {
    let custom: ISiteTemplate[] = [];
    try {
      const j = await this.client.post<unknown>(`${this.home}/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`, {}, { keepCache: true }, this.home);
      custom = rowsOf<{ Id?: string; Title?: string; Description?: string; WebTemplate?: string }>(j)
        .filter((d) => d.Id && (d.WebTemplate === '64' || d.WebTemplate === '68' || d.WebTemplate === '1'))
        .map((d) => {
          const base = d.WebTemplate === '64' ? 'GROUP#0' : d.WebTemplate === '68' ? 'SITEPAGEPUBLISHING#0' : 'STS#3';
          return {
            key: 'custom-' + String(d.Id),
            title: String(d.Title || 'Site template'),
            description: (d.Description ? String(d.Description) + ' ' : '') + `(based on ${d.WebTemplate === '64' ? 'team site with group' : d.WebTemplate === '68' ? 'communication site' : 'team site without group'})`,
            group: 'Your organisation' as const,
            webTemplate: base,
            route: d.WebTemplate === '64' ? ('group' as const) : ('spsite' as const),
            customDesignId: String(d.Id),
            icon: 'Org'
          };
        });
    } catch {
      /* No custom templates, or not allowed to list them: built-ins only. */
    }
    return BUILT_IN_TEMPLATES.slice(0, 5).concat(custom, BUILT_IN_TEMPLATES.slice(5));
  }

  public async languages(): Promise<Array<{ lcid: number; name: string }>> {
    try {
      const j = entityOf<{ Items?: unknown }>(await this.client.get(`${this.home}/_api/web/RegionalSettings/InstalledLanguages`));
      const items = rowsOf<{ Lcid?: number; DisplayName?: string }>(j.Items || j);
      const list = items.filter((l) => l.Lcid).map((l) => ({ lcid: Number(l.Lcid), name: String(l.DisplayName || l.Lcid) }));
      return list.length ? list.sort((a, b) => a.name.localeCompare(b.name)) : [{ lcid: 1033, name: 'English' }];
    } catch {
      return [{ lcid: 1033, name: 'English' }];
    }
  }

  public async hubs(): Promise<Array<{ id: string; title: string; url: string }>> {
    try {
      const j = await this.client.get(`${this.home}/_api/HubSites?$select=ID,Title,SiteUrl`);
      return rowsOf<{ ID?: string; Title?: string; SiteUrl?: string }>(j)
        .filter((h) => h.ID)
        .map((h) => ({ id: String(h.ID), title: String(h.Title || h.SiteUrl), url: String(h.SiteUrl || '') }))
        .sort((a, b) => a.title.localeCompare(b.title));
    } catch {
      return [];
    }
  }

  /** 0 = the address is free. */
  public async status(url: string): Promise<number> {
    const j = entityOf<{ SiteStatus?: number; Status?: { SiteStatus?: number } }>(
      await this.client.get(`${this.origin}/_api/SPSiteManager/status?url='${encodeURIComponent(odataString(url))}'`, { cacheTtlMs: 0 })
    );
    return Number(j.SiteStatus !== undefined ? j.SiteStatus : j.Status && j.Status.SiteStatus) || 0;
  }

  /** For group sites: the URL SharePoint would use for this alias (a different one means it is taken). */
  public async groupSiteUrlFor(alias: string, managedPath: string): Promise<string | undefined> {
    const j = entityOf<{ value?: string; GetValidSiteUrlFromAlias?: string }>(
      await this.client.get(`${this.home}/_api/GroupSiteManager/GetValidSiteUrlFromAlias?alias='${encodeURIComponent(odataString(alias))}'&managedPath='${encodeURIComponent(managedPath)}'`, { cacheTtlMs: 0 })
    );
    const v = j.value || j.GetValidSiteUrlFromAlias;
    return typeof v === 'string' ? trimSlash(v) : undefined;
  }

  /** Is the address free? Returns a reason when it is not. */
  /**
   * Is the address free? Returns a reason when it is not. For group sites SharePoint picks
   * the managed path (/sites/ or /teams/, a tenant setting), so the URL it reports is used.
   */
  public async checkAvailable(r: ICreateRequest): Promise<{ ok: boolean; message: string; suggestion?: string; url?: string }> {
    if (!validAlias(r.alias)) {
      return { ok: false, message: 'Use letters, numbers, - and _ only (no spaces), starting with a letter or number.' };
    }
    let url = `${this.origin}/${r.managedPath}/${r.alias}`;
    if (r.template.route === 'group') {
      try {
        const valid = await this.groupSiteUrlFor(r.alias, r.managedPath);
        if (valid) {
          const name = (valid.split('/').pop() || '').toLowerCase();
          if (name !== r.alias.toLowerCase()) {
            return { ok: false, message: 'That name is already used by a site or group.', suggestion: valid.split('/').pop(), url: valid };
          }
          url = valid;
        }
      } catch {
        /* fall through to the status check */
      }
    }
    const s = await this.status(url);
    return s === 0
      ? { ok: true, message: `${url} is available.`, url }
      : { ok: false, message: `${url} is already taken (${STATUS_TEXT[s] || 'exists'}). Deleted sites keep their address until they are removed from the site recycle bin.`, url };
  }

  public async create(r: ICreateRequest): Promise<ICreateResult> {
    if (r.template.route === 'powershell') {
      throw new Error('Classic templates can only be created with PowerShell or the SharePoint admin center. Copy the command shown.');
    }
    const url = `${this.origin}/${r.managedPath}/${r.alias}`;
    const followUps: string[] = [];
    let result: ICreateResult;
    if (r.template.route === 'group') {
      const optionalParams: { [k: string]: unknown } = { Description: r.description || '', CreationOptions: [`SPSiteLanguage:${r.lcid}`] };
      if (r.owner.trim()) {
        optionalParams.Owners = [r.owner.trim()];
      }
      const j = entityOf<{ SiteUrl?: string; SiteStatus?: number; GroupId?: string; ErrorMessage?: string; CreateGroupEx?: { SiteUrl?: string; SiteStatus?: number; GroupId?: string; ErrorMessage?: string } }>(
        await this.client.post(`${this.home}/_api/GroupSiteManager/CreateGroupEx`, { displayName: r.title, alias: r.alias, isPublic: r.isPublic, optionalParams }, {}, this.home)
      );
      const v = j.CreateGroupEx || j;
      if (v.ErrorMessage) {
        throw new Error(String(v.ErrorMessage));
      }
      result = { url: trimSlash(String(v.SiteUrl || url)), status: Number(v.SiteStatus) || 1, groupId: v.GroupId && v.GroupId !== EMPTY ? String(v.GroupId) : undefined, followUps };
    } else {
      const request: { [k: string]: unknown } = {
        Title: r.title,
        Url: url,
        Lcid: r.lcid,
        Description: r.description || '',
        WebTemplate: r.template.webTemplate,
        ShareByEmailEnabled: false
      };
      if (r.template.siteDesignId) {
        request.SiteDesignId = r.template.siteDesignId;
      }
      if (r.template.customDesignId) {
        request.WebTemplateExtensionId = r.template.customDesignId;
      }
      if (r.owner.trim()) {
        request.Owner = r.owner.trim();
      }
      const j = entityOf<{ SiteUrl?: string; SiteStatus?: number; Create?: { SiteUrl?: string; SiteStatus?: number } }>(
        await this.client.post(`${this.origin}/_api/SPSiteManager/create`, { request }, {}, this.home)
      );
      const v = j.Create || j;
      result = { url: trimSlash(String(v.SiteUrl || url)), status: Number(v.SiteStatus) || 0, followUps };
      if (result.status === 3) {
        throw new Error('SharePoint could not create the site. Check that the address is free and that you are allowed to create sites.');
      }
    }
    return result;
  }

  /** Re-checks a site being provisioned. */
  public async refreshStatus(res: ICreateResult): Promise<number> {
    if (res.groupId) {
      try {
        const j = entityOf<{ SiteStatus?: number; SiteUrl?: string; GetSiteStatus?: { SiteStatus?: number; SiteUrl?: string } }>(
          await this.client.get(`${this.home}/_api/GroupSiteManager/GetSiteStatus('${res.groupId}')`, { cacheTtlMs: 0 })
        );
        const v = j.GetSiteStatus || j;
        if (v.SiteUrl) {
          res.url = trimSlash(String(v.SiteUrl));
        }
        return Number(v.SiteStatus) || 0;
      } catch (e) {
        if (!(e instanceof HttpError)) {
          throw e;
        }
      }
    }
    return this.status(res.url);
  }

  /** Steps that need the site to exist: custom template on group sites, hub association. */
  public async finish(r: ICreateRequest, res: ICreateResult): Promise<string[]> {
    const done: string[] = [];
    if (r.template.route === 'group' && r.template.customDesignId) {
      try {
        await this.client.post(`${res.url}/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility.AddSiteDesignTaskToCurrentWeb`, { siteDesignId: r.template.customDesignId }, {}, res.url);
        done.push(`Applying "${r.template.title}" (runs in the background for a few minutes).`);
      } catch (e) {
        done.push(`Could not apply the site template: ${(e as Error).message}`);
      }
    }
    if (r.hubId) {
      try {
        await this.client.post(`${res.url}/_api/site/JoinHubSite('${r.hubId}')`, undefined, {}, res.url);
        done.push('Joined the hub.');
      } catch (e) {
        done.push(`Could not join the hub (you may need hub join permission): ${(e as Error).message}`);
      }
    }
    return done;
  }
}
