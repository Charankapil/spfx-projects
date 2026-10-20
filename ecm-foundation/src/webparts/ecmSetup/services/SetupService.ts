import { SPHttpClient } from '@microsoft/sp-http';

import { normalizeUrl, SpRestClient } from '../../../common/services/SpRestClient';
import { buildSharePointTheme, ECM_THEME_NAME, IEcmBrand } from '../../../common/theme/ecmTheme';
import { NAVIGATION_SEED } from '../data/navigationSeed';

/** The navigator's component id, from its manifest. */
export const NAVIGATOR_COMPONENT_ID = '293dd12c-1019-42fd-b262-79b6f949e9b7';

export type Log = (line: string) => void;

/** What ECM Setup keeps in the navigator's settings, so it is there next time. */
export interface ISavedSettings {
  brandPrimary?: string;
  brandDark?: string;
  logoUrl?: string;
  brandBadge?: string;
  homeTitle?: string;
}

export interface ILinkCheck {
  title: string;
  type: string;
  url: string;
  status: 'ok' | 'no-access' | 'missing' | 'error';
  detail: string;
}

interface IFieldValue {
  FieldName: string;
  FieldValue: string;
}

interface IValidateResult {
  value?: { FieldName: string; ErrorMessage: string | null; HasException: boolean }[];
}

interface INavItem {
  Id: number;
  Title: string;
  EcmNavType: string;
  EcmParentHub: string | null;
  EcmUrl: { Url: string } | null;
}

/**
 * The columns of the navigation list. Name is the internal name the
 * navigator reads; option 8 (AddFieldInternalNameHint) makes SharePoint keep
 * it instead of deriving one from the display name, 16 adds it to the view.
 */
const FIELDS: { name: string; xml: string }[] = [
  {
    name: 'EcmNavType',
    xml: '<Field Type="Choice" Name="EcmNavType" StaticName="EcmNavType" DisplayName="Type" Required="TRUE" Format="Dropdown"><CHOICES><CHOICE>Hub</CHOICE><CHOICE>Department site</CHOICE><CHOICE>Shared service</CHOICE><CHOICE>Hub resource</CHOICE><CHOICE>Header link</CHOICE></CHOICES></Field>'
  },
  {
    name: 'EcmParentHub',
    xml: '<Field Type="Text" Name="EcmParentHub" StaticName="EcmParentHub" DisplayName="Parent hub" Description="Title of the hub this entry belongs to (department sites and hub resources)." />'
  },
  { name: 'EcmUrl', xml: '<Field Type="URL" Name="EcmUrl" StaticName="EcmUrl" DisplayName="URL" Format="Hyperlink" Required="TRUE" />' },
  {
    name: 'EcmDescription',
    xml: '<Field Type="Text" Name="EcmDescription" StaticName="EcmDescription" DisplayName="Description" Description="Shown under the site name; for a hub, the governance note." />'
  },
  { name: 'EcmOrder', xml: '<Field Type="Number" Name="EcmOrder" StaticName="EcmOrder" DisplayName="Order" Decimals="0"><Default>100</Default></Field>' },
  {
    name: 'EcmRestricted',
    xml: '<Field Type="Boolean" Name="EcmRestricted" StaticName="EcmRestricted" DisplayName="Restricted" Description="Shown only to people who can open the site."><Default>0</Default></Field>'
  },
  { name: 'EcmActive', xml: '<Field Type="Boolean" Name="EcmActive" StaticName="EcmActive" DisplayName="Active"><Default>1</Default></Field>' },
  {
    name: 'EcmColor',
    xml: '<Field Type="Text" Name="EcmColor" StaticName="EcmColor" DisplayName="Hub colour" Description="Optional hex colour for a hub, e.g. #B5470E." />'
  }
];

const READ_ROLE_TYPE = 2;

function odataString(value: string): string {
  return value.replace(/'/g, "''");
}

/**
 * Everything ECM Setup does, through SharePoint REST as the signed-in user:
 * no app registration, client id or API permission. Each step only adds what
 * is missing, so it is safe to run again.
 */
export class SetupService {
  private rest: SpRestClient;
  private origin: string;

  constructor(httpClient: SPHttpClient, private webUrl: string, private listTitle: string) {
    this.rest = new SpRestClient(httpClient, 'ECM Setup');
    this.origin = /^(https?:\/\/[^/]+)/i.exec(webUrl)?.[1] || '';
  }

  private get listApi(): string {
    return `${this.webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(odataString(this.listTitle))}')`;
  }

  public get serverRelativeWebUrl(): string {
    return this.webUrl.substring(this.origin.length) || '/';
  }

  // --- Step: navigation list --------------------------------------------------

  public async ensureNavigationList(log: Log): Promise<void> {
    let list = await this.getList();
    if (!list) {
      await this.rest.postJson(`${this.webUrl}/_api/web/lists`, {
        Title: this.listTitle,
        BaseTemplate: 100,
        Description: 'Drives the ECM Hub Navigator. Managed by the ECM governance team.'
      });
      log(`Created the list "${this.listTitle}".`);
      list = await this.getList();
    } else {
      log(`The list "${this.listTitle}" already exists.`);
    }
    if (!list) {
      throw new Error('The list was created but could not be read back.');
    }

    const fields = await this.rest.getJson<{ value: { InternalName: string }[] }>(`${this.listApi}/fields?$select=InternalName`);
    const existing = (fields.value || []).map((f) => f.InternalName);
    for (const field of FIELDS) {
      if (existing.indexOf(field.name) < 0) {
        await this.rest.postJson(`${this.listApi}/fields/CreateFieldAsXml`, {
          parameters: { SchemaXml: field.xml, Options: 8 | 16 }
        });
        log(`Added column ${field.name}.`);
      }
    }

    const items = await this.getNavItems();
    const have = items.map((i) => `${i.Title}|${i.EcmNavType}|${i.EcmParentHub || ''}`.toLowerCase());
    let added = 0;
    for (const row of NAVIGATION_SEED) {
      if (have.indexOf(`${row.title}|${row.type}|${row.parentHub}`.toLowerCase()) >= 0) {
        continue;
      }
      const url = row.url.indexOf('/') === 0 ? this.origin + row.url : row.url;
      await this.addItem(list.RootFolder.ServerRelativeUrl, [
        { FieldName: 'Title', FieldValue: row.title },
        { FieldName: 'EcmNavType', FieldValue: row.type },
        { FieldName: 'EcmParentHub', FieldValue: row.parentHub },
        // URL values are "address, description"; a comma in the address is escaped by doubling it.
        { FieldName: 'EcmUrl', FieldValue: `${url.replace(/,/g, ',,')}, ${row.title}` },
        { FieldName: 'EcmDescription', FieldValue: row.description },
        { FieldName: 'EcmOrder', FieldValue: String(row.order) },
        { FieldName: 'EcmRestricted', FieldValue: row.restricted ? '1' : '0' },
        { FieldName: 'EcmActive', FieldValue: '1' }
      ]);
      added++;
    }
    log(added ? `Added ${added} navigation entries.` : 'All navigation entries were already there; nothing was changed.');
    this.clearNavigatorCache();
  }

  private async getList(): Promise<{ Id: string; RootFolder: { ServerRelativeUrl: string } } | undefined> {
    try {
      return await this.rest.getJson(`${this.listApi}?$select=Id,RootFolder/ServerRelativeUrl&$expand=RootFolder`);
    } catch (err) {
      if ((err as { status?: number }).status === 404) {
        return undefined;
      }
      throw err;
    }
  }

  private async getNavItems(): Promise<INavItem[]> {
    const json = await this.rest.getJson<{ value: INavItem[] }>(
      `${this.listApi}/items?$select=Id,Title,EcmNavType,EcmParentHub,EcmUrl&$top=5000`
    );
    return json.value || [];
  }

  /** AddValidateUpdateItemUsingPath takes display-format strings, so no entity type names are needed. */
  private async addItem(folderUrl: string, values: IFieldValue[]): Promise<void> {
    const result = await this.rest.postJson<IValidateResult>(`${this.listApi}/AddValidateUpdateItemUsingPath`, {
      listItemCreateInfo: { FolderPath: { DecodedUrl: folderUrl }, UnderlyingObjectType: 0 },
      formValues: values,
      bNewDocumentUpdate: false
    });
    const failed = (result?.value || []).filter((v) => v.HasException);
    if (failed.length) {
      throw new Error(failed.map((f) => `${f.FieldName}: ${f.ErrorMessage}`).join('; '));
    }
  }

  /** So the person running setup sees the new menu straight away instead of after the cache expires. */
  private clearNavigatorCache(): void {
    try {
      for (let i = localStorage.length - 1; i >= 0; i--) {
        const key = localStorage.key(i);
        if (key && key.indexOf('ecm-nav:') === 0) {
          localStorage.removeItem(key);
        }
      }
      for (let i = sessionStorage.length - 1; i >= 0; i--) {
        const key = sessionStorage.key(i);
        if (key && key.indexOf('ecm-nav-access:') === 0) {
          sessionStorage.removeItem(key);
        }
      }
    } catch {
      // Storage blocked: nothing cached to clear.
    }
  }

  // --- Step: everyone can read the list ---------------------------------------

  public async grantEveryoneRead(tenantId: string, log: Log): Promise<void> {
    const list = await this.rest.getJson<{ HasUniqueRoleAssignments: boolean }>(`${this.listApi}?$select=HasUniqueRoleAssignments`);
    if (!list.HasUniqueRoleAssignments) {
      await this.rest.postJson(`${this.listApi}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)`);
      log('Gave the list its own permissions (copied from the site, so owners keep access).');
    }
    const claim = `c:0-.f|rolemanager|spo-grid-all-users/${tenantId}`;
    let user: { Id: number };
    try {
      user = await this.rest.postJson<{ Id: number }>(`${this.webUrl}/_api/web/ensureuser`, { logonName: claim });
    } catch (err) {
      throw new Error(
        `"Everyone except external users" is not available in this tenant (${(err as Error).message}). ` +
          'Give the list read access to a group that contains every employee instead.'
      );
    }
    const role = await this.rest.getJson<{ Id: number }>(`${this.webUrl}/_api/web/roledefinitions/getbytype(${READ_ROLE_TYPE})?$select=Id`);
    await this.rest.postJson(`${this.listApi}/roleassignments/addroleassignment(principalid=${user.Id},roledefid=${role.Id})`);
    log('"Everyone except external users" can now read the navigation list.');
  }

  // --- Step: tenant theme -----------------------------------------------------

  private themeJson(brand: IEcmBrand): string {
    const theme = buildSharePointTheme(brand);
    return JSON.stringify({ palette: theme.palette, isInverted: theme.isInverted });
  }

  /** Adds the theme to the "Change the look" picker for every site. Needs a SharePoint administrator. */
  public async registerTenantTheme(brand: IEcmBrand, log: Log): Promise<void> {
    const api = `${this.webUrl}/_api/thememanager`;
    const body = { name: ECM_THEME_NAME, themeJson: this.themeJson(brand) };
    let added = false;
    try {
      const result = await this.rest.postJson<{ value?: boolean }>(`${api}/AddTenantTheme`, body);
      added = !!result?.value;
    } catch (err) {
      if ((err as { status?: number }).status === 403 || (err as { status?: number }).status === 401) {
        throw new Error('Only a SharePoint administrator can register a tenant theme. The next step still applies the theme to sites.');
      }
      // Most often "already exists": fall through to update.
    }
    if (added) {
      log(`Registered "${ECM_THEME_NAME}" as a tenant theme.`);
      return;
    }
    await this.rest.postJson(`${api}/UpdateTenantTheme`, body);
    log(`Updated the existing tenant theme "${ECM_THEME_NAME}".`);
  }

  // --- Step: apply theme ------------------------------------------------------

  /** The hub plus every hub, department and shared-service site in the navigation list. */
  public async getThemeTargets(): Promise<string[]> {
    const urls = [this.webUrl];
    try {
      const items = await this.getNavItems();
      items
        .filter((i) => i.EcmNavType === 'Hub' || i.EcmNavType === 'Department site' || i.EcmNavType === 'Shared service')
        .forEach((i) => {
          const site = /^(https?:\/\/[^/]+\/(?:sites|teams)\/[^/?#]+)/i.exec(i.EcmUrl?.Url || '');
          if (site && urls.map((u) => u.toLowerCase()).indexOf(site[1].toLowerCase()) < 0) {
            urls.push(site[1]);
          }
        });
    } catch {
      // No list yet: just the hub.
    }
    return urls;
  }

  /** ApplyTheme sets the palette directly, so it works even if the tenant theme was not registered. */
  public async applyTheme(brand: IEcmBrand, siteUrls: string[], log: Log): Promise<{ applied: number; failed: number }> {
    const themeJson = this.themeJson(brand);
    let applied = 0;
    let failed = 0;
    for (const raw of siteUrls) {
      const site = raw.trim().replace(/\/+$/, '');
      if (!site) {
        continue;
      }
      try {
        await this.rest.postJson(`${site}/_api/ThemeManager/ApplyTheme`, { name: ECM_THEME_NAME, themeJson });
        log(`Applied to ${site}`);
        applied++;
      } catch (err) {
        const status = (err as { status?: number }).status;
        log(
          `Skipped ${site}: ${
            status === 403 || status === 401 ? 'you are not an owner of this site' : status === 404 ? 'site not found' : (err as Error).message
          }`
        );
        failed++;
      }
    }
    return { applied, failed };
  }

  // --- Step: navigator settings ----------------------------------------------

  /**
   * Points the tenant-wide navigator at this site's list. Deploying the
   * package with "make available to all sites" creates the Tenant Wide
   * Extensions entry; if it is missing this adds it. Needs owner rights on the
   * tenant app catalog (normally a SharePoint administrator).
   */
  public async configureNavigator(brand: IEcmBrand, brandBadge: string, homeTitle: string, log: Log): Promise<void> {
    const settings = await this.rest.getJson<{ CorporateCatalogUrl?: string }>(`${this.webUrl}/_api/SP_TenantSettings_Current`);
    const catalog = settings.CorporateCatalogUrl;
    if (!catalog) {
      throw new Error('This tenant has no app catalog yet. Create one in the SharePoint admin center and deploy the package first.');
    }
    const listApi = `${catalog}/_api/web/lists/getbytitle('Tenant Wide Extensions')`;
    let entries: { Id: number; TenantWideExtensionComponentId: string; TenantWideExtensionComponentProperties: string | null }[];
    let rootFolder: string;
    try {
      const list = await this.rest.getJson<{ RootFolder: { ServerRelativeUrl: string } }>(`${listApi}?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`);
      rootFolder = list.RootFolder.ServerRelativeUrl;
      const items = await this.rest.getJson<{ value: typeof entries }>(
        `${listApi}/items?$select=Id,TenantWideExtensionComponentId,TenantWideExtensionComponentProperties&$top=500`
      );
      entries = (items.value || []).filter((i) => (i.TenantWideExtensionComponentId || '').toLowerCase().indexOf(NAVIGATOR_COMPONENT_ID) >= 0);
    } catch (err) {
      const status = (err as { status?: number }).status;
      if (status === 403 || status === 401 || status === 404) {
        throw new Error(
          `You need owner access to the app catalog (${catalog}) for this step, normally a SharePoint administrator. ` +
            `Alternatively, set configSiteUrl to "${this.serverRelativeWebUrl}" in its Tenant Wide Extensions list by hand.`
        );
      }
      throw err;
    }

    const wanted: { [key: string]: string } = {
      configSiteUrl: this.serverRelativeWebUrl,
      listTitle: this.listTitle,
      brandPrimary: brand.primary,
      brandDark: brand.dark,
      logoUrl: brand.logoUrl
    };
    if (homeTitle.trim()) {
      wanted.homeTitle = homeTitle.trim();
    }
    if (brandBadge.trim()) {
      wanted.brandBadge = brandBadge.trim();
    }

    if (!entries.length) {
      await this.rest.postJson(`${listApi}/AddValidateUpdateItemUsingPath`, {
        listItemCreateInfo: { FolderPath: { DecodedUrl: rootFolder }, UnderlyingObjectType: 0 },
        formValues: [
          { FieldName: 'Title', FieldValue: 'ECM Hub Navigator' },
          { FieldName: 'TenantWideExtensionComponentId', FieldValue: NAVIGATOR_COMPONENT_ID },
          { FieldName: 'TenantWideExtensionLocation', FieldValue: 'ClientSideExtension.ApplicationCustomizer' },
          { FieldName: 'TenantWideExtensionComponentProperties', FieldValue: JSON.stringify(wanted) },
          { FieldName: 'TenantWideExtensionSequence', FieldValue: '0' },
          { FieldName: 'TenantWideExtensionListTemplate', FieldValue: '0' },
          { FieldName: 'TenantWideExtensionWebTemplate', FieldValue: '' },
          { FieldName: 'TenantWideExtensionDisabled', FieldValue: '0' }
        ],
        bNewDocumentUpdate: false
      });
      log('Added the navigator to every site (Tenant Wide Extensions), pointing at this hub.');
      return;
    }

    for (const entry of entries) {
      let current: { [key: string]: unknown } = {};
      try {
        current = JSON.parse(entry.TenantWideExtensionComponentProperties || '{}');
      } catch {
        // Unreadable properties are replaced.
      }
      const merged = JSON.stringify({ ...current, ...wanted });
      await this.rest.postJson(`${listApi}/items(${entry.Id})/ValidateUpdateListItem`, {
        formValues: [{ FieldName: 'TenantWideExtensionComponentProperties', FieldValue: merged }],
        bNewDocumentUpdate: false
      });
    }
    log(`The navigator now reads its menu from this site (${this.serverRelativeWebUrl}). Sites pick this up on their next page load.`);
  }

  // --- Brand: saved settings and logo ----------------------------------------

  private async findNavigatorEntries(): Promise<{ catalog: string; entries: { Id: number; TenantWideExtensionComponentProperties: string | null }[] }> {
    const settings = await this.rest.getJson<{ CorporateCatalogUrl?: string }>(`${this.webUrl}/_api/SP_TenantSettings_Current`);
    const catalog = settings.CorporateCatalogUrl || '';
    if (!catalog) {
      return { catalog, entries: [] };
    }
    const items = await this.rest.getJson<{ value: { Id: number; TenantWideExtensionComponentId: string; TenantWideExtensionComponentProperties: string | null }[] }>(
      `${catalog}/_api/web/lists/getbytitle('Tenant Wide Extensions')/items?$select=Id,TenantWideExtensionComponentId,TenantWideExtensionComponentProperties&$top=500`
    );
    return {
      catalog,
      entries: (items.value || []).filter((i) => (i.TenantWideExtensionComponentId || '').toLowerCase().indexOf(NAVIGATOR_COMPONENT_ID) >= 0)
    };
  }

  /** The brand saved by an earlier run (step 5), if this person can read the app catalog. */
  public async loadSavedSettings(): Promise<ISavedSettings> {
    try {
      const { entries } = await this.findNavigatorEntries();
      return entries.length ? (JSON.parse(entries[0].TenantWideExtensionComponentProperties || '{}') as ISavedSettings) : {};
    } catch {
      return {};
    }
  }

  /**
   * Uploads the logo to this site's Site Assets library and returns its URL.
   * Site Assets follows the site's permissions, so when `everyoneRead` is set
   * the file alone is opened to Everyone except external users; otherwise
   * people who cannot open the hub see the text badge instead of the logo.
   */
  public async uploadLogo(file: File, everyoneRead: boolean, tenantId: string, log: Log): Promise<string> {
    if (!/^image\/(png|svg\+xml|jpeg|gif|webp)$/.test(file.type)) {
      throw new Error('Choose a PNG, SVG, JPG, GIF or WebP image.');
    }
    if (file.size > 1024 * 1024) {
      throw new Error('The logo should be under 1 MB; it loads on every page.');
    }
    const library = await this.rest.postJson<{ RootFolder?: { ServerRelativeUrl: string }; Id: string }>(
      `${this.webUrl}/_api/web/lists/EnsureSiteAssetsLibrary`
    );
    const folder = library.RootFolder?.ServerRelativeUrl ||
      (await this.rest.getJson<{ ServerRelativeUrl: string }>(`${this.webUrl}/_api/web/lists(guid'${library.Id}')/RootFolder?$select=ServerRelativeUrl`)).ServerRelativeUrl;
    const extension = (file.name.split('.').pop() || 'png').toLowerCase().replace(/[^a-z0-9]/g, '');
    const fileName = `ecm-logo.${extension}`;
    const content = await file.arrayBuffer();
    const uploaded = await this.rest.postBinary<{ ServerRelativeUrl: string }>(
      `${this.webUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${encodeURIComponent(odataString(folder))}')/Files/AddUsingPath(decodedurl='${fileName}',overwrite=true)`,
      content
    );
    log(`Uploaded ${file.name} to ${folder}.`);
    if (everyoneRead) {
      const itemApi = `${this.webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${encodeURIComponent(odataString(uploaded.ServerRelativeUrl))}')/ListItemAllFields`;
      const item = await this.rest.getJson<{ HasUniqueRoleAssignments: boolean }>(`${itemApi}?$select=HasUniqueRoleAssignments`);
      if (!item.HasUniqueRoleAssignments) {
        await this.rest.postJson(`${itemApi}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)`);
      }
      const user = await this.rest.postJson<{ Id: number }>(`${this.webUrl}/_api/web/ensureuser`, {
        logonName: `c:0-.f|rolemanager|spo-grid-all-users/${tenantId}`
      });
      const role = await this.rest.getJson<{ Id: number }>(`${this.webUrl}/_api/web/roledefinitions/getbytype(${READ_ROLE_TYPE})?$select=Id`);
      await this.rest.postJson(`${itemApi}/roleassignments/addroleassignment(principalid=${user.Id},roledefid=${role.Id})`);
      log('Everyone except external users can see the logo.');
    }
    // A query string busts browser caches when the logo is replaced under the same name.
    return `${uploaded.ServerRelativeUrl}?v=${Date.now()}`;
  }

  // --- Step: link check -------------------------------------------------------

  /** Opens every site the menu links to, so placeholder URLs from the seed stand out. */
  public async checkLinks(onResult: (result: ILinkCheck) => void): Promise<void> {
    const items = (await this.getNavItems()).filter(
      (i) => i.EcmNavType === 'Hub' || i.EcmNavType === 'Department site' || i.EcmNavType === 'Shared service'
    );
    const queue = items.slice();
    const worker = async (): Promise<void> => {
      for (let item = queue.shift(); item; item = queue.shift()) {
        const url = item.EcmUrl?.Url || '';
        const base = { title: item.Title, type: item.EcmNavType, url };
        if (!url || normalizeUrl(url, this.origin).indexOf(this.origin.toLowerCase()) !== 0) {
          onResult({ ...base, status: 'error', detail: 'Not a site in this tenant' });
          continue;
        }
        try {
          const web = await this.rest.getJson<{ Title: string }>(`${url.replace(/\/+$/, '')}/_api/web?$select=Title`);
          onResult({ ...base, status: 'ok', detail: web.Title });
        } catch (err) {
          const status = (err as { status?: number }).status;
          onResult({
            ...base,
            status: status === 404 ? 'missing' : status === 401 || status === 403 ? 'no-access' : 'error',
            detail: status === 404 ? 'No site at this URL' : status === 401 || status === 403 ? 'Exists, but you have no access' : (err as Error).message
          });
        }
      }
    };
    await Promise.all([worker(), worker(), worker(), worker()]);
  }

  /** Where the navigation list lives, to link to it once it exists. */
  public async getNavigationListUrl(): Promise<string | undefined> {
    const list = await this.getList().catch(() => undefined);
    return list ? this.origin + list.RootFolder.ServerRelativeUrl : undefined;
  }
}
