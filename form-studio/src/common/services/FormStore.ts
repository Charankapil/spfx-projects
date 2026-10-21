import { SPHttpClient } from '@microsoft/sp-http';

import { FORMS_SCHEMA_VERSION, IFormDefinition, IListForms } from '../model/IFormDefinition';
import { ensureOk, getJson, guidOnly, HttpError, ISpContext, quoteForUrl } from './http';

const FOLDER = 'FormStudio';

interface IRootFolder {
  RootFolder: { ServerRelativeUrl: string };
}

/**
 * Form definitions live as one JSON file per list in the site's Site Assets
 * library: SiteAssets/FormStudio/<list id>.json. Everyone who can read Site
 * Assets (by default, everyone who can read the site) can load the forms;
 * only people who can edit Site Assets can change them, and the designer
 * additionally requires Manage Lists on the list itself.
 *
 * Plain files keep forms portable: copy the file to another site or check it
 * into source control, and the designer's Export / Import uses the same JSON.
 */
export class FormStore {
  private siteAssets: Promise<string | undefined> | undefined;

  constructor(private ctx: ISpContext, private webUrl: string = ctx.pageContext.web.absoluteUrl) {}

  private findSiteAssets(): Promise<string | undefined> {
    if (!this.siteAssets) {
      // Site Assets' URL name is "SiteAssets" regardless of the site's language.
      this.siteAssets = getJson<{ value: IRootFolder[] }>(
        this.ctx,
        `${this.webUrl}/_api/web/lists?$filter=BaseTemplate eq 101&$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`
      ).then((json) => {
        const match = json.value.filter((l) => /\/siteassets$/i.test(l.RootFolder.ServerRelativeUrl))[0];
        return match ? match.RootFolder.ServerRelativeUrl : undefined;
      });
      this.siteAssets.catch(() => (this.siteAssets = undefined));
    }
    return this.siteAssets;
  }

  private fileName(listId: string): string {
    return `${guidOnly(listId).toLowerCase()}.json`;
  }

  /** Returns an empty set when no forms were saved for the list yet. */
  public async load(listId: string, listTitle = ''): Promise<IListForms> {
    const empty: IListForms = { schemaVersion: FORMS_SCHEMA_VERSION, listId: guidOnly(listId), listTitle, forms: [] };
    const assets = await this.findSiteAssets();
    if (!assets) {
      return empty;
    }
    const url = `${this.webUrl}/_api/web/GetFileByServerRelativeUrl('${quoteForUrl(`${assets}/${FOLDER}/${this.fileName(listId)}`)}')/$value`;
    const response = await this.ctx.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (response.status === 404) {
      return empty;
    }
    await ensureOk(response);
    try {
      const parsed = JSON.parse(await response.text()) as IListForms;
      if (!parsed || !Array.isArray(parsed.forms)) {
        return empty;
      }
      return parsed;
    } catch {
      return empty;
    }
  }

  public async save(data: IListForms, modifiedBy: string): Promise<void> {
    const assets = await this.ensureSiteAssets();
    const folder = `${assets}/${FOLDER}`;
    await this.ensureFolder(assets, folder);
    const now = new Date().toISOString();
    const body: IListForms = {
      ...data,
      schemaVersion: FORMS_SCHEMA_VERSION,
      forms: data.forms.map((f) => ({ ...f, modified: f.modified || now, modifiedBy: f.modifiedBy || modifiedBy }))
    };
    const url = `${this.webUrl}/_api/web/GetFolderByServerRelativeUrl('${quoteForUrl(folder)}')/Files/add(url='${this.fileName(data.listId)}',overwrite=true)`;
    await ensureOk(await this.ctx.spHttpClient.post(url, SPHttpClient.configurations.v1, { body: JSON.stringify(body, undefined, 2) }));
  }

  private async ensureSiteAssets(): Promise<string> {
    let assets = await this.findSiteAssets();
    if (!assets) {
      await ensureOk(
        await this.ctx.spHttpClient.post(`${this.webUrl}/_api/web/lists/EnsureSiteAssetsLibrary()`, SPHttpClient.configurations.v1, {})
      );
      this.siteAssets = undefined;
      assets = await this.findSiteAssets();
      if (!assets) {
        throw new Error('The Site Assets library could not be found or created.');
      }
    }
    return assets;
  }

  private async ensureFolder(parent: string, folder: string): Promise<void> {
    const check = await this.ctx.spHttpClient.get(
      `${this.webUrl}/_api/web/GetFolderByServerRelativeUrl('${quoteForUrl(folder)}')?$select=Exists`,
      SPHttpClient.configurations.v1
    );
    if (check.ok) {
      const json = await check.json();
      if (json.Exists !== false) {
        return;
      }
    } else if (check.status !== 404) {
      throw new HttpError(check.status, `Could not open ${folder}.`);
    }
    await ensureOk(
      await this.ctx.spHttpClient.post(
        `${this.webUrl}/_api/web/GetFolderByServerRelativeUrl('${quoteForUrl(parent)}')/Folders/add('${FOLDER}')`,
        SPHttpClient.configurations.v1,
        {}
      )
    );
  }
}

/** Stamps a form as modified now (the store keeps an existing stamp otherwise). */
export function touch(form: IFormDefinition, by: string): IFormDefinition {
  return { ...form, modified: new Date().toISOString(), modifiedBy: by };
}
