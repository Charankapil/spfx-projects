import { CancelToken, httpStatusOf } from './RequestQueue';
import { quote, SpClient } from './SpClient';

/**
 * Lists the toolkit owns are created with a fixed URL name and addressed by
 * that URL, never by title: owners can rename them and the site language can
 * differ without anything breaking.
 */
export interface IListDefinition {
  /** URL name, letters and digits only, e.g. "ToolkitFaq" -> /Lists/ToolkitFaq. */
  urlName: string;
  title: string;
  description: string;
  fields: string[];
  /** Index the Created By column, for "my items" queries on large lists. */
  indexAuthor?: boolean;
  /** Index the Created column, for date-range queries on large lists. */
  indexCreated?: boolean;
  /** People can edit and delete only the items they created. */
  ownItemsOnly?: boolean;
}

function xmlEscape(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function fieldXml(type: string, name: string, displayName: string, attributes: string, inner = ''): string {
  const head =
    `<Field Type="${type}" Name="${name}" StaticName="${name}" DisplayName="${xmlEscape(displayName)}" ${attributes}`;
  return inner ? `${head}>${inner}</Field>` : `${head}/>`;
}

export const field = {
  text: (name: string, displayName: string, indexed = false): string =>
    fieldXml('Text', name, displayName, `MaxLength="255"${indexed ? ' Indexed="TRUE"' : ''}`),
  note: (name: string, displayName: string): string =>
    fieldXml('Note', name, displayName, 'NumLines="6" RichText="FALSE"'),
  number: (name: string, displayName: string, indexed = false): string =>
    fieldXml('Number', name, displayName, `Decimals="0"${indexed ? ' Indexed="TRUE"' : ''}`),
  dateTime: (name: string, displayName: string): string =>
    fieldXml('DateTime', name, displayName, 'Format="DateTime"'),
  user: (name: string, displayName: string): string =>
    fieldXml('User', name, displayName, 'UserSelectionMode="PeopleOnly" List="UserInfo"'),
  choice: (name: string, displayName: string, choices: string[]): string =>
    fieldXml(
      'Choice',
      name,
      displayName,
      'Format="Dropdown"',
      `<Default>${xmlEscape(choices[0])}</Default><CHOICES>${choices
        .map((c) => `<CHOICE>${xmlEscape(c)}</CHOICE>`)
        .join('')}</CHOICES>`
    )
};

/** AddFieldInternalNameHint (8) + AddFieldToDefaultView (16). */
const ADD_FIELD_OPTIONS = 24;

export interface IListItemBase {
  Id: number;
}

export class ListService {
  /**
   * @param sp      request client
   * @param baseUrl absolute URL of the web that holds the list
   * @param baseServerRelativeUrl server-relative URL of that web ('' for a root site)
   */
  constructor(private sp: SpClient, private baseUrl: string, private baseServerRelativeUrl: string) {}

  public static forWeb(sp: SpClient): ListService {
    return new ListService(sp, sp.webUrl, sp.webServerRelativeUrl);
  }

  /** Lists that live once per site collection, on its root web. */
  public static forSiteRoot(sp: SpClient): ListService {
    return new ListService(sp, sp.siteUrl, sp.siteServerRelativeUrl);
  }

  public listPath(urlName: string): string {
    return `${this.baseServerRelativeUrl}/Lists/${urlName}`;
  }

  public listUrl(urlName: string): string {
    return `${this.baseUrl}/Lists/${urlName}`;
  }

  public api(urlName: string): string {
    return `${this.baseUrl}/_api/web/GetList('${quote(this.listPath(urlName))}')`;
  }

  public libraryApi(title: string): string {
    return `${this.baseUrl}/_api/web/lists/getbytitle('${quote(title)}')`;
  }

  /**
   * Creates the list and its columns. Needs the "Manage Lists" permission,
   * so it only ever runs when a site owner presses a Set up button.
   */
  public async ensure(definition: IListDefinition): Promise<void> {
    let listId: string | undefined;
    let existingFields: string[] = [];
    try {
      const found = await this.sp.getJson<{ Id: string }>(`${this.api(definition.urlName)}?$select=Id`);
      listId = found.Id;
      const fields = await this.sp.getJson<{ value: Array<{ InternalName: string }> }>(
        `${this.api(definition.urlName)}/fields?$select=InternalName&$filter=Hidden eq false`
      );
      existingFields = (fields.value || []).map((f) => f.InternalName);
    } catch (err) {
      if (httpStatusOf(err) !== 404) {
        throw err;
      }
    }

    if (!listId) {
      // Created under its URL name first so the address is predictable, then given its display title.
      const created = await this.sp.post<{ Id: string }>(`${this.baseUrl}/_api/web/lists`, {
        Title: definition.urlName,
        Description: definition.description,
        BaseTemplate: 100,
        ContentTypesEnabled: false
      });
      if (!created || !created.Id) {
        throw new Error(`The list "${definition.title}" could not be created.`);
      }
      listId = created.Id;
      await this.sp.merge(`${this.baseUrl}/_api/web/lists(guid'${listId}')`, { Title: definition.title });
    }

    const listApi = `${this.baseUrl}/_api/web/lists(guid'${listId}')`;
    for (const schemaXml of definition.fields) {
      const nameMatch = / Name="([^"]+)"/.exec(schemaXml);
      if (nameMatch && existingFields.indexOf(nameMatch[1]) >= 0) {
        continue;
      }
      await this.sp.post(`${listApi}/fields/createfieldasxml`, {
        parameters: { SchemaXml: schemaXml, Options: ADD_FIELD_OPTIONS }
      });
    }

    // Hardening that is useful but not essential: a failure here must not block setup.
    const optional: Array<() => Promise<void>> = [];
    if (definition.indexAuthor) {
      optional.push(() => this.sp.merge(`${listApi}/fields/getbyinternalnameortitle('Author')`, { Indexed: true }));
    }
    if (definition.indexCreated) {
      optional.push(() => this.sp.merge(`${listApi}/fields/getbyinternalnameortitle('Created')`, { Indexed: true }));
    }
    if (definition.ownItemsOnly) {
      optional.push(() => this.sp.merge(listApi, { WriteSecurity: 2 }));
    }
    for (const step of optional) {
      try {
        await step();
      } catch {
        // Ignored on purpose, see above.
      }
    }
  }

  /** Creates a document library with the given title if it does not exist yet. */
  public async ensureLibrary(title: string, description: string): Promise<void> {
    try {
      await this.sp.getJson(`${this.libraryApi(title)}?$select=Id`);
      return;
    } catch (err) {
      if (httpStatusOf(err) !== 404) {
        throw err;
      }
    }
    await this.sp.post(`${this.baseUrl}/_api/web/lists`, {
      Title: title,
      Description: description,
      BaseTemplate: 101,
      ContentTypesEnabled: false
    });
  }

  public async getItems<T>(
    urlName: string,
    query: string,
    maxItems: number,
    token?: CancelToken
  ): Promise<{ items: T[]; more: boolean }> {
    return this.sp.getAll<T>(`${this.api(urlName)}/items?${query}`, maxItems, { token });
  }

  public async addItem<T extends IListItemBase>(urlName: string, values: Record<string, unknown>): Promise<T> {
    const created = await this.sp.post<T>(`${this.api(urlName)}/items`, values);
    if (!created) {
      throw new Error('SharePoint did not confirm that the item was saved.');
    }
    return created;
  }

  public async updateItem(urlName: string, id: number, values: Record<string, unknown>): Promise<void> {
    await this.sp.merge(`${this.api(urlName)}/items(${id})`, values);
  }

  /** Sends the item to the recycle bin (restorable), rather than deleting it outright. */
  public async recycleItem(urlName: string, id: number): Promise<void> {
    await this.sp.post(`${this.api(urlName)}/items(${id})/recycle()`);
  }
}

/** True when the error means "that list or library is not there". */
export function isMissingList(err: unknown): boolean {
  return httpStatusOf(err) === 404;
}
