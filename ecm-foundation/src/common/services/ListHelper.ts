import { SpRestClient } from './SpRestClient';

export interface IFieldDefinition {
  /** Internal name the code reads. */
  name: string;
  /** Field schema XML; Name must equal `name`. */
  xml: string;
}

export interface IFieldValue {
  FieldName: string;
  FieldValue: string;
}

interface IValidateResult {
  value?: { FieldName: string; ErrorMessage: string | null; HasException: boolean }[];
}

export interface IListInfo {
  Id: string;
  RootFolder: { ServerRelativeUrl: string };
}

export function odataString(value: string): string {
  return value.replace(/'/g, "''");
}

/**
 * Creates and fills SharePoint lists through REST as the signed-in user,
 * without entity type names (untyped OData 4 payloads and the
 * AddValidateUpdateItemUsingPath display-format API).
 */
export class ListHelper {
  constructor(private rest: SpRestClient, private webUrl: string, private listTitle: string) {}

  public get listApi(): string {
    return `${this.webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(odataString(this.listTitle))}')`;
  }

  /** The list, or undefined when it does not exist (or the user cannot see it). */
  public async getList(): Promise<IListInfo | undefined> {
    try {
      return await this.rest.getJson<IListInfo>(`${this.listApi}?$select=Id,RootFolder/ServerRelativeUrl&$expand=RootFolder`);
    } catch (err) {
      if ((err as { status?: number }).status === 404) {
        return undefined;
      }
      throw err;
    }
  }

  /**
   * Creates the list if needed and adds any missing columns. Option 8
   * (AddFieldInternalNameHint) keeps the internal name from the XML instead
   * of deriving one from the display name; 16 adds the column to the view.
   * Returns whether the list was created.
   */
  public async ensure(description: string, fields: IFieldDefinition[], log: (line: string) => void): Promise<{ list: IListInfo; created: boolean }> {
    let list = await this.getList();
    const created = !list;
    if (!list) {
      await this.rest.postJson(`${this.webUrl}/_api/web/lists`, { Title: this.listTitle, BaseTemplate: 100, Description: description });
      log(`Created the list "${this.listTitle}".`);
      list = await this.getList();
      if (!list) {
        throw new Error('The list was created but could not be read back.');
      }
    }
    const existing = ((await this.rest.getJson<{ value: { InternalName: string }[] }>(`${this.listApi}/fields?$select=InternalName`)).value || []).map(
      (f) => f.InternalName
    );
    for (const field of fields) {
      if (existing.indexOf(field.name) < 0) {
        await this.rest.postJson(`${this.listApi}/fields/CreateFieldAsXml`, { parameters: { SchemaXml: field.xml, Options: 8 | 16 } });
        log(`Added column ${field.name}.`);
      }
    }
    return { list, created };
  }

  public async addItem(folderUrl: string, values: IFieldValue[]): Promise<void> {
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
}
