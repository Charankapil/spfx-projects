import { SpClient, SpError, odataString } from './SpClient';

/** Helpers for creating lists and fields over REST. Every helper checks before it creates. */

export interface IListInfo {
  Id: string;
  Title: string;
}

export async function findList(sp: SpClient, title: string): Promise<IListInfo | null> {
  try {
    return await sp.get<IListInfo>(`web/lists/getbytitle(${odataString(title)})?$select=Id,Title`);
  } catch (e) {
    if (e instanceof SpError && e.isNotFound) {
      return null;
    }
    throw e;
  }
}

export interface IEnsureListOptions {
  description?: string;
  hidden?: boolean;
  versioning?: boolean;
  attachments?: boolean;
  /** 2 = users read (and edit) only their own items. */
  ownItemsOnly?: boolean;
}

export async function ensureList(sp: SpClient, title: string, options: IEnsureListOptions = {}): Promise<IListInfo> {
  let list = await findList(sp, title);
  if (!list) {
    list = await sp.post<IListInfo>('web/lists', {
      __metadata: { type: 'SP.List' },
      BaseTemplate: 100,
      Title: title,
      Description: options.description || '',
      ContentTypesEnabled: false
    });
    if (!list || !list.Id) {
      list = await findList(sp, title);
    }
    if (!list) {
      throw new Error(`The list "${title}" could not be created.`);
    }
  }
  const settings: { [k: string]: unknown } = { __metadata: { type: 'SP.List' } };
  if (options.hidden !== undefined) {
    settings.Hidden = options.hidden;
  }
  if (options.versioning !== undefined) {
    settings.EnableVersioning = options.versioning;
  }
  if (options.attachments !== undefined) {
    settings.EnableAttachments = options.attachments;
  }
  if (options.ownItemsOnly) {
    settings.ReadSecurity = 2;
    settings.WriteSecurity = 2;
  }
  if (Object.keys(settings).length > 1) {
    await sp.merge(`web/lists(guid'${list.Id}')`, settings);
  }
  return list;
}

export async function fieldExists(sp: SpClient, listId: string, internalName: string): Promise<boolean> {
  try {
    await sp.get(`web/lists(guid'${listId}')/fields/getbyinternalnameortitle(${odataString(internalName)})?$select=InternalName`);
    return true;
  } catch (e) {
    if (e instanceof SpError && e.isNotFound) {
      return false;
    }
    throw e;
  }
}

function xmlUnescape(value: string): string {
  return value
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}

/**
 * Create a field with the given internal name.
 *
 * Options 8 (AddFieldInternalNameHint) makes SharePoint take the internal name from the field's
 * DisplayName, not from Name. So the field is created with DisplayName set to the internal name,
 * then renamed to its friendly title. A failed rename is harmless: the internal name is what the
 * app uses.
 */
export async function ensureField(sp: SpClient, listId: string, internalName: string, schemaXml: string): Promise<void> {
  if (await fieldExists(sp, listId, internalName)) {
    return;
  }
  const match = /DisplayName="([^"]*)"/.exec(schemaXml);
  const friendly = match ? xmlUnescape(match[1]) : internalName;
  const xml = match ? schemaXml.replace(match[0], `DisplayName="${internalName}"`) : schemaXml;
  await sp.post(`web/lists(guid'${listId}')/fields/createfieldasxml`, {
    parameters: {
      __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' },
      SchemaXml: xml,
      Options: 8
    }
  });
  if (friendly !== internalName) {
    try {
      await renameField(sp, listId, internalName, friendly);
    } catch {
      // Cosmetic only.
    }
  }
}

export async function deleteField(sp: SpClient, listId: string, internalName: string): Promise<void> {
  if (!(await fieldExists(sp, listId, internalName))) {
    return;
  }
  await sp.remove(`web/lists(guid'${listId}')/fields/getbyinternalnameortitle(${odataString(internalName)})`);
}

export async function renameField(sp: SpClient, listId: string, internalName: string, title: string): Promise<void> {
  await sp.merge(`web/lists(guid'${listId}')/fields/getbyinternalnameortitle(${odataString(internalName)})`, {
    __metadata: { type: 'SP.Field' },
    Title: title
  });
}

/** Move a list to the recycle bin. */
export async function recycleList(sp: SpClient, listId: string): Promise<void> {
  try {
    await sp.post(`web/lists(guid'${listId}')/recycle`);
  } catch (e) {
    if (!(e instanceof SpError && e.isNotFound)) {
      throw e;
    }
  }
}

/** Rename the list's Title field (for example to "Item") so the list reads well in SharePoint too. */
export async function setTitleFieldName(sp: SpClient, listId: string, displayName: string): Promise<void> {
  await sp.merge(`web/lists(guid'${listId}')/fields/getbyinternalnameortitle('Title')`, {
    __metadata: { type: 'SP.Field' },
    Title: displayName
  });
}

const entityTypeCache: Map<string, string> = new Map();

/** ListItemEntityTypeFullName, needed as __metadata.type when writing items with verbose OData. */
export async function itemEntityType(sp: SpClient, listId: string): Promise<string> {
  const key = listId.toLowerCase();
  const cached = entityTypeCache.get(key);
  if (cached) {
    return cached;
  }
  const info = await sp.get<{ ListItemEntityTypeFullName: string }>(`web/lists(guid'${listId}')?$select=ListItemEntityTypeFullName`);
  entityTypeCache.set(key, info.ListItemEntityTypeFullName);
  return info.ListItemEntityTypeFullName;
}
