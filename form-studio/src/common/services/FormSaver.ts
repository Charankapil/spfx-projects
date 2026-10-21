import { FormMode } from '../model/IFormDefinition';
import { IFieldInfo } from '../model/IFieldInfo';
import { IAttachmentsValue, IFormValues, IUserValue } from '../form/formValues';
import { buildPayload } from '../form/valueCodec';
import { HttpError } from './http';
import { ListService } from './ListService';

export interface ISaveRequest {
  listService: ListService;
  listId: string;
  mode: FormMode;
  itemId?: number;
  /** Fields the form may write (FormEngine.writableFields). */
  fields: IFieldInfo[];
  values: IFormValues;
  /** Values as loaded; only changed fields are written. Undefined for new items. */
  original?: IFormValues;
  etag?: string;
  detectConflicts: boolean;
  contentTypeId?: string;
  folderPath?: string;
}

export class ConflictError extends Error {
  constructor() {
    super('Someone else changed this item after you opened it. Reload the form to see their changes, then make yours again.');
    this.name = 'ConflictError';
  }
}

/** People set by expressions (e.g. an e-mail) only become writable once SharePoint gives them an id. */
async function resolveUsers(listService: ListService, fields: IFieldInfo[], values: IFormValues): Promise<IFormValues> {
  const next: IFormValues = { ...values };
  for (const f of fields) {
    if (f.kind !== 'User' && f.kind !== 'UserMulti') {
      continue;
    }
    const raw = values[f.internalName];
    const users = (Array.isArray(raw) ? raw : raw ? [raw] : []) as IUserValue[];
    if (!users.some((u) => u.unresolved)) {
      continue;
    }
    const resolved: IUserValue[] = [];
    for (const u of users) {
      if (!u.unresolved) {
        resolved.push(u);
      } else if (u.id) {
        resolved.push({ ...u, unresolved: false });
      } else if (u.email || u.title) {
        resolved.push(await listService.ensureUser(u.email || u.title));
      }
    }
    next[f.internalName] = f.kind === 'User' ? resolved[0] || null : resolved;
  }
  return next;
}

/** Writes the item (create or update), then its attachments. Returns the item id. */
export async function saveItem(request: ISaveRequest): Promise<number> {
  const { listService, listId, fields } = request;
  const values = await resolveUsers(listService, fields, request.values);
  const payload = buildPayload(fields, values, request.mode === 'New' ? undefined : request.original);

  let itemId: number;
  if (request.mode === 'New' || !request.itemId) {
    itemId = await listService.createItem(listId, payload, { contentTypeId: request.contentTypeId, folderPath: request.folderPath });
  } else {
    itemId = request.itemId;
    try {
      await listService.updateItem(listId, itemId, payload, request.detectConflicts && request.etag ? request.etag : '*');
    } catch (e) {
      // Status check rather than instanceof: Error subclasses lose their prototype when compiled to ES5.
      if ((e as HttpError).status === 412) {
        throw new ConflictError();
      }
      throw e;
    }
  }

  const attachmentsField = fields.filter((f) => f.kind === 'Attachments')[0];
  const attachments = attachmentsField ? (values[attachmentsField.internalName] as IAttachmentsValue | undefined) : undefined;
  if (attachments) {
    for (const name of attachments.removed) {
      await listService.deleteAttachment(listId, itemId, name);
    }
    for (const file of attachments.added) {
      await listService.addAttachment(listId, itemId, file);
    }
  }
  return itemId;
}
