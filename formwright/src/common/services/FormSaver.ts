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
  /**
   * Work already done by an earlier attempt that failed part-way. Saving
   * records each step here, so pressing Save again after (say) an attachment
   * upload failed updates the item it already created instead of creating a
   * second one, and doesn't repeat uploads or deletions.
   */
  progress?: ISaveProgress;
}

export interface ISaveProgress {
  /** Id of the item a previous attempt created. */
  createdItemId?: number;
  uploaded: string[];
  deleted: string[];
}

export function newSaveProgress(): ISaveProgress {
  return { uploaded: [], deleted: [] };
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
  const progress = request.progress || newSaveProgress();
  const values = await resolveUsers(listService, fields, request.values);
  const creating = (request.mode === 'New' || !request.itemId) && !progress.createdItemId;
  // A retry of a create writes every field again, onto the item already created.
  const payload = buildPayload(fields, values, request.mode === 'New' || progress.createdItemId ? undefined : request.original);

  let itemId: number;
  if (creating) {
    itemId = await listService.createItem(listId, payload, { contentTypeId: request.contentTypeId, folderPath: request.folderPath });
    progress.createdItemId = itemId;
  } else {
    itemId = progress.createdItemId || (request.itemId as number);
    // The ETag is the one read when the form opened; our own earlier attempt changed it.
    const etag = request.detectConflicts && request.etag && !progress.createdItemId && !progress.uploaded.length && !progress.deleted.length ? request.etag : '*';
    try {
      await listService.updateItem(listId, itemId, payload, etag);
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
      if (progress.deleted.indexOf(name) < 0) {
        await listService.deleteAttachment(listId, itemId, name);
        progress.deleted.push(name);
      }
    }
    for (const file of attachments.added) {
      if (progress.uploaded.indexOf(file.name) < 0) {
        await listService.addAttachment(listId, itemId, file);
        progress.uploaded.push(file.name);
      }
    }
  }
  return itemId;
}
