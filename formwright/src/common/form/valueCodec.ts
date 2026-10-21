import { IFieldInfo } from '../model/IFieldInfo';
import {
  emptyValue,
  IAttachmentsValue,
  IFormValues,
  ILookupValue,
  ITermValue,
  IUrlValue,
  IUserValue
} from './formValues';

/**
 * Maps between SharePoint REST (JSON light) item payloads and form values.
 * Kept free of HTTP so the mapping can be unit-tested.
 */

/* eslint-disable @typescript-eslint/no-explicit-any */
type RawItem = { [key: string]: any };

export interface ISelectExpand {
  select: string[];
  expand: string[];
}

export function buildSelect(fields: IFieldInfo[]): ISelectExpand {
  const select = ['Id', 'ContentTypeId'];
  const expand: string[] = [];
  fields.forEach((f) => {
    const p = f.entityPropertyName;
    switch (f.kind) {
      case 'Lookup':
      case 'LookupMulti': {
        const display = f.lookupField || 'Title';
        select.push(`${p}/Id`, `${p}/${display}`);
        expand.push(p);
        break;
      }
      case 'User':
      case 'UserMulti':
        select.push(`${p}/Id`, `${p}/Title`, `${p}/EMail`, `${p}/Name`);
        expand.push(p);
        break;
      case 'Attachments':
        select.push('Attachments', 'AttachmentFiles/FileName', 'AttachmentFiles/ServerRelativeUrl');
        expand.push('AttachmentFiles');
        break;
      case 'Unsupported':
        break;
      default:
        select.push(p);
    }
  });
  // Term labels and calculated / computed text come from FieldValuesAsText.
  if (fields.some((f) => f.kind === 'TaxonomyFieldType' || f.kind === 'TaxonomyFieldTypeMulti')) {
    select.push('FieldValuesAsText');
    expand.push('FieldValuesAsText');
  }
  return { select: dedupe(select), expand: dedupe(expand) };
}

function dedupe(values: string[]): string[] {
  return values.filter((v, i) => values.indexOf(v) === i);
}

function asArray(value: any): any[] {
  if (!value) {
    return [];
  }
  if (Array.isArray(value)) {
    return value;
  }
  if (Array.isArray(value.results)) {
    return value.results;
  }
  return [value];
}

/** FieldValuesAsText escapes "_x" in internal names as "_x005f_x". */
function asText(item: RawItem, field: IFieldInfo): string {
  const bag = item.FieldValuesAsText || {};
  const escaped = field.internalName.replace(/_x/g, '_x005f_x');
  const value = bag[escaped] !== undefined ? bag[escaped] : bag[field.internalName];
  return typeof value === 'string' ? value : '';
}

function readLookup(raw: any, field: IFieldInfo): ILookupValue | null {
  if (!raw || raw.Id === undefined || raw.Id === null) {
    return null;
  }
  const display = raw[field.lookupField || 'Title'];
  return { id: raw.Id, title: display === undefined || display === null ? '' : String(display) };
}

function readUser(raw: any): IUserValue | null {
  if (!raw || raw.Id === undefined || raw.Id === null) {
    return null;
  }
  return { id: raw.Id, title: raw.Title || '', email: raw.EMail || '', loginName: raw.Name };
}

export function readValue(field: IFieldInfo, item: RawItem): unknown {
  const raw = field.kind === 'Attachments' ? item.AttachmentFiles : item[field.entityPropertyName];
  switch (field.kind) {
    case 'Text':
    case 'Note':
    case 'Choice':
    case 'Calculated':
    case 'Computed':
      return raw === undefined || raw === null ? null : String(raw);
    case 'MultiChoice':
      return asArray(raw).map(String);
    case 'Number':
    case 'Currency':
    case 'Integer':
    case 'Counter':
      return typeof raw === 'number' ? raw : raw === null || raw === undefined || raw === '' ? null : Number(raw);
    case 'Boolean':
      return raw === true || raw === 1 || raw === '1';
    case 'DateTime':
      return raw ? new Date(raw) : null;
    case 'Lookup':
      return readLookup(raw, field);
    case 'LookupMulti':
      return asArray(raw).map((r) => readLookup(r, field)).filter((v) => v !== null);
    case 'User':
      return readUser(raw);
    case 'UserMulti':
      return asArray(raw).map(readUser).filter((v) => v !== null);
    case 'URL':
      return raw && raw.Url ? ({ url: raw.Url, description: raw.Description || '' } as IUrlValue) : null;
    case 'TaxonomyFieldType':
    case 'TaxonomyFieldTypeMulti': {
      // The REST value carries term ids; labels are only reliable in FieldValuesAsText.
      const terms = asArray(raw).filter((t) => t && t.TermGuid);
      const labels = asText(item, field).split(';').map((s) => s.trim());
      const values: ITermValue[] = terms.map((t, i) => ({
        termGuid: t.TermGuid,
        label: labels[i] || (isNaN(Number(t.Label)) ? t.Label : '') || ''
      }));
      return field.kind === 'TaxonomyFieldType' ? values[0] || null : values;
    }
    case 'Attachments':
      return {
        existing: asArray(raw).map((a) => ({ fileName: a.FileName, serverRelativeUrl: a.ServerRelativeUrl })),
        removed: [],
        added: []
      } as IAttachmentsValue;
    default:
      return raw === undefined ? emptyValue(field) : raw;
  }
}

export function readItem(fields: IFieldInfo[], item: RawItem): IFormValues {
  const values: IFormValues = {};
  fields.forEach((f) => {
    values[f.internalName] = readValue(f, item);
  });
  return values;
}

function toIso(value: unknown): string | null {
  return value instanceof Date && !isNaN(value.getTime()) ? value.toISOString() : null;
}

/** Comparable form of a value, to work out which fields changed. */
function fingerprint(field: IFieldInfo, value: unknown): string {
  const write = writeValue(field, value);
  return JSON.stringify(write ? write.value : null);
}

interface IWrite {
  property: string;
  value: unknown;
}

function writeValue(field: IFieldInfo, value: unknown): IWrite | undefined {
  const p = field.entityPropertyName;
  switch (field.kind) {
    case 'Text':
    case 'Note':
    case 'Choice':
      return { property: p, value: value === undefined || value === null || value === '' ? null : String(value) };
    case 'MultiChoice':
      return { property: p, value: Array.isArray(value) ? value : [] };
    case 'Number':
    case 'Currency':
    case 'Integer':
      return { property: p, value: typeof value === 'number' && !isNaN(value) ? value : null };
    case 'Boolean':
      return { property: p, value: !!value };
    case 'DateTime':
      return { property: p, value: toIso(value) };
    case 'Lookup':
    case 'User':
      return { property: `${p}Id`, value: value ? (value as ILookupValue).id : null };
    case 'LookupMulti':
    case 'UserMulti':
      return { property: `${p}Id`, value: Array.isArray(value) ? (value as ILookupValue[]).map((v) => v.id) : [] };
    case 'URL': {
      const url = value as IUrlValue | null;
      return { property: p, value: url && url.url ? { Url: url.url, Description: url.description || url.url } : null };
    }
    default:
      return undefined;
  }
}

/** Taxonomy values go through validateUpdateListItem, as "Label|guid;Label|guid". */
export function taxonomyFormValue(value: unknown): string {
  const terms = (Array.isArray(value) ? value : value ? [value] : []) as ITermValue[];
  return terms.map((t) => `${t.label}|${t.termGuid}`).join(';');
}

export interface IItemPayload {
  /** JSON body for items POST / MERGE. */
  body: { [property: string]: unknown };
  /** validateUpdateListItem form values for managed metadata fields. */
  taxonomy: { FieldName: string; FieldValue: string }[];
  /** User values that still need resolving to ids before saving. */
  unresolvedUsers: string[];
}

/**
 * Builds the write payload. With `original`, only changed fields are sent, so
 * saving a form never overwrites a column the form did not touch (another
 * form, a flow or a person may own it).
 */
export function buildPayload(fields: IFieldInfo[], values: IFormValues, original?: IFormValues): IItemPayload {
  const payload: IItemPayload = { body: {}, taxonomy: [], unresolvedUsers: [] };
  fields.forEach((f) => {
    if (f.readOnly || !(f.internalName in values)) {
      return;
    }
    const value = values[f.internalName];
    // A new item with a lookup or person column left blank: leave it out, so a
    // default SharePoint applies server-side isn't overwritten with nothing.
    if (!original && f.defaultValue && (f.kind === 'Lookup' || f.kind === 'LookupMulti' || f.kind === 'User' || f.kind === 'UserMulti') && isBlankValue(value)) {
      return;
    }
    if (original && f.kind !== 'Attachments' && fingerprintAll(f, value) === fingerprintAll(f, original[f.internalName])) {
      return;
    }
    if (f.kind === 'TaxonomyFieldType' || f.kind === 'TaxonomyFieldTypeMulti') {
      payload.taxonomy.push({ FieldName: f.internalName, FieldValue: taxonomyFormValue(value) });
      return;
    }
    if (f.kind === 'User' || f.kind === 'UserMulti') {
      const users = (Array.isArray(value) ? value : value ? [value] : []) as IUserValue[];
      users.filter((u) => u.unresolved).forEach((u) => payload.unresolvedUsers.push(u.email || u.title));
    }
    const write = writeValue(f, value);
    if (write) {
      payload.body[write.property] = write.value;
    }
  });
  return payload;
}

function isBlankValue(value: unknown): boolean {
  return value === null || value === undefined || (Array.isArray(value) && value.length === 0);
}

function fingerprintAll(field: IFieldInfo, value: unknown): string {
  if (field.kind === 'TaxonomyFieldType' || field.kind === 'TaxonomyFieldTypeMulti') {
    return taxonomyFormValue(value);
  }
  return fingerprint(field, value);
}
