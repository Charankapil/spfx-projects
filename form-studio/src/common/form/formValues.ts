import { isBlank, toBool, toPrimitive, toText } from '../expressions';
import { IFieldInfo } from '../model/IFieldInfo';

/**
 * In-memory representation of field values while a form is open. Each kind
 * of field has one shape, so controls, rules and the REST mapping agree:
 *
 *   Text, Note, Choice   string | null
 *   MultiChoice          string[]
 *   Number, Currency     number | null
 *   Boolean              boolean
 *   DateTime             Date | null
 *   Lookup / User        ILookupValue | IUserValue | null
 *   *Multi               arrays of the above
 *   URL                  IUrlValue | null
 *   Taxonomy             ITermValue | ITermValue[]
 *   Attachments          IAttachmentsValue
 */

export interface ILookupValue {
  id: number;
  title: string;
}

export interface IUserValue extends ILookupValue {
  email: string;
  loginName?: string;
  /** Set when the value came from an expression (e.g. an e-mail) and still has to be resolved to a site user id. */
  unresolved?: boolean;
}

export interface IUrlValue {
  url: string;
  description: string;
}

export interface ITermValue {
  label: string;
  termGuid: string;
  /** "Parent > Child" path, shown in the picker only. */
  path?: string;
}

export interface IAttachment {
  fileName: string;
  serverRelativeUrl: string;
}

export interface IAttachmentsValue {
  existing: IAttachment[];
  /** File names of existing attachments to delete on save. */
  removed: string[];
  added: File[];
}

export interface IFormValues {
  [internalName: string]: unknown;
}

export function emptyValue(field: IFieldInfo): unknown {
  switch (field.kind) {
    case 'MultiChoice':
    case 'LookupMulti':
    case 'UserMulti':
    case 'TaxonomyFieldTypeMulti':
      return [];
    case 'Boolean':
      return false;
    case 'Attachments':
      return { existing: [], removed: [], added: [] } as IAttachmentsValue;
    default:
      return null;
  }
}

function toDateValue(value: unknown): Date | null {
  const v = toPrimitive(value);
  if (isBlank(v)) {
    return null;
  }
  const d = v instanceof Date ? v : new Date(v as string);
  return isNaN(d.getTime()) ? null : d;
}

function toNumberValue(value: unknown): number | null {
  const v = toPrimitive(value);
  if (isBlank(v)) {
    return null;
  }
  const n = typeof v === 'number' ? v : Number(String(v).replace(/,/g, ''));
  return isNaN(n) ? null : n;
}

function toList(value: unknown): unknown[] {
  if (isBlank(value)) {
    return [];
  }
  if (Array.isArray(value)) {
    return value;
  }
  if (typeof value === 'string') {
    return value.split(';').map((s) => s.trim()).filter((s) => s !== '');
  }
  return [value];
}

function toLookup(value: unknown): ILookupValue | null {
  if (isBlank(value)) {
    return null;
  }
  if (typeof value === 'object' && value !== null && typeof (value as ILookupValue).id === 'number') {
    return value as ILookupValue;
  }
  if (typeof value === 'number') {
    return { id: value, title: '' };
  }
  const n = Number(value);
  return isNaN(n) ? null : { id: n, title: '' };
}

function toUser(value: unknown): IUserValue | null {
  if (isBlank(value)) {
    return null;
  }
  if (typeof value === 'object' && value !== null && typeof (value as IUserValue).id === 'number') {
    return value as IUserValue;
  }
  if (typeof value === 'number') {
    return { id: value, title: '', email: '', unresolved: true };
  }
  const text = toText(value);
  return { id: 0, title: text, email: text, unresolved: true };
}

/**
 * Converts whatever an expression produced (or a SharePoint default value)
 * into the field's value shape. Used for default values, calculated values
 * and assignments.
 */
export function coerceToField(field: IFieldInfo, value: unknown): unknown {
  switch (field.kind) {
    case 'Text':
    case 'Note':
    case 'Choice': {
      const v = toPrimitive(value);
      return isBlank(v) ? null : toText(v);
    }
    case 'MultiChoice':
      return toList(toPrimitive(value)).map(toText);
    case 'Number':
    case 'Currency':
    case 'Integer':
      return toNumberValue(value);
    case 'Boolean':
      return toBool(value);
    case 'DateTime':
      return toDateValue(value);
    case 'Lookup':
      return toLookup(Array.isArray(value) ? value[0] : value);
    case 'LookupMulti':
      return toList(value).map(toLookup).filter((v): v is ILookupValue => v !== null);
    case 'User':
      return toUser(Array.isArray(value) ? value[0] : value);
    case 'UserMulti':
      return toList(value).map(toUser).filter((v): v is IUserValue => v !== null);
    case 'URL': {
      if (isBlank(value)) {
        return null;
      }
      if (typeof value === 'object' && value !== null && (value as IUrlValue).url !== undefined) {
        return value;
      }
      const url = toText(value);
      return { url, description: url } as IUrlValue;
    }
    case 'TaxonomyFieldType':
    case 'TaxonomyFieldTypeMulti': {
      const terms = toList(value).filter(
        (t): t is ITermValue => typeof t === 'object' && t !== null && typeof (t as ITermValue).termGuid === 'string'
      );
      return field.kind === 'TaxonomyFieldType' ? terms[0] || null : terms;
    }
    default:
      return value;
  }
}

/**
 * SharePoint stores column defaults as text: "1" for Yes, "[today]" for
 * today's date, a choice name, a number. Lookup, person and term defaults
 * need ids SharePoint itself fills in server-side, so they are left blank.
 */
export function sharePointDefault(field: IFieldInfo): unknown {
  const raw = field.defaultValue;
  if (raw === undefined || raw === null || raw === '') {
    return emptyValue(field);
  }
  switch (field.kind) {
    case 'Boolean':
      return raw === '1' || raw.toLowerCase() === 'true';
    case 'DateTime': {
      if (raw.toLowerCase() === '[today]') {
        const now = new Date();
        return new Date(now.getFullYear(), now.getMonth(), now.getDate());
      }
      return toDateValue(raw);
    }
    case 'MultiChoice':
      return raw.split(';#').filter((s) => s !== '');
    case 'Text':
    case 'Note':
    case 'Choice':
    case 'Number':
    case 'Currency':
    case 'Integer':
    case 'URL':
      return coerceToField(field, raw);
    default:
      return emptyValue(field);
  }
}

/** Display text for read-only rendering and exports. */
export function displayText(field: IFieldInfo, value: unknown): string {
  if (isBlank(value)) {
    return '';
  }
  switch (field.kind) {
    case 'Boolean':
      return value ? 'Yes' : 'No';
    case 'DateTime': {
      const d = value as Date;
      return field.dateOnly ? d.toLocaleDateString() : `${d.toLocaleDateString()} ${d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}`;
    }
    case 'Number':
    case 'Integer':
      if (field.showAsPercentage) {
        return `${Math.round((value as number) * 10000) / 100}%`;
      }
      return (value as number).toLocaleString(undefined, { maximumFractionDigits: field.decimals !== undefined ? field.decimals : 10 });
    case 'Currency':
      return (value as number).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    case 'URL':
      return (value as IUrlValue).description || (value as IUrlValue).url;
    case 'Attachments':
      return (value as IAttachmentsValue).existing.map((a) => a.fileName).join(', ');
    default:
      return toText(value);
  }
}

export function valueIsBlank(field: IFieldInfo, value: unknown): boolean {
  if (field.kind === 'Attachments') {
    const a = value as IAttachmentsValue | undefined;
    return !a || a.existing.length - a.removed.length + a.added.length <= 0;
  }
  if (field.kind === 'URL') {
    return isBlank(value) || !(value as IUrlValue).url;
  }
  return isBlank(value);
}
