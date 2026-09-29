import { CellValue, IBoardColumn, IPerson, ITimelineValue, IWorkItem, ColumnType } from '../models/types';
import { fromSpDate, toSpDate } from './dates';

/** Fields every board list has. Cross-board features (My Work) rely on these names. */
export const F = {
  GroupId: 'WB_GroupId',
  SortOrder: 'WB_SortOrder',
  ParentId: 'WB_ParentId',
  Owner: 'WB_Owner',
  Status: 'WB_Status',
  StartDate: 'WB_StartDate',
  DueDate: 'WB_DueDate'
};

/** Fields on each board's Updates list. */
export const UF = {
  ItemId: 'WB_ItemId',
  ParentId: 'WB_ParentId',
  Body: 'WB_Body',
  Mentions: 'WB_Mentions',
  LikedBy: 'WB_LikedBy'
};

export function xmlEscape(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function attrs(name: string, displayName: string, extra: string): string {
  return `Name="${name}" StaticName="${name}" DisplayName="${xmlEscape(displayName)}" ${extra}`.trim();
}

/** SharePoint field schema XML for one field. */
export function fieldXml(
  kind: ColumnType | 'sortnumber' | 'id',
  name: string,
  displayName: string,
  options: { indexed?: boolean; choices?: string[] } = {}
): string {
  const indexed = options.indexed ? ' Indexed="TRUE"' : '';
  switch (kind) {
    case 'status': {
      const choices = (options.choices || []).map(c => `<CHOICE>${xmlEscape(c)}</CHOICE>`).join('');
      return `<Field Type="Choice" ${attrs(name, displayName, 'Format="Dropdown" FillInChoice="TRUE"' + indexed)}><CHOICES>${choices}</CHOICES></Field>`;
    }
    case 'dropdown': {
      const choices = (options.choices || []).map(c => `<CHOICE>${xmlEscape(c)}</CHOICE>`).join('');
      return `<Field Type="MultiChoice" ${attrs(name, displayName, 'FillInChoice="TRUE"')}><CHOICES>${choices}</CHOICES></Field>`;
    }
    case 'people':
      return `<Field Type="UserMulti" ${attrs(name, displayName, 'Mult="TRUE" UserSelectionMode="PeopleOnly" ShowField="ImnName"')} />`;
    case 'text':
    case 'link':
      return `<Field Type="Text" ${attrs(name, displayName, 'MaxLength="255"' + indexed)} />`;
    case 'longtext':
      return `<Field Type="Note" ${attrs(name, displayName, 'NumLines="6" RichText="FALSE" UnlimitedLengthInDocumentLibrary="TRUE"')} />`;
    case 'number':
    case 'sortnumber':
    case 'id':
      return `<Field Type="Number" ${attrs(name, displayName, indexed.trim())} />`;
    case 'date':
    case 'timeline':
      return `<Field Type="DateTime" ${attrs(name, displayName, 'Format="DateOnly"' + indexed)} />`;
    case 'checkbox':
      return `<Field Type="Boolean" ${attrs(name, displayName, '')}><Default>0</Default></Field>`;
    default:
      throw new Error('Unknown field kind ' + kind);
  }
}

/** Internal field names a column occupies. */
export function columnFields(col: IBoardColumn): string[] {
  return col.type === 'timeline' && col.fieldEnd ? [col.field, col.fieldEnd] : [col.field];
}

/** $select and $expand parts for loading items with the given columns. */
export function itemQueryParts(columns: IBoardColumn[]): { select: string[]; expand: string[] } {
  const select = ['Id', 'Title', F.GroupId, F.SortOrder, F.ParentId, 'Created', 'Modified', 'Attachments',
    'Author/Id', 'Author/Title', 'Author/EMail', 'Editor/Id', 'Editor/Title', 'Editor/EMail'];
  const expand = ['Author', 'Editor'];
  columns.forEach(col => {
    if (col.type === 'people') {
      select.push(`${col.field}/Id`, `${col.field}/Title`, `${col.field}/EMail`);
      expand.push(col.field);
    } else {
      columnFields(col).forEach(f => {
        if (select.indexOf(f) < 0) {
          select.push(f);
        }
      });
    }
  });
  return { select, expand };
}

/* eslint-disable @typescript-eslint/no-explicit-any */
function readPerson(raw: any): IPerson | null {
  if (!raw || raw.Id === undefined) {
    return null;
  }
  return { id: raw.Id, title: raw.Title || '', email: raw.EMail || '' };
}

function readPeople(raw: any): IPerson[] {
  if (!raw) {
    return [];
  }
  const arr: any[] = Array.isArray(raw) ? raw : raw.results || [raw];
  return arr.map(readPerson).filter((p): p is IPerson => p !== null);
}

function readStrings(raw: any): string[] {
  if (!raw) {
    return [];
  }
  if (Array.isArray(raw)) {
    return raw.filter(s => typeof s === 'string' && s.length > 0);
  }
  if (Array.isArray(raw.results)) {
    return raw.results;
  }
  if (typeof raw === 'string') {
    // Some endpoints return ";#a;#b;#".
    return raw.split(';#').filter((s: string) => s.length > 0);
  }
  return [];
}

export function readCell(col: IBoardColumn, row: any): CellValue {
  const raw = row[col.field];
  switch (col.type) {
    case 'people':
      return readPeople(raw);
    case 'dropdown':
      return readStrings(raw);
    case 'number':
      return typeof raw === 'number' ? raw : raw === null || raw === undefined || raw === '' ? null : Number(raw);
    case 'checkbox':
      return raw === true || raw === 1 || raw === '1' || raw === 'Yes';
    case 'date':
      return fromSpDate(raw);
    case 'timeline': {
      const v: ITimelineValue = { start: fromSpDate(raw), end: fromSpDate(col.fieldEnd ? row[col.fieldEnd] : null) };
      return v;
    }
    default:
      return raw === undefined || raw === null ? null : String(raw);
  }
}

export function readItem(columns: IBoardColumn[], row: any): IWorkItem {
  const values: { [id: string]: CellValue } = {};
  columns.forEach(col => {
    values[col.id] = readCell(col, row);
  });
  return {
    id: row.Id,
    title: row.Title || '',
    groupId: row[F.GroupId] || '',
    sortOrder: typeof row[F.SortOrder] === 'number' ? row[F.SortOrder] : 0,
    parentId: typeof row[F.ParentId] === 'number' && row[F.ParentId] > 0 ? row[F.ParentId] : null,
    values,
    created: row.Created,
    modified: row.Modified,
    author: readPerson(row.Author),
    editor: readPerson(row.Editor),
    attachments: row.Attachments === true,
    etag: row['odata.etag'] || row['@odata.etag'] || (row.__metadata && row.__metadata.etag) || '*'
  };
}

/** Verbose-OData payload properties for writing one cell. */
export function writeCell(col: IBoardColumn, value: CellValue): { [key: string]: any } {
  switch (col.type) {
    case 'people': {
      const ids = ((value as IPerson[]) || []).map(p => p.id);
      return { [col.field + 'Id']: { results: ids } };
    }
    case 'dropdown':
      return { [col.field]: { __metadata: { type: 'Collection(Edm.String)' }, results: (value as string[]) || [] } };
    case 'number':
      return { [col.field]: value === null || value === undefined || value === '' ? null : Number(value) };
    case 'checkbox':
      return { [col.field]: value === true };
    case 'date':
      return { [col.field]: toSpDate(value as string | null) };
    case 'timeline': {
      const tv = (value as ITimelineValue) || { start: null, end: null };
      const out: { [key: string]: any } = { [col.field]: toSpDate(tv.start) };
      if (col.fieldEnd) {
        out[col.fieldEnd] = toSpDate(tv.end);
      }
      return out;
    }
    default:
      return { [col.field]: value === null || value === undefined ? null : String(value) };
  }
}
/* eslint-enable @typescript-eslint/no-explicit-any */

/** Plain text for a cell, used by search, CSV export and the activity log. */
export function cellText(col: IBoardColumn, value: CellValue): string {
  if (value === null || value === undefined) {
    return '';
  }
  switch (col.type) {
    case 'people':
      return (value as IPerson[]).map(p => p.title).join(', ');
    case 'dropdown':
      return (value as string[]).join(', ');
    case 'checkbox':
      return value ? 'Yes' : '';
    case 'timeline': {
      const tv = value as ITimelineValue;
      if (!tv.start && !tv.end) {
        return '';
      }
      return (tv.start || '') + ' – ' + (tv.end || '');
    }
    case 'number':
      return typeof value === 'number' ? String(value) : '';
    default:
      return String(value);
  }
}

export function isEmptyCell(col: IBoardColumn, value: CellValue): boolean {
  if (col.type === 'checkbox') {
    return value !== true;
  }
  return cellText(col, value) === '';
}
