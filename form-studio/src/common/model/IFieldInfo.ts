/**
 * The subset of a SharePoint field's schema the designer and the renderer
 * need. Built from `_api/web/lists(...)/fields`, so it describes the list as
 * it is right now; the form definition only stores internal names and looks
 * the rest up here at run time.
 */
export type FieldKind =
  | 'Text'
  | 'Note'
  | 'Number'
  | 'Currency'
  | 'Integer'
  | 'Choice'
  | 'MultiChoice'
  | 'Boolean'
  | 'DateTime'
  | 'Lookup'
  | 'LookupMulti'
  | 'User'
  | 'UserMulti'
  | 'URL'
  | 'TaxonomyFieldType'
  | 'TaxonomyFieldTypeMulti'
  | 'Attachments'
  | 'Calculated'
  | 'Computed'
  | 'Counter'
  | 'Unsupported';

export interface IFieldInfo {
  internalName: string;
  /** Property name used in REST payloads; differs from the internal name for some fields (e.g. names with _x0020_). */
  entityPropertyName: string;
  title: string;
  kind: FieldKind;
  /** Raw TypeAsString from SharePoint, kept for diagnostics and unsupported types. */
  typeAsString: string;
  description: string;
  required: boolean;
  /** Never editable through a form: calculated, computed, system fields. */
  readOnly: boolean;
  defaultValue?: string;
  choices?: string[];
  fillIn?: boolean;
  maxLength?: number;
  min?: number;
  max?: number;
  decimals?: number;
  showAsPercentage?: boolean;
  richText?: boolean;
  appendOnly?: boolean;
  dateOnly?: boolean;
  lookupListId?: string;
  lookupField?: string;
  /** User fields: 0 = people only, 1 = people and groups. */
  userSelectionMode?: number;
  /** User fields: limit to members of this SharePoint group (0 = everyone). */
  userSelectionGroup?: number;
  /** URL fields: 0 = hyperlink, 1 = picture. */
  urlFormat?: number;
  termSetId?: string;
  anchorId?: string;
}

/** Fields a form can write to. */
export function isEditableKind(kind: FieldKind): boolean {
  return (
    kind !== 'Calculated' && kind !== 'Computed' && kind !== 'Counter' && kind !== 'Unsupported'
  );
}

export function isMultiValueKind(kind: FieldKind): boolean {
  return (
    kind === 'MultiChoice' ||
    kind === 'LookupMulti' ||
    kind === 'UserMulti' ||
    kind === 'TaxonomyFieldTypeMulti'
  );
}
