/**
 * A form is stored as plain JSON (see FormStore) so it can be exported,
 * imported and diffed. Every behaviour that Plumsail or Power Apps would put
 * in JavaScript or Power Fx is an expression string evaluated by the safe
 * expression engine in ../expressions - nothing in a form definition is ever
 * executed as code.
 */

export type FormMode = 'New' | 'Edit' | 'Display';

export const ALL_MODES: FormMode[] = ['New', 'Edit', 'Display'];

export type FormLayoutKind = 'sections' | 'tabs' | 'wizard';

export type ControlType =
  | 'field'
  | 'heading'
  | 'text'
  | 'divider'
  | 'spacer'
  | 'image'
  | 'button'
  | 'relatedItems';

export type FieldVariant =
  | 'default'
  | 'dropdown'
  | 'radio'
  | 'checkboxes'
  | 'toggle'
  | 'checkbox'
  | 'multiline';

export interface IValidationRule {
  /** Must evaluate to true for the value to be valid. */
  expression: string;
  message: string;
}

export interface IFieldAssignment {
  /** Internal name of the field to set. */
  field: string;
  /** Expression whose result is written to the field. */
  expression: string;
}

export interface ILookupCascade {
  /** Field in the lookup's source list to filter on. */
  sourceField: string;
  /** Field on this form whose value the source field must equal. */
  formField: string;
}

export interface IRelatedItemsConfig {
  /** Id (GUID) of the child list. */
  listId: string;
  listTitle?: string;
  /** Internal name of the lookup field in the child list that points at this list. */
  parentLookupField: string;
  /** Internal names of the child list fields to show as columns. */
  columns: string[];
  allowAdd: boolean;
  allowEdit: boolean;
  allowDelete: boolean;
}

export type ButtonAction = 'save' | 'close' | 'link' | 'print';

export interface IControl {
  id: string;
  type: ControlType;
  /** Number of grid columns this control spans inside its section. */
  span: number;

  /** type = field: internal name of the bound list field. */
  field?: string;
  /** Overrides the field's display name. Headings, buttons and related items use it as their caption. */
  label?: string;
  description?: string;
  placeholder?: string;
  variant?: FieldVariant;
  /** Number of visible lines for multi-line text. */
  rows?: number;

  /** Expressions. Empty / undefined means "use the default". */
  visibleWhen?: string;
  requiredWhen?: string;
  readOnlyWhen?: string;
  /** Initial value for new items. */
  defaultValue?: string;
  /** Recalculated on every change; makes the control read-only. */
  calculatedValue?: string;
  validations?: IValidationRule[];

  /** Lookup fields. */
  lookupCascade?: ILookupCascade;
  /** Lookup fields: extra OData $filter. `{{ expression }}` placeholders are replaced with form values. */
  lookupFilter?: string;

  /** heading / text: content. `{{ expression }}` placeholders are replaced with form values. */
  content?: string;
  /** heading: 1-4. */
  level?: number;

  /** image */
  imageUrl?: string;
  altText?: string;
  imageWidth?: number;

  /** button */
  action?: ButtonAction;
  /** button: fields set before saving, e.g. Status = "Approved". */
  assignments?: IFieldAssignment[];
  /** button action=link, or redirect after save. Supports `{{ expression }}` placeholders. */
  url?: string;
  buttonStyle?: 'primary' | 'default';
  /** button: skip validation (e.g. "Save draft"). */
  skipValidation?: boolean;

  /** spacer height in px. */
  height?: number;

  relatedItems?: IRelatedItemsConfig;
}

export interface ISection {
  id: string;
  title?: string;
  description?: string;
  columns: number;
  collapsible?: boolean;
  collapsedByDefault?: boolean;
  visibleWhen?: string;
  controls: IControl[];
}

export interface IFormPage {
  id: string;
  title: string;
  visibleWhen?: string;
  sections: ISection[];
}

export interface IFormSettings {
  /** Header shown above the form. Supports `{{ expression }}` placeholders. */
  title?: string;
  subtitle?: string;
  labelPosition: 'top' | 'left';
  maxWidth: number;
  accentColor?: string;
  /** Show the standard Save / Cancel bar. Turn off to use button controls only. */
  showDefaultButtons: boolean;
  submitLabel: string;
  cancelLabel: string;
  /** Shown by the Form web part after a successful save. */
  successMessage: string;
  /** Where to go after saving. Supports `{{ expression }}`; [ID] is the saved item's id. */
  redirectUrl?: string;
  /** Fields set on every save, before the item is written. */
  onSave: IFieldAssignment[];
  /** Form-level rules, checked on submit after field validation. */
  validations: IValidationRule[];
  /** Detect and refuse to overwrite changes made by someone else since the form opened. */
  detectConflicts: boolean;
}

export interface IFormDefinition {
  id: string;
  name: string;
  /** Modes this form is used for. */
  modes: FormMode[];
  /** SharePoint group names. Empty = everyone. */
  audience: string[];
  /** Content type id prefixes this form applies to. Empty = all content types. */
  contentTypes: string[];
  layout: FormLayoutKind;
  pages: IFormPage[];
  settings: IFormSettings;
  modified?: string;
  modifiedBy?: string;
}

/** Everything saved for one list. */
export interface IListForms {
  schemaVersion: number;
  listId: string;
  listTitle: string;
  /** Evaluated top to bottom - the first form whose mode, audience and content type match is used. */
  forms: IFormDefinition[];
}

export const FORMS_SCHEMA_VERSION = 1;
