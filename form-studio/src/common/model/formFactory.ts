import { IFieldInfo } from './IFieldInfo';
import {
  ALL_MODES,
  ControlType,
  IControl,
  IFormDefinition,
  IFormPage,
  IFormSettings,
  ISection
} from './IFormDefinition';

export function newId(prefix: string): string {
  return `${prefix}_${Math.random().toString(36).substring(2, 10)}`;
}

export function defaultSettings(): IFormSettings {
  return {
    labelPosition: 'top',
    maxWidth: 960,
    showDefaultButtons: true,
    submitLabel: 'Save',
    cancelLabel: 'Cancel',
    successMessage: 'Thank you - your response has been saved.',
    onSave: [],
    validations: [],
    detectConflicts: true
  };
}

export function createControl(type: ControlType, field?: IFieldInfo): IControl {
  const control: IControl = { id: newId('c'), type, span: 1 };
  switch (type) {
    case 'field':
      control.field = field ? field.internalName : undefined;
      if (field && (field.kind === 'Note' || field.kind === 'Attachments')) {
        control.span = 2;
      }
      break;
    case 'heading':
      control.content = 'Heading';
      control.level = 2;
      control.span = 12;
      break;
    case 'text':
      control.content = 'Add instructions or information for the person filling in the form.';
      control.span = 12;
      break;
    case 'divider':
      control.span = 12;
      break;
    case 'spacer':
      control.height = 24;
      control.span = 12;
      break;
    case 'image':
      control.imageUrl = '';
      control.altText = '';
      break;
    case 'button':
      control.label = 'Approve';
      control.action = 'save';
      control.buttonStyle = 'primary';
      control.assignments = [];
      break;
    case 'relatedItems':
      control.label = 'Related items';
      control.span = 12;
      control.relatedItems = {
        listId: '',
        parentLookupField: '',
        columns: [],
        allowAdd: true,
        allowEdit: true,
        allowDelete: true
      };
      break;
  }
  return control;
}

export function createSection(columns = 2): ISection {
  return { id: newId('s'), title: '', columns, controls: [] };
}

export function createPage(title: string): IFormPage {
  return { id: newId('p'), title, sections: [createSection()] };
}

/**
 * The fields a generated form starts with: what the default SharePoint form
 * would show, in list order.
 */
export function defaultFormFields(fields: IFieldInfo[]): IFieldInfo[] {
  return fields.filter(
    (f) =>
      !f.readOnly &&
      f.kind !== 'Unsupported' &&
      f.internalName !== 'ContentType' &&
      f.internalName !== 'FileLeafRef'
  );
}

/** A working form with every editable field, in a two-column section. */
export function createDefaultForm(name: string, fields: IFieldInfo[]): IFormDefinition {
  const section = createSection(2);
  section.controls = defaultFormFields(fields).map((f) => createControl('field', f));
  return {
    id: newId('f'),
    name,
    modes: ALL_MODES.slice(),
    audience: [],
    contentTypes: [],
    layout: 'sections',
    pages: [{ id: newId('p'), title: 'Page 1', sections: [section] }],
    settings: defaultSettings()
  };
}

/** Deep copy with fresh ids, for "Duplicate form". */
export function cloneForm(form: IFormDefinition, name: string): IFormDefinition {
  const copy: IFormDefinition = JSON.parse(JSON.stringify(form));
  copy.id = newId('f');
  copy.name = name;
  copy.pages.forEach((p) => {
    p.id = newId('p');
    p.sections.forEach((s) => {
      s.id = newId('s');
      s.controls.forEach((c) => {
        c.id = newId('c');
      });
    });
  });
  return copy;
}
