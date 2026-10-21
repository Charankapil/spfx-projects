import {
  evaluate,
  evaluateOr,
  IExpressionEnv,
  isBlank,
  renderTemplate,
  toBool
} from '../expressions';
import { IFieldInfo, isEditableKind } from '../model/IFieldInfo';
import {
  FormMode,
  IControl,
  IFieldAssignment,
  IFormDefinition,
  IFormPage,
  ISection
} from '../model/IFormDefinition';
import {
  coerceToField,
  displayText,
  IFormValues,
  IUrlValue,
  sharePointDefault,
  valueIsBlank
} from './formValues';

export interface IRuntimeContext {
  mode: FormMode;
  user: { name: string; email: string; loginName: string };
  groups: string[];
  itemId?: number;
  now?: () => Date;
}

export interface IControlState {
  visible: boolean;
  required: boolean;
  readOnly: boolean;
}

export interface IFormState {
  controls: { [controlId: string]: IControlState };
  sections: { [sectionId: string]: boolean };
  pages: { [pageId: string]: boolean };
}

export interface IControlLocation {
  control: IControl;
  section: ISection;
  page: IFormPage;
}

/** Errors keyed by control id. */
export interface IValidationErrors {
  [controlId: string]: string;
}

const MAX_CALCULATION_PASSES = 5;

/**
 * Everything a form does between "values changed" and "render": rule
 * evaluation, calculated values, defaults and validation. Pure TypeScript
 * with no SharePoint or React dependency, so the same logic drives the
 * designer preview, the list form and the page web part, and is unit-tested.
 */
export class FormEngine {
  private fieldMap: { [internalName: string]: IFieldInfo } = {};
  private locations: IControlLocation[] = [];

  constructor(public readonly form: IFormDefinition, fields: IFieldInfo[], public readonly runtime: IRuntimeContext) {
    fields.forEach((f) => {
      this.fieldMap[f.internalName] = f;
    });
    form.pages.forEach((page) =>
      page.sections.forEach((section) =>
        section.controls.forEach((control) => this.locations.push({ control, section, page }))
      )
    );
  }

  public field(internalName: string | undefined): IFieldInfo | undefined {
    return internalName ? this.fieldMap[internalName] : undefined;
  }

  public get controls(): IControlLocation[] {
    return this.locations;
  }

  public fieldControls(): { control: IControl; field: IFieldInfo }[] {
    const out: { control: IControl; field: IFieldInfo }[] = [];
    this.locations.forEach((l) => {
      const field = l.control.type === 'field' ? this.field(l.control.field) : undefined;
      if (field) {
        out.push({ control: l.control, field });
      }
    });
    return out;
  }

  public env(values: IFormValues): IExpressionEnv {
    return {
      getField: (name) => values[name],
      hasField: (name) => !!this.fieldMap[name],
      user: this.runtime.user,
      groups: this.runtime.groups,
      mode: this.runtime.mode,
      itemId: this.runtime.itemId,
      now: this.runtime.now
    };
  }

  public label(control: IControl): string {
    const field = this.field(control.field);
    return control.label || (field ? field.title : control.field || '');
  }

  public template(text: string | undefined, values: IFormValues, escape?: (value: string) => string): string {
    return renderTemplate(text, this.env(values), escape);
  }

  public computeState(values: IFormValues): IFormState {
    const env = this.env(values);
    const isDisplay = this.runtime.mode === 'Display';
    const state: IFormState = { controls: {}, sections: {}, pages: {} };
    this.form.pages.forEach((page) => {
      const pageVisible = evaluateOr(page.visibleWhen, env, true, toBool);
      state.pages[page.id] = pageVisible;
      page.sections.forEach((section) => {
        const sectionVisible = pageVisible && evaluateOr(section.visibleWhen, env, true, toBool);
        state.sections[section.id] = sectionVisible;
        section.controls.forEach((control) => {
          let visible = sectionVisible && evaluateOr(control.visibleWhen, env, true, toBool);
          let readOnly = isDisplay;
          let required = false;
          if (control.type === 'field') {
            const field = this.field(control.field);
            if (!field) {
              // The column was deleted from the list after the form was designed.
              visible = false;
            } else {
              readOnly =
                readOnly ||
                field.readOnly ||
                !isEditableKind(field.kind) ||
                !!(control.calculatedValue && control.calculatedValue.trim()) ||
                evaluateOr(control.readOnlyWhen, env, false, toBool);
              required = !readOnly && (field.required || evaluateOr(control.requiredWhen, env, false, toBool));
            }
          }
          state.controls[control.id] = { visible, readOnly, required };
        });
      });
    });
    return state;
  }

  /**
   * Starting values. New items get control defaults, then SharePoint column
   * defaults; existing items keep what was loaded.
   */
  public initialValues(loaded: IFormValues): IFormValues {
    const values: IFormValues = { ...loaded };
    if (this.runtime.mode === 'New') {
      Object.keys(this.fieldMap).forEach((name) => {
        if (!(name in values) || isBlank(values[name])) {
          values[name] = sharePointDefault(this.fieldMap[name]);
        }
      });
      const env = this.env(values);
      this.fieldControls().forEach(({ control, field }) => {
        if (control.defaultValue && control.defaultValue.trim()) {
          try {
            values[field.internalName] = coerceToField(field, evaluate(control.defaultValue, env));
          } catch {
            // A broken default leaves the field at its SharePoint default.
          }
        }
      });
    }
    return this.applyCalculated(values);
  }

  /** Re-runs calculated values until they settle (they can depend on each other). */
  public applyCalculated(values: IFormValues): IFormValues {
    const calculated = this.fieldControls().filter(({ control }) => control.calculatedValue && control.calculatedValue.trim());
    if (!calculated.length || this.runtime.mode === 'Display') {
      return values;
    }
    let current = values;
    for (let pass = 0; pass < MAX_CALCULATION_PASSES; pass++) {
      let changed = false;
      const env = this.env(current);
      const next: IFormValues = { ...current };
      calculated.forEach(({ control, field }) => {
        let result: unknown;
        try {
          result = coerceToField(field, evaluate(control.calculatedValue as string, env));
        } catch {
          result = null;
        }
        if (JSON.stringify(result) !== JSON.stringify(current[field.internalName])) {
          next[field.internalName] = result;
          changed = true;
        }
      });
      current = next;
      if (!changed) {
        break;
      }
    }
    return current;
  }

  /** Applies "set field" assignments (on-save rules and button actions). */
  public applyAssignments(assignments: IFieldAssignment[] | undefined, values: IFormValues): IFormValues {
    if (!assignments || !assignments.length) {
      return values;
    }
    const next: IFormValues = { ...values };
    assignments.forEach((a) => {
      const field = this.field(a.field);
      if (!field || field.readOnly || !a.expression) {
        return;
      }
      // Each assignment sees the results of the ones before it.
      next[field.internalName] = coerceToField(field, evaluate(a.expression, this.env(next)));
    });
    return next;
  }

  private validateControl(
    control: IControl,
    field: IFieldInfo,
    values: IFormValues,
    env: IExpressionEnv,
    state: IControlState
  ): string | undefined {
    const value = values[field.internalName];
    if (typeof value === 'number' && isNaN(value)) {
      return 'Enter a number.';
    }
    if (state.required && valueIsBlank(field, value)) {
      return 'You can\'t leave this blank.';
    }
    if (valueIsBlank(field, value)) {
      return undefined;
    }
    if ((field.kind === 'Text' || field.kind === 'Note') && field.maxLength && String(value).length > field.maxLength) {
      return `Use ${field.maxLength} characters or fewer.`;
    }
    if (field.kind === 'Number' || field.kind === 'Currency' || field.kind === 'Integer') {
      const n = value as number;
      const scale = field.showAsPercentage ? 100 : 1;
      if (field.min !== undefined && field.min !== null && n < field.min) {
        return `Enter ${field.min * scale}${field.showAsPercentage ? '%' : ''} or more.`;
      }
      if (field.max !== undefined && field.max !== null && n > field.max) {
        return `Enter ${field.max * scale}${field.showAsPercentage ? '%' : ''} or less.`;
      }
      if (field.kind === 'Integer' && Math.round(n) !== n) {
        return 'Enter a whole number.';
      }
    }
    if (field.kind === 'URL' && !/^(https?:\/\/|\/|mailto:)/i.test((value as IUrlValue).url)) {
      return 'Enter a link starting with https://';
    }
    for (const rule of control.validations || []) {
      if (!rule.expression || !rule.expression.trim()) {
        continue;
      }
      let ok: boolean;
      try {
        ok = toBool(evaluate(rule.expression, env));
      } catch {
        ok = false;
      }
      if (!ok) {
        return rule.message || 'This value is not valid.';
      }
    }
    return undefined;
  }

  /**
   * Validates visible, editable fields - on one page (wizard "Next") or the
   * whole form. Hidden fields are never validated, matching how a person
   * experiences the form.
   */
  public validate(values: IFormValues, pageId?: string): IValidationErrors {
    const errors: IValidationErrors = {};
    const state = this.computeState(values);
    const env = this.env(values);
    this.locations.forEach(({ control, page }) => {
      if (pageId && page.id !== pageId) {
        return;
      }
      const cs = state.controls[control.id];
      if (!cs || !cs.visible || cs.readOnly || control.type !== 'field') {
        return;
      }
      const field = this.field(control.field);
      if (!field) {
        return;
      }
      const error = this.validateControl(control, field, values, env, cs);
      if (error) {
        errors[control.id] = error;
      }
    });
    return errors;
  }

  /** Form-level rules from settings.validations. */
  public formErrors(values: IFormValues): string[] {
    const env = this.env(values);
    const errors: string[] = [];
    (this.form.settings.validations || []).forEach((rule) => {
      if (!rule.expression || !rule.expression.trim()) {
        return;
      }
      let ok: boolean;
      try {
        ok = toBool(evaluate(rule.expression, env));
      } catch {
        ok = false;
      }
      if (!ok) {
        errors.push(rule.message || 'The form is not valid.');
      }
    });
    return errors;
  }

  /**
   * Fields a save may write: those placed on the form plus those set by
   * assignments. Anything else on the list is never sent, and the payload
   * builder further drops fields whose value did not change.
   */
  public writableFields(extraAssignments?: IFieldAssignment[]): IFieldInfo[] {
    const names: string[] = [];
    const add = (name: string | undefined): void => {
      const field = this.field(name);
      if (field && !field.readOnly && names.indexOf(field.internalName) < 0) {
        names.push(field.internalName);
      }
    };
    this.fieldControls().forEach(({ field }) => add(field.internalName));
    (this.form.settings.onSave || []).forEach((a) => add(a.field));
    (extraAssignments || []).forEach((a) => add(a.field));
    return names.map((n) => this.fieldMap[n]);
  }

  public displayValue(field: IFieldInfo, values: IFormValues): string {
    return displayText(field, values[field.internalName]);
  }
}

/**
 * Picks the form for a list: the first one whose mode, audience and content
 * type all match, in the order the designer lists them.
 */
export function selectForm(
  forms: IFormDefinition[],
  mode: FormMode,
  userGroups: string[],
  contentTypeId?: string
): IFormDefinition | undefined {
  const groups = userGroups.map((g) => g.toLowerCase());
  const ct = (contentTypeId || '').toLowerCase();
  for (const form of forms) {
    if (form.modes.indexOf(mode) < 0) {
      continue;
    }
    if (form.audience.length && !form.audience.some((a) => groups.indexOf(a.toLowerCase()) >= 0)) {
      continue;
    }
    if (form.contentTypes.length && ct && !form.contentTypes.some((prefix) => ct.indexOf(prefix.toLowerCase()) === 0)) {
      continue;
    }
    return form;
  }
  return undefined;
}
