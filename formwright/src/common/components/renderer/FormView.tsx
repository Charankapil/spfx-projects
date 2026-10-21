import * as React from 'react';
import {
  DefaultButton,
  Icon,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Spinner,
  SpinnerSize
} from '@fluentui/react';

import { FormEngine, IFormState, IRuntimeContext, IValidationErrors } from '../../form/FormEngine';
import { IFormValues } from '../../form/formValues';
import { IFieldInfo } from '../../model/IFieldInfo';
import { IControl, IFormDefinition, IFormPage, ISection } from '../../model/IFormDefinition';
import { errorMessage } from '../../services/http';
import { ISaveProgress, newSaveProgress, saveItem } from '../../services/FormSaver';
import { ListService } from '../../services/ListService';
import { renderRichText } from '../richContent';
import { FieldEditor } from './FieldControl';
import styles from './FormRenderer.module.scss';
import { IRelatedItemsHandle, RelatedItems } from './RelatedItems';

export interface IFormViewProps {
  form: IFormDefinition;
  fields: IFieldInfo[];
  runtime: IRuntimeContext;
  listId: string;
  listService?: ListService;
  /** Values of the loaded item; empty for a new item. */
  loadedValues: IFormValues;
  etag?: string;
  contentTypeId?: string;
  folderPath?: string;
  /** Designer preview: everything works except writing to SharePoint. */
  preview?: boolean;
  onSaved?: (itemId: number, values: IFormValues, redirectUrl?: string, successMessage?: string) => void;
  onCancel?: () => void;
  /** Display mode: switch to the edit form. */
  onEdit?: () => void;
}

const SPAN_MAX = 12;

function clampSpan(span: number, columns: number): number {
  return Math.max(1, Math.min(span || 1, columns, SPAN_MAX));
}

/**
 * Renders a form definition against a list item. All behaviour comes from
 * FormEngine; this component only draws it and talks to SharePoint on save.
 */
export const FormView: React.FC<IFormViewProps> = (props) => {
  const { form, fields, runtime, listService, preview } = props;
  const engine = React.useMemo(() => new FormEngine(form, fields, runtime), [form, fields, runtime]);
  const [values, setValues] = React.useState<IFormValues>(() => engine.initialValues(props.loadedValues));
  const [errors, setErrors] = React.useState<IValidationErrors>({});
  const [formErrors, setFormErrors] = React.useState<string[]>([]);
  const [saving, setSaving] = React.useState<boolean>(false);
  const [saveError, setSaveError] = React.useState<string | undefined>();
  const [previewNote, setPreviewNote] = React.useState<string | undefined>();
  const [dirty, setDirty] = React.useState<boolean>(false);
  const [collapsed, setCollapsed] = React.useState<{ [sectionId: string]: boolean }>(() => {
    const initial: { [id: string]: boolean } = {};
    form.pages.forEach((p) => p.sections.forEach((s) => (initial[s.id] = !!(s.collapsible && s.collapsedByDefault))));
    return initial;
  });
  const related = React.useRef<{ [controlId: string]: IRelatedItemsHandle }>({});
  // Survives a failed save so the retry resumes instead of repeating work.
  const saveProgress = React.useRef<ISaveProgress>(newSaveProgress());

  // Re-initialise when the designer swaps the definition under a live preview.
  React.useEffect(() => {
    setValues(engine.initialValues(props.loadedValues));
    setErrors({});
    setFormErrors([]);
  }, [engine, props.loadedValues]);

  const state: IFormState = React.useMemo(() => engine.computeState(values), [engine, values]);
  const visiblePages = form.pages.filter((p) => state.pages[p.id]);
  const [pageId, setPageId] = React.useState<string>(form.pages[0] ? form.pages[0].id : '');
  const currentPage = visiblePages.filter((p) => p.id === pageId)[0] || visiblePages[0];
  const pageIndex = currentPage ? visiblePages.indexOf(currentPage) : 0;
  const isDisplay = runtime.mode === 'Display';

  // Resolve titles of lookups and people set by defaults (they arrive as bare ids / e-mails).
  React.useEffect(() => {
    if (!listService) {
      return;
    }
    let cancelled = false;
    const work = async (): Promise<void> => {
      for (const { field } of engine.fieldControls()) {
        const v = values[field.internalName] as { id?: number; title?: string; unresolved?: boolean } | null;
        if (!v || Array.isArray(v)) {
          continue;
        }
        try {
          if (field.kind === 'Lookup' && v.id && !v.title) {
            const title = await listService.getLookupTitle(field, v.id);
            if (!cancelled) {
              setValues((prev) => ({ ...prev, [field.internalName]: { id: v.id, title } }));
            }
          } else if (field.kind === 'User' && v.unresolved && v.title) {
            const user = await listService.ensureUser(v.title);
            if (!cancelled) {
              setValues((prev) => ({ ...prev, [field.internalName]: user }));
            }
          }
        } catch {
          // Leave the value unresolved; saving will try again and report a real error.
        }
      }
    };
    work().catch(() => undefined);
    return () => {
      cancelled = true;
    };
  }, [engine]);

  const setFieldValue = (name: string, value: unknown): void => {
    setValues((prev) => engine.applyCalculated({ ...prev, [name]: value }));
    setDirty(true);
    setErrors((prev) => {
      if (!Object.keys(prev).length) {
        return prev;
      }
      const next = { ...prev };
      engine.fieldControls().forEach(({ control, field }) => {
        if (field.internalName === name) {
          delete next[control.id];
        }
      });
      return next;
    });
  };

  const goToFirstError = (errs: IValidationErrors): void => {
    const first = engine.controls.filter((l) => errs[l.control.id])[0];
    if (!first) {
      return;
    }
    setPageId(first.page.id);
    if (collapsed[first.section.id]) {
      setCollapsed({ ...collapsed, [first.section.id]: false });
    }
    window.setTimeout(() => {
      const el = document.getElementById(`fs-${first.control.id}`);
      if (el) {
        el.focus();
        el.scrollIntoView({ block: 'center', behavior: 'smooth' });
      }
    }, 50);
  };

  const submit = async (button?: IControl): Promise<void> => {
    setSaveError(undefined);
    setPreviewNote(undefined);
    let next: IFormValues;
    try {
      next = engine.applyAssignments(button ? button.assignments : undefined, values);
      next = engine.applyAssignments(form.settings.onSave, next);
      next = engine.applyCalculated(next);
    } catch (e) {
      setSaveError(`A "set field" rule failed: ${errorMessage(e)}`);
      return;
    }

    if (!(button && button.skipValidation) && !isDisplay) {
      const errs = engine.validate(next);
      const formLevel = engine.formErrors(next);
      const relatedErrors = Object.keys(related.current)
        .map((k) => related.current[k].validate())
        .filter((m): m is string => !!m);
      setErrors(errs);
      setFormErrors(formLevel.concat(relatedErrors));
      if (Object.keys(errs).length || formLevel.length || relatedErrors.length) {
        goToFirstError(errs);
        return;
      }
    }

    if (preview || !listService) {
      setValues(next);
      setPreviewNote('Preview: the form is valid. Nothing was saved.');
      return;
    }

    setSaving(true);
    try {
      const itemId = await saveItem({
        listService,
        listId: props.listId,
        mode: runtime.mode,
        itemId: runtime.itemId,
        fields: engine.writableFields(button ? button.assignments : undefined),
        values: next,
        original: props.loadedValues,
        etag: props.etag,
        detectConflicts: form.settings.detectConflicts,
        contentTypeId: props.contentTypeId,
        folderPath: props.folderPath,
        progress: saveProgress.current
      });
      for (const key of Object.keys(related.current)) {
        await related.current[key].commit(itemId);
      }
      setDirty(false);
      saveProgress.current = newSaveProgress();
      const savedValues = { ...next, ID: itemId, Id: itemId };
      const savedEnv = new FormEngine(form, fields, { ...runtime, itemId });
      const redirectTemplate = (button && button.url) || form.settings.redirectUrl;
      const redirect = redirectTemplate ? savedEnv.template(redirectTemplate, savedValues, encodeURIComponent) : undefined;
      if (props.onSaved) {
        props.onSaved(itemId, savedValues, redirect, savedEnv.template(form.settings.successMessage, savedValues));
      }
    } catch (e) {
      setSaveError(errorMessage(e));
    } finally {
      setSaving(false);
    }
  };

  const cancel = (): void => {
    // eslint-disable-next-line no-alert
    if (dirty && !preview && !window.confirm('Discard your changes?')) {
      return;
    }
    if (props.onCancel) {
      props.onCancel();
    }
  };

  const next = (): void => {
    if (!currentPage) {
      return;
    }
    if (!isDisplay) {
      const errs = engine.validate(values, currentPage.id);
      setErrors(errs);
      if (Object.keys(errs).length) {
        goToFirstError(errs);
        return;
      }
    }
    const target = visiblePages[pageIndex + 1];
    if (target) {
      setPageId(target.id);
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
  };

  const renderControl = (control: IControl, section: ISection): React.ReactNode => {
    const cs = state.controls[control.id];
    if (!cs || !cs.visible) {
      return null;
    }
    const span = clampSpan(control.span, section.columns);
    const style = { gridColumn: `span ${span}` };
    const id = `fs-${control.id}`;
    switch (control.type) {
      case 'field': {
        const field = engine.field(control.field);
        if (!field) {
          return null;
        }
        const label = engine.label(control);
        const error = errors[control.id];
        return (
          <div key={control.id} className={`${styles.field} ${error ? styles.hasError : ''}`} style={style}>
            <label id={`${id}-label`} htmlFor={id} className={styles.label}>
              {label}
              {cs.required && <span className={styles.required} aria-hidden="true"> *</span>}
            </label>
            <div className={styles.control}>
              <FieldEditor
                id={id}
                control={control}
                field={field}
                value={values[field.internalName]}
                values={values}
                state={cs}
                label={label}
                engine={engine}
                listService={listService}
                onChange={(v) => setFieldValue(field.internalName, v)}
              />
              {(control.description || field.description) && !cs.readOnly && (
                <div className={styles.description}>{control.description || field.description}</div>
              )}
              {error && (
                <div className={styles.error} role="alert">
                  <Icon iconName="Error" /> {error}
                </div>
              )}
            </div>
          </div>
        );
      }
      case 'heading': {
        const text = engine.template(control.content, values);
        const level = Math.max(1, Math.min(control.level || 2, 4));
        const Tag = `h${level + 1}` as 'h2' | 'h3' | 'h4' | 'h5';
        return (
          <Tag key={control.id} className={`${styles.heading} ${styles[`h${level}`]}`} style={style}>
            {text}
          </Tag>
        );
      }
      case 'text':
        return (
          <div key={control.id} className={styles.text} style={style}>
            {renderRichText(engine.template(control.content, values))}
          </div>
        );
      case 'divider':
        return <hr key={control.id} className={styles.divider} style={style} />;
      case 'spacer':
        return <div key={control.id} style={{ ...style, height: control.height || 24 }} aria-hidden="true" />;
      case 'image': {
        const src = engine.template(control.imageUrl, values);
        if (!src) {
          return null;
        }
        return (
          <div key={control.id} className={styles.image} style={style}>
            <img src={src} alt={control.altText || ''} style={control.imageWidth ? { width: control.imageWidth, maxWidth: '100%' } : undefined} />
          </div>
        );
      }
      case 'button': {
        const action = control.action || 'save';
        const onClick = (): void => {
          if (action === 'save') {
            submit(control).catch(() => undefined);
          } else if (action === 'close') {
            cancel();
          } else if (action === 'print') {
            window.print();
          } else if (action === 'link') {
            const url = engine.template(control.url, values, encodeURIComponent);
            if (url && /^(https?:\/\/|\/)/i.test(url)) {
              window.open(url, '_blank', 'noopener');
            }
          }
        };
        const Button = control.buttonStyle === 'primary' ? PrimaryButton : DefaultButton;
        return (
          <div key={control.id} className={styles.buttonCell} style={style}>
            <Button text={control.label || 'Button'} onClick={onClick} disabled={saving} />
          </div>
        );
      }
      case 'relatedItems':
        if (!control.relatedItems) {
          return null;
        }
        return (
          <div key={control.id} className={styles.field} style={style}>
            <div className={styles.label}>{control.label}</div>
            <RelatedItems
              config={control.relatedItems}
              label={control.label || 'Related items'}
              parentId={runtime.itemId}
              readOnly={cs.readOnly || preview === true}
              listService={listService}
              register={(handle) => {
                if (handle) {
                  related.current[control.id] = handle;
                } else {
                  delete related.current[control.id];
                }
              }}
              onDirty={() => setDirty(true)}
            />
          </div>
        );
      default:
        return null;
    }
  };

  const renderSection = (section: ISection): React.ReactNode => {
    if (!state.sections[section.id]) {
      return null;
    }
    const isCollapsed = !!collapsed[section.id];
    const columns = Math.max(1, Math.min(section.columns || 1, 4));
    const titleId = `fs-${section.id}-title`;
    return (
      <section key={section.id} className={styles.section} aria-labelledby={section.title ? titleId : undefined}>
        {section.title &&
          (section.collapsible ? (
            <button
              type="button"
              className={styles.sectionToggle}
              aria-expanded={!isCollapsed}
              onClick={() => setCollapsed({ ...collapsed, [section.id]: !isCollapsed })}
            >
              <Icon iconName={isCollapsed ? 'ChevronRight' : 'ChevronDown'} />
              <span id={titleId} className={styles.sectionTitle}>{section.title}</span>
            </button>
          ) : (
            <div id={titleId} className={styles.sectionTitle}>{section.title}</div>
          ))}
        {section.description && !isCollapsed && <div className={styles.sectionDescription}>{section.description}</div>}
        {!isCollapsed && (
          <div className={styles.grid} style={{ gridTemplateColumns: `repeat(${columns}, minmax(0, 1fr))` }}>
            {section.controls.map((c) => renderControl(c, section))}
          </div>
        )}
      </section>
    );
  };

  const pageHasErrors = (page: IFormPage): boolean => page.sections.some((s) => s.controls.some((c) => !!errors[c.id]));

  const title = engine.template(form.settings.title, values);
  const subtitle = engine.template(form.settings.subtitle, values);
  const isWizard = form.layout === 'wizard' && visiblePages.length > 1;
  const isTabs = form.layout === 'tabs' && visiblePages.length > 1;
  // Every page stays mounted (inactive tabs / steps are only hidden) so
  // state inside controls, such as unsaved related-item rows, survives
  // switching pages.
  const paged = isWizard || isTabs;
  const lastStep = !isWizard || pageIndex === visiblePages.length - 1;
  const accentColor = form.settings.accentColor;
  const accent = accentColor && /^(#[0-9a-f]{3,8}|[a-z]+)$/i.test(accentColor.trim()) ? ({ '--fs-accent': accentColor.trim() } as React.CSSProperties) : {};

  return (
    <div
      className={`${styles.form} ${form.settings.labelPosition === 'left' ? styles.labelsLeft : ''}`}
      style={{ maxWidth: form.settings.maxWidth || undefined, ...accent }}
    >
      {(title || subtitle) && (
        <header className={styles.header}>
          {title && <h1 className={styles.title}>{title}</h1>}
          {subtitle && <div className={styles.subtitle}>{subtitle}</div>}
        </header>
      )}

      {isTabs && (
        <div className={styles.tabs} role="tablist">
          {visiblePages.map((p) => (
            <button
              key={p.id}
              type="button"
              role="tab"
              aria-selected={p.id === (currentPage && currentPage.id)}
              className={`${styles.tab} ${p.id === (currentPage && currentPage.id) ? styles.tabActive : ''}`}
              onClick={() => setPageId(p.id)}
            >
              {p.title}
              {pageHasErrors(p) && <Icon iconName="AlertSolid" className={styles.tabError} aria-label="has errors" />}
            </button>
          ))}
        </div>
      )}

      {isWizard && (
        <ol className={styles.steps}>
          {visiblePages.map((p, i) => (
            <li key={p.id} className={`${styles.step} ${i === pageIndex ? styles.stepActive : ''} ${i < pageIndex ? styles.stepDone : ''}`} aria-current={i === pageIndex ? 'step' : undefined}>
              <span className={styles.stepNumber}>{i < pageIndex ? <Icon iconName="CheckMark" /> : i + 1}</span>
              <span className={styles.stepTitle}>{p.title}</span>
            </li>
          ))}
        </ol>
      )}

      <div role={isTabs ? 'tabpanel' : undefined}>
        {visiblePages.map((page) => (
          <div key={page.id} className={styles.page} hidden={paged && page !== currentPage}>
            {!isWizard && !isTabs && form.pages.length > 1 && page.title && <h2 className={styles.pageTitle}>{page.title}</h2>}
            {page.sections.map(renderSection)}
          </div>
        ))}
      </div>

      {formErrors.length > 0 && (
        <MessageBar messageBarType={MessageBarType.error} className={styles.messages}>
          {formErrors.map((m, i) => (
            <div key={i}>{m}</div>
          ))}
        </MessageBar>
      )}
      {Object.keys(errors).length > 0 && !formErrors.length && (
        <MessageBar messageBarType={MessageBarType.error} className={styles.messages}>
          Some fields need your attention.
        </MessageBar>
      )}
      {saveError && (
        <MessageBar messageBarType={MessageBarType.error} className={styles.messages} onDismiss={() => setSaveError(undefined)}>
          {saveError}
        </MessageBar>
      )}
      {previewNote && (
        <MessageBar messageBarType={MessageBarType.success} className={styles.messages} onDismiss={() => setPreviewNote(undefined)}>
          {previewNote}
        </MessageBar>
      )}

      {form.settings.showDefaultButtons && (
        <div className={styles.footer}>
          {saving && <Spinner size={SpinnerSize.small} label="Saving..." labelPosition="right" />}
          {isWizard && pageIndex > 0 && <DefaultButton text="Back" iconProps={{ iconName: 'ChevronLeft' }} onClick={() => setPageId(visiblePages[pageIndex - 1].id)} disabled={saving} />}
          {isWizard && !lastStep && <PrimaryButton text="Next" onClick={next} disabled={saving} />}
          {!isDisplay && lastStep && <PrimaryButton text={form.settings.submitLabel || 'Save'} onClick={() => submit().catch(() => undefined)} disabled={saving} />}
          {isDisplay && props.onEdit && <PrimaryButton text="Edit" iconProps={{ iconName: 'Edit' }} onClick={props.onEdit} />}
          {props.onCancel && <DefaultButton text={isDisplay ? 'Close' : form.settings.cancelLabel || 'Cancel'} onClick={cancel} disabled={saving} />}
        </div>
      )}
    </div>
  );
};
