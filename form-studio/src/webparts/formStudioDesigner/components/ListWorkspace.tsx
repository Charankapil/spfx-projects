import * as React from 'react';
import {
  ChoiceGroup,
  CommandBar,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  Icon,
  ICommandBarItemProps,
  IconButton,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
  Pivot,
  PivotItem,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  SpinButton,
  TextField
} from '@fluentui/react';

import { FormView } from '../../../common/components/renderer/FormView';
import { IRuntimeContext } from '../../../common/form/FormEngine';
import { IFormValues } from '../../../common/form/formValues';
import { cloneForm, createControl, createDefaultForm } from '../../../common/model/formFactory';
import { IFieldInfo } from '../../../common/model/IFieldInfo';
import { ALL_MODES, ControlType, FormMode, FORMS_SCHEMA_VERSION, IFormDefinition, IListForms } from '../../../common/model/IFormDefinition';
import { FormDeployment, IContentTypeBinding, listFormUrl } from '../../../common/services/FormDeployment';
import { FormStore, touch } from '../../../common/services/FormStore';
import { errorMessage, ISpContext } from '../../../common/services/http';
import { ICurrentUser, IContentTypeInfo, IListInfo, IListSummary, ListService } from '../../../common/services/ListService';
import { Canvas } from './Canvas';
import styles from './Designer.module.scss';
import { addControl, findControl, findSection, IHistory, placedFields, pushHistory, redo, Selection, undo } from './designerOps';
import { Palette } from './Palette';
import { PropertiesPanel } from './PropertiesPanel';

export interface IListWorkspaceProps {
  ctx: ISpContext;
  listId: string;
  onChangeList: () => void;
}

interface IWorkspaceData {
  list: IListInfo;
  fields: IFieldInfo[];
  user: ICurrentUser;
  canManage: boolean;
  siteGroups: string[];
  contentTypes: IContentTypeInfo[];
  lists: IListSummary[];
}

function formFields(data: IWorkspaceData): IFieldInfo[] {
  return data.fields.filter((f) => f.kind !== 'Attachments' || data.list.enableAttachments);
}

function download(fileName: string, content: string): void {
  const blob = new Blob([content], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/** Accepts an exported list file or a single exported form; gives every imported form a fresh id. */
function parseImport(text: string): IFormDefinition[] {
  const parsed = JSON.parse(text);
  const forms: IFormDefinition[] = Array.isArray(parsed && parsed.forms) ? parsed.forms : parsed && Array.isArray(parsed.pages) ? [parsed] : [];
  if (!forms.length) {
    throw new Error('The file does not contain Form Studio forms.');
  }
  return forms.map((f) => {
    if (!f.name || !Array.isArray(f.pages) || !f.settings) {
      throw new Error('The file contains a form in an unknown format.');
    }
    return cloneForm(f, f.name);
  });
}

export const ListWorkspace: React.FC<IListWorkspaceProps> = (props) => {
  const { ctx, listId } = props;
  const listService = React.useMemo(() => new ListService(ctx), [ctx]);
  const store = React.useMemo(() => new FormStore(ctx), [ctx]);
  const deployment = React.useMemo(() => new FormDeployment(ctx), [ctx]);

  const [data, setData] = React.useState<IWorkspaceData | undefined>();
  const [loadError, setLoadError] = React.useState<string | undefined>();
  const [history, setHistory] = React.useState<IHistory<IListForms> | undefined>();
  const [savedJson, setSavedJson] = React.useState<string>('');
  const [formId, setFormId] = React.useState<string>('');
  const [selection, setSelection] = React.useState<Selection>({ kind: 'form' });
  const [activePageId, setActivePageId] = React.useState<string>('');
  const [view, setView] = React.useState<'design' | 'preview'>('design');
  const [saving, setSaving] = React.useState<boolean>(false);
  const [notice, setNotice] = React.useState<{ type: MessageBarType; text: string } | undefined>();
  const [deployOpen, setDeployOpen] = React.useState<boolean>(false);
  const [newFormOpen, setNewFormOpen] = React.useState<boolean>(false);
  const [newFormName, setNewFormName] = React.useState<string>('');
  const importInput = React.useRef<HTMLInputElement>(null);
  const propertiesPane = React.useRef<HTMLElement>(null);

  // Show the top of the properties for whatever was just selected.
  React.useEffect(() => {
    if (propertiesPane.current) {
      propertiesPane.current.scrollTop = 0;
    }
  }, [selection.kind === 'form' ? 'form' : selection.id]);

  React.useEffect(() => {
    let cancelled = false;
    const load = async (): Promise<void> => {
      const [list, fields, user, canManage, saved] = await Promise.all([
        listService.getList(listId),
        listService.getFields(listId),
        listService.getCurrentUser(),
        listService.canManageList(listId),
        store.load(listId)
      ]);
      const [siteGroups, contentTypes, lists] = await Promise.all([
        listService.getSiteGroups().catch(() => [] as string[]),
        listService.getContentTypes(listId).catch(() => [] as IContentTypeInfo[]),
        listService.getLists().catch(() => [] as IListSummary[])
      ]);
      if (cancelled) {
        return;
      }
      const loaded: IWorkspaceData = { list, fields, user, canManage, siteGroups, contentTypes, lists };
      const doc: IListForms = { ...saved, listTitle: list.title };
      const isNew = !doc.forms.length;
      if (isNew) {
        // Start from a working form so there is something to design straight away.
        doc.forms = [createDefaultForm(`${list.title} form`, formFields(loaded))];
      }
      setData(loaded);
      setHistory({ past: [], present: doc, future: [] });
      setSavedJson(isNew ? '' : JSON.stringify(saved.forms));
      setFormId(doc.forms[0].id);
      setActivePageId(doc.forms[0].pages[0].id);
    };
    load().catch((e) => !cancelled && setLoadError(errorMessage(e)));
    return () => {
      cancelled = true;
    };
  }, [listId]);

  const doc = history ? history.present : undefined;
  const form = doc ? doc.forms.filter((f) => f.id === formId)[0] || doc.forms[0] : undefined;
  const dirty = !!doc && JSON.stringify(doc.forms) !== savedJson;

  // Warn before leaving the page with unsaved changes.
  React.useEffect(() => {
    const handler = (e: BeforeUnloadEvent): void => {
      if (dirty) {
        e.preventDefault();
        e.returnValue = '';
      }
    };
    window.addEventListener('beforeunload', handler);
    return () => window.removeEventListener('beforeunload', handler);
  }, [dirty]);

  const commit = (next: IListForms): void => {
    if (history) {
      setHistory(pushHistory(history, next));
    }
  };

  const changeForm = (next: IFormDefinition, nextSelection?: Selection): void => {
    if (!doc) {
      return;
    }
    commit({ ...doc, forms: doc.forms.map((f) => (f.id === next.id ? next : f)) });
    if (nextSelection) {
      setSelection(nextSelection);
    }
  };

  const targetSectionId = (): string | undefined => {
    if (!form) {
      return undefined;
    }
    if (selection.kind === 'section' && findSection(form, selection.id)) {
      return selection.id;
    }
    if (selection.kind === 'control') {
      const loc = findControl(form, selection.id);
      if (loc) {
        return loc.section.id;
      }
    }
    const page = form.pages.filter((p) => p.id === activePageId)[0] || form.pages[0];
    return page.sections[page.sections.length - 1].id;
  };

  const addFromPalette = (type: ControlType, field?: IFieldInfo): void => {
    if (!form) {
      return;
    }
    if (field) {
      const existing = form.pages
        .map((p) => p.sections.map((s) => s.controls.filter((c) => c.field === field.internalName)[0]).filter((c) => !!c)[0])
        .filter((c) => !!c)[0];
      if (existing) {
        setSelection({ kind: 'control', id: existing.id });
        return;
      }
    }
    const sectionId = targetSectionId();
    if (!sectionId) {
      return;
    }
    const control = createControl(type, field);
    const loc = selection.kind === 'control' ? findControl(form, selection.id) : undefined;
    changeForm(addControl(form, sectionId, control, loc ? loc.index + 1 : undefined), { kind: 'control', id: control.id });
  };

  const save = async (): Promise<void> => {
    if (!doc || !data) {
      return;
    }
    setSaving(true);
    setNotice(undefined);
    try {
      const previous: IFormDefinition[] = savedJson ? JSON.parse(savedJson) : [];
      const stamped: IListForms = {
        ...doc,
        schemaVersion: FORMS_SCHEMA_VERSION,
        forms: doc.forms.map((f) => {
          const before = previous.filter((p) => p.id === f.id)[0];
          return !before || JSON.stringify(before) !== JSON.stringify(f) ? touch(f, data.user.name) : f;
        })
      };
      await store.save(stamped, data.user.name);
      setHistory({ past: history ? history.past : [], present: stamped, future: history ? history.future : [] });
      setSavedJson(JSON.stringify(stamped.forms));
      setNotice({ type: MessageBarType.success, text: 'Saved. People see the new form the next time they open it.' });
    } catch (e) {
      setNotice({ type: MessageBarType.error, text: `Could not save: ${errorMessage(e)}` });
    } finally {
      setSaving(false);
    }
  };

  // Ctrl+S / Ctrl+Z / Ctrl+Y.
  React.useEffect(() => {
    const handler = (e: KeyboardEvent): void => {
      if (!(e.ctrlKey || e.metaKey) || !history) {
        return;
      }
      const key = e.key.toLowerCase();
      const inField = /input|textarea/i.test((e.target as HTMLElement).tagName) || (e.target as HTMLElement).isContentEditable;
      if (key === 's') {
        e.preventDefault();
        save().catch(() => undefined);
      } else if (key === 'z' && !inField) {
        e.preventDefault();
        setHistory(e.shiftKey ? redo(history) : undo(history));
      } else if (key === 'y' && !inField) {
        e.preventDefault();
        setHistory(redo(history));
      }
    };
    document.addEventListener('keydown', handler);
    return () => document.removeEventListener('keydown', handler);
  });

  if (loadError) {
    return (
      <MessageBar messageBarType={MessageBarType.error} actions={<DefaultButton text="Choose another list" onClick={props.onChangeList} />}>
        {loadError}
      </MessageBar>
    );
  }
  if (!data || !doc || !form || !history) {
    return (
      <div className={styles.centered}>
        <Spinner size={SpinnerSize.large} label="Loading list..." />
      </div>
    );
  }

  const fields = formFields(data);
  const readOnly = !data.canManage;

  const selectForm = (id: string): void => {
    const f = doc.forms.filter((x) => x.id === id)[0];
    if (f) {
      setFormId(id);
      setSelection({ kind: 'form' });
      setActivePageId(f.pages[0].id);
    }
  };

  const moveForm = (id: string, delta: number): void => {
    const index = doc.forms.map((f) => f.id).indexOf(id);
    const target = index + delta;
    if (target < 0 || target >= doc.forms.length) {
      return;
    }
    const forms = doc.forms.slice();
    const [item] = forms.splice(index, 1);
    forms.splice(target, 0, item);
    commit({ ...doc, forms });
  };

  const deleteForm = (id: string): void => {
    if (doc.forms.length <= 1) {
      return;
    }
    // eslint-disable-next-line no-alert
    if (!window.confirm('Delete this form? You can undo until you save.')) {
      return;
    }
    const forms = doc.forms.filter((f) => f.id !== id);
    commit({ ...doc, forms });
    selectForm(forms[0].id);
  };

  const commands: ICommandBarItemProps[] = [
    { key: 'undo', text: 'Undo', ariaLabel: 'Undo (Ctrl+Z)', title: 'Undo (Ctrl+Z)', iconOnly: true, iconProps: { iconName: 'Undo' }, disabled: !history.past.length, onClick: () => setHistory(undo(history)) },
    { key: 'redo', text: 'Redo', ariaLabel: 'Redo (Ctrl+Y)', title: 'Redo (Ctrl+Y)', iconOnly: true, iconProps: { iconName: 'Redo' }, disabled: !history.future.length, onClick: () => setHistory(redo(history)) },
    {
      key: 'view',
      text: view === 'design' ? 'Preview' : 'Back to designer',
      iconProps: { iconName: view === 'design' ? 'Play' : 'Edit' },
      onClick: () => setView(view === 'design' ? 'preview' : 'design')
    },
    {
      key: 'save',
      text: saving ? 'Saving...' : dirty ? 'Save' : 'Saved',
      iconProps: { iconName: 'Save' },
      disabled: saving || readOnly || !dirty,
      onClick: () => {
        save().catch(() => undefined);
      }
    },
    { key: 'deploy', text: 'Use on list', iconProps: { iconName: 'PlugConnected' }, disabled: readOnly, onClick: () => setDeployOpen(true) }
  ];
  const farCommands: ICommandBarItemProps[] = [
    {
      key: 'more',
      text: 'More',
      ariaLabel: 'More actions',
      iconProps: { iconName: 'More' },
      iconOnly: true,
      subMenuProps: {
        items: [
          { key: 'export', text: 'Export forms (JSON)', iconProps: { iconName: 'Download' }, onClick: () => download(`${data.list.title} forms.json`, JSON.stringify(doc, undefined, 2)) },
          { key: 'import', text: 'Import forms...', iconProps: { iconName: 'Upload' }, disabled: readOnly, onClick: () => {
              if (importInput.current) {
                importInput.current.click();
              }
            } },
          {
            key: 'open',
            text: 'Open the list',
            iconProps: { iconName: 'OpenInNewWindow' },
            href: data.list.rootFolderUrl,
            target: '_blank'
          },
          { key: 'list', text: 'Design another list', iconProps: { iconName: 'BulletedList' }, onClick: props.onChangeList }
        ]
      }
    }
  ];

  return (
    <div className={styles.workspace}>
      <div className={styles.topBar}>
        <div className={styles.breadcrumb}>
          <Icon iconName="FormLibrary" className={styles.brandIcon} />
          <span className={styles.listName}>{data.list.title}</span>
          <Icon iconName="ChevronRight" className={styles.crumbSep} />
          <span className={styles.formName}>{form.name || 'Untitled form'}</span>
          {dirty && <span className={styles.dirty}>Unsaved changes</span>}
        </div>
        <CommandBar items={commands} farItems={farCommands} className={styles.commandBar} />
      </div>
      {readOnly && (
        <MessageBar messageBarType={MessageBarType.warning}>
          You need permission to manage this list to change its forms. You can look around and preview, but not save.
        </MessageBar>
      )}
      {notice && (
        <MessageBar messageBarType={notice.type} onDismiss={() => setNotice(undefined)}>
          {notice.text}
        </MessageBar>
      )}
      <input
        ref={importInput}
        type="file"
        accept=".json,application/json"
        style={{ display: 'none' }}
        onChange={(e) => {
          const file = e.target.files && e.target.files[0];
          e.target.value = '';
          if (!file) {
            return;
          }
          const reader = new FileReader();
          reader.onload = () => {
            try {
              const imported = parseImport(String(reader.result));
              commit({ ...doc, forms: doc.forms.concat(imported) });
              selectForm(imported[0].id);
              setNotice({ type: MessageBarType.success, text: `Imported ${imported.length} form${imported.length > 1 ? 's' : ''}. Fields that do not exist in this list show a warning on the canvas.` });
            } catch (err) {
              setNotice({ type: MessageBarType.error, text: errorMessage(err) });
            }
          };
          reader.readAsText(file);
        }}
      />

      {view === 'preview' ? (
        <PreviewPane form={form} fields={data.fields} user={data.user} listId={listId} listService={listService} />
      ) : (
        <div className={styles.designer}>
          <aside className={styles.leftPane}>
            <Pivot className={styles.leftPivot}>
              <PivotItem headerText="Add" itemIcon="Add">
                <Palette fields={fields} placed={placedFields(form)} onAddField={(f) => addFromPalette('field', f)} onAddElement={(t) => addFromPalette(t)} />
              </PivotItem>
              <PivotItem headerText={`Forms (${doc.forms.length})`} itemIcon="Documentation">
                <div className={styles.formsList}>
                  <div className={styles.help}>The first form that matches the mode, the person&apos;s groups and the content type is used.</div>
                  {doc.forms.map((f, i) => (
                    <div key={f.id} className={`${styles.formItem} ${f.id === form.id ? styles.formItemActive : ''}`}>
                      <button type="button" className={styles.formItemMain} onClick={() => selectForm(f.id)}>
                        <span className={styles.formItemName}>
                          {i + 1}. {f.name || 'Untitled form'}
                        </span>
                        <span className={styles.formItemMeta}>
                          {f.modes.join(', ') || 'No modes'}
                          {f.audience.length ? ` · ${f.audience.join(', ')}` : ' · Everyone'}
                        </span>
                      </button>
                      <IconButton iconProps={{ iconName: 'Up' }} title="Higher priority" ariaLabel="Higher priority" disabled={i === 0} onClick={() => moveForm(f.id, -1)} />
                      <IconButton iconProps={{ iconName: 'Down' }} title="Lower priority" ariaLabel="Lower priority" disabled={i === doc.forms.length - 1} onClick={() => moveForm(f.id, 1)} />
                      <IconButton
                        iconProps={{ iconName: 'Copy' }}
                        title="Duplicate"
                        ariaLabel="Duplicate"
                        onClick={() => {
                          const copy = cloneForm(f, `${f.name} (copy)`);
                          commit({ ...doc, forms: doc.forms.concat([copy]) });
                          selectForm(copy.id);
                        }}
                      />
                      <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" ariaLabel="Delete" disabled={doc.forms.length <= 1} onClick={() => deleteForm(f.id)} />
                    </div>
                  ))}
                  <DefaultButton
                    iconProps={{ iconName: 'Add' }}
                    text="New form"
                    onClick={() => {
                      setNewFormName(`${data.list.title} form ${doc.forms.length + 1}`);
                      setNewFormOpen(true);
                    }}
                    className={styles.newFormButton}
                  />
                </div>
              </PivotItem>
            </Pivot>
          </aside>
          <main className={styles.centerPane}>
            <Canvas
              form={form}
              fields={data.fields}
              selection={selection}
              activePageId={activePageId}
              onActivePage={setActivePageId}
              onSelect={setSelection}
              onChange={changeForm}
            />
          </main>
          <aside className={styles.rightPane} ref={propertiesPane}>
            <PropertiesPanel
              form={form}
              fields={data.fields}
              selection={selection}
              listId={listId}
              listService={listService}
              siteGroups={data.siteGroups}
              contentTypes={data.contentTypes}
              lists={data.lists}
              onChange={changeForm}
            />
          </aside>
        </div>
      )}

      <Dialog
        hidden={!newFormOpen}
        onDismiss={() => setNewFormOpen(false)}
        dialogContentProps={{ type: DialogType.normal, title: 'New form', subText: 'Starts with every editable field. Use a different form per audience (e.g. Approvers) or per mode (e.g. a simpler New form).' }}
      >
        <TextField label="Name" value={newFormName} onChange={(_e, v) => setNewFormName(v || '')} />
        <DialogFooter>
          <PrimaryButton
            text="Create"
            disabled={!newFormName.trim()}
            onClick={() => {
              const created = createDefaultForm(newFormName.trim(), fields);
              commit({ ...doc, forms: doc.forms.concat([created]) });
              selectForm(created.id);
              setNewFormOpen(false);
            }}
          />
          <DefaultButton text="Cancel" onClick={() => setNewFormOpen(false)} />
        </DialogFooter>
      </Dialog>

      {deployOpen && (
        <DeployPanel
          listId={listId}
          list={data.list}
          deployment={deployment}
          webUrl={ctx.pageContext.web.absoluteUrl}
          dirty={dirty}
          onDismiss={() => setDeployOpen(false)}
        />
      )}
    </div>
  );
};

function PreviewPane(props: { form: IFormDefinition; fields: IFieldInfo[]; user: ICurrentUser; listId: string; listService: ListService }): React.ReactElement {
  const [mode, setMode] = React.useState<FormMode>(props.form.modes[0] || 'New');
  const [itemId, setItemId] = React.useState<number>(0);
  const [values, setValues] = React.useState<IFormValues>({});
  const [error, setError] = React.useState<string | undefined>();
  const [run, setRun] = React.useState<number>(0);

  React.useEffect(() => {
    setError(undefined);
    if (mode === 'New' || !itemId) {
      setValues({});
      return;
    }
    props.listService
      .getItem(props.listId, itemId, props.fields)
      .then((item) => setValues(item.values))
      .catch((e) => {
        setValues({});
        setError(`Item ${itemId}: ${errorMessage(e)}`);
      });
  }, [mode, itemId]);

  const runtime: IRuntimeContext = React.useMemo(
    () => ({
      mode,
      itemId: mode === 'New' ? undefined : itemId || undefined,
      user: { name: props.user.name, email: props.user.email, loginName: props.user.loginName },
      groups: props.user.groups
    }),
    [mode, itemId]
  );

  return (
    <div className={styles.preview}>
      <div className={styles.previewBar}>
        <ChoiceGroup
          label="Preview as"
          selectedKey={mode}
          options={ALL_MODES.map((m) => ({ key: m, text: m }))}
          onChange={(_e, o) => o && setMode(o.key as FormMode)}
          styles={{ flexContainer: { display: 'flex', gap: 16 } }}
        />
        {mode !== 'New' && (
          <SpinButton label="Item ID" min={0} max={100000000} value={String(itemId)} onChange={(_e, v) => setItemId(Number(v) || 0)} styles={{ root: { width: 160 } }} />
        )}
        <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Reset" onClick={() => setRun(run + 1)} />
        <span className={styles.help}>Nothing is saved in preview. Rules use your own groups: {props.user.groups.join(', ') || 'none'}.</span>
      </div>
      {props.form.modes.indexOf(mode) < 0 && <MessageBar>This form is not set to be used for {mode.toLowerCase()} items; the preview shows it anyway.</MessageBar>}
      {error && <MessageBar messageBarType={MessageBarType.warning}>{error}</MessageBar>}
      <div className={styles.previewSurface}>
        <FormView
          key={`${run}-${mode}-${itemId}`}
          form={props.form}
          fields={props.fields}
          runtime={runtime}
          listId={props.listId}
          listService={props.listService}
          loadedValues={values}
          preview
          onCancel={() => setRun(run + 1)}
        />
      </div>
    </div>
  );
}

function DeployPanel(props: { listId: string; list: IListInfo; deployment: FormDeployment; webUrl: string; dirty: boolean; onDismiss: () => void }): React.ReactElement {
  const [bindings, setBindings] = React.useState<IContentTypeBinding[] | undefined>();
  const [modes, setModes] = React.useState<FormMode[]>(props.list.isLibrary ? ['Edit', 'Display'] : ALL_MODES.slice());
  const [busy, setBusy] = React.useState<boolean>(false);
  const [message, setMessage] = React.useState<{ type: MessageBarType; text: string } | undefined>();

  const load = (): void => {
    props.deployment
      .getBindings(props.listId)
      .then(setBindings)
      .catch((e) => setMessage({ type: MessageBarType.error, text: errorMessage(e) }));
  };
  React.useEffect(load, [props.listId]);

  const apply = async (enabled: boolean): Promise<void> => {
    setBusy(true);
    setMessage(undefined);
    try {
      await props.deployment.apply(props.listId, modes, enabled);
      setMessage({
        type: MessageBarType.success,
        text: enabled ? 'Done. The list now opens Form Studio forms for the selected modes.' : 'Done. The list uses SharePoint forms again for the selected modes.'
      });
      load();
    } catch (e) {
      setMessage({ type: MessageBarType.error, text: `${errorMessage(e)}. See the README for the PnP PowerShell alternative.` });
    } finally {
      setBusy(false);
    }
  };

  const mark = (on: boolean): React.ReactNode => (on ? <Icon iconName="CheckMark" className={styles.on} aria-label="Form Studio" /> : <span className={styles.off}>SharePoint</span>);

  return (
    <Panel isOpen type={PanelType.medium} headerText="Use Form Studio on this list" onDismiss={props.onDismiss} closeButtonAriaLabel="Close">
      <p>
        This points the list&apos;s New, Edit and Display forms at Form Studio, in place of the SharePoint form (or a Power Apps form). It applies to everyone
        opening items in this list, in the browser and in Microsoft Lists.
      </p>
      {props.dirty && <MessageBar messageBarType={MessageBarType.warning}>You have unsaved changes. Save first, or people will see the previous version.</MessageBar>}
      {props.list.isLibrary && <MessageBar>For document libraries, only the Edit and Display (properties) forms can be replaced.</MessageBar>}
      <h3>Current forms</h3>
      {!bindings ? (
        <Spinner size={SpinnerSize.small} />
      ) : (
        <table className={styles.bindingTable}>
          <thead>
            <tr>
              <th scope="col">Content type</th>
              <th scope="col">New</th>
              <th scope="col">Edit</th>
              <th scope="col">Display</th>
            </tr>
          </thead>
          <tbody>
            {bindings.map((b) => (
              <tr key={b.id}>
                <td>{b.name}</td>
                <td>{mark(b.modes.New)}</td>
                <td>{mark(b.modes.Edit)}</td>
                <td>{mark(b.modes.Display)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
      <h3>Change</h3>
      <div className={styles.checkRow}>
        {ALL_MODES.map((m) => (
          <label key={m} className={styles.inlineCheck}>
            <input
              type="checkbox"
              checked={modes.indexOf(m) >= 0}
              disabled={props.list.isLibrary && m === 'New'}
              onChange={(e) => setModes(e.target.checked ? ALL_MODES.filter((x) => x === m || modes.indexOf(x) >= 0) : modes.filter((x) => x !== m))}
            />{' '}
            {m} form
          </label>
        ))}
      </div>
      {message && <MessageBar messageBarType={message.type}>{message.text}</MessageBar>}
      <div className={styles.buttonRow}>
        <PrimaryButton text="Use Form Studio" disabled={busy || !modes.length} onClick={() => apply(true).catch(() => undefined)} />
        <DefaultButton text="Restore SharePoint forms" disabled={busy || !modes.length} onClick={() => apply(false).catch(() => undefined)} />
        {busy && <Spinner size={SpinnerSize.small} />}
      </div>
      <h3>Try it</h3>
      <div className={styles.buttonRow}>
        {!props.list.isLibrary && <DefaultButton text="Open the New form" href={listFormUrl(props.webUrl, props.listId, 'New')} target="_blank" />}
        <DefaultButton text="Open the list" href={props.list.rootFolderUrl} target="_blank" />
      </div>
      <p className={styles.help}>
        Forms are picked per item: the first form in the Forms tab whose mode, audience and content type match. If none matches, an automatic form with every
        field is shown.
      </p>
    </Panel>
  );
}

