import * as React from 'react';
import {
  ActionButton,
  Checkbox,
  ChoiceGroup,
  DefaultButton,
  Dropdown,
  IconButton,
  IDropdownOption,
  Label,
  MessageBar,
  Slider,
  SpinButton,
  TextField,
  Toggle
} from '@fluentui/react';

import { IFieldInfo } from '../common/model/IFieldInfo';
import {
  ALL_MODES,
  FieldVariant,
  FormMode,
  IControl,
  IFieldAssignment,
  IFormDefinition,
  IFormSettings,
  IFormPage,
  ISection,
  IValidationRule
} from '../common/model/IFormDefinition';
import { IContentTypeInfo, IListSummary, ListService } from '../common/services/ListService';
import styles from './Designer.module.scss';
import { findControl, findSection, removePage, nudgePage, Selection, updateControl, updatePage, updateSection } from './designerOps';
import { ExpressionField } from './ExpressionField';

export interface IPropertiesPanelProps {
  form: IFormDefinition;
  fields: IFieldInfo[];
  selection: Selection;
  listId: string;
  listService: ListService;
  siteGroups: string[];
  contentTypes: IContentTypeInfo[];
  lists: IListSummary[];
  onChange: (form: IFormDefinition, selection?: Selection) => void;
}

const Group: React.FC<{ title: string; defaultOpen?: boolean }> = (props) => {
  const [open, setOpen] = React.useState<boolean>(props.defaultOpen !== false);
  return (
    <div className={styles.propGroup}>
      <button type="button" className={styles.propGroupHeader} aria-expanded={open} onClick={() => setOpen(!open)}>
        <span>{props.title}</span>
        <span aria-hidden="true">{open ? '−' : '+'}</span>
      </button>
      {open && <div className={styles.propGroupBody}>{props.children}</div>}
    </div>
  );
};

function RuleList(props: { label: string; rules: IValidationRule[]; fields: IFieldInfo[]; help: string; onChange: (rules: IValidationRule[]) => void }): React.ReactElement {
  const { rules } = props;
  return (
    <div>
      <Label>{props.label}</Label>
      <div className={styles.help}>{props.help}</div>
      {rules.map((rule, i) => (
        <div key={i} className={styles.ruleCard}>
          <ExpressionField
            label="Valid when"
            value={rule.expression}
            fields={props.fields}
            placeholder='e.g. [EndDate] >= [StartDate]'
            onChange={(v) => props.onChange(rules.map((r, j) => (j === i ? { ...r, expression: v } : r)))}
          />
          <TextField label="Message" value={rule.message} onChange={(_e, v) => props.onChange(rules.map((r, j) => (j === i ? { ...r, message: v || '' } : r)))} />
          <ActionButton iconProps={{ iconName: 'Delete' }} onClick={() => props.onChange(rules.filter((_r, j) => j !== i))}>
            Remove rule
          </ActionButton>
        </div>
      ))}
      <ActionButton iconProps={{ iconName: 'Add' }} onClick={() => props.onChange(rules.concat([{ expression: '', message: '' }]))}>
        Add rule
      </ActionButton>
    </div>
  );
}

function AssignmentList(props: { label: string; help: string; items: IFieldAssignment[]; fields: IFieldInfo[]; onChange: (items: IFieldAssignment[]) => void }): React.ReactElement {
  const writable = props.fields.filter((f) => !f.readOnly);
  const options: IDropdownOption[] = writable.map((f) => ({ key: f.internalName, text: f.title }));
  return (
    <div>
      <Label>{props.label}</Label>
      <div className={styles.help}>{props.help}</div>
      {props.items.map((a, i) => (
        <div key={i} className={styles.ruleCard}>
          <Dropdown label="Field" options={options} selectedKey={a.field || null} onChange={(_e, o) => o && props.onChange(props.items.map((x, j) => (j === i ? { ...x, field: String(o.key) } : x)))} />
          <ExpressionField
            label="Value"
            value={a.expression}
            fields={props.fields}
            placeholder='e.g. "Approved" or Now() or UserEmail()'
            onChange={(v) => props.onChange(props.items.map((x, j) => (j === i ? { ...x, expression: v } : x)))}
          />
          <ActionButton iconProps={{ iconName: 'Delete' }} onClick={() => props.onChange(props.items.filter((_x, j) => j !== i))}>
            Remove
          </ActionButton>
        </div>
      ))}
      <ActionButton iconProps={{ iconName: 'Add' }} onClick={() => props.onChange(props.items.concat([{ field: '', expression: '' }]))}>
        Add field to set
      </ActionButton>
    </div>
  );
}

function variantOptions(field: IFieldInfo): { key: FieldVariant; text: string }[] | undefined {
  switch (field.kind) {
    case 'Choice':
      return [
        { key: 'dropdown', text: 'Drop-down' },
        { key: 'radio', text: 'Radio buttons' }
      ];
    case 'MultiChoice':
      return [
        { key: 'checkboxes', text: 'Checkboxes' },
        { key: 'dropdown', text: 'Drop-down' }
      ];
    case 'Boolean':
      return [
        { key: 'toggle', text: 'Toggle' },
        { key: 'checkbox', text: 'Checkbox' }
      ];
    default:
      return undefined;
  }
}

function WidthSlider(props: { control: IControl; columns: number; onChange: (span: number) => void }): React.ReactElement | null {
  if (props.columns <= 1) {
    return null;
  }
  const value = Math.min(props.control.span || 1, props.columns);
  return (
    <Slider
      label="Width (columns)"
      min={1}
      max={props.columns}
      step={1}
      value={value}
      showValue
      valueFormat={(v) => (v === props.columns ? 'Full width' : `${v} of ${props.columns}`)}
      onChange={(v) => props.onChange(v === props.columns ? 12 : v)}
    />
  );
}

function LookupSettings(props: { control: IControl; field: IFieldInfo; fields: IFieldInfo[]; listService: ListService; update: (patch: Partial<IControl>) => void }): React.ReactElement {
  const { control, field } = props;
  const [sourceFields, setSourceFields] = React.useState<IFieldInfo[]>([]);
  React.useEffect(() => {
    if (field.lookupListId) {
      props.listService.getFields(field.lookupListId).then(setSourceFields).catch(() => setSourceFields([]));
    }
  }, [field.lookupListId]);
  const cascade = control.lookupCascade || { sourceField: '', formField: '' };
  return (
    <Group title="Filter options" defaultOpen={!!(control.lookupCascade || control.lookupFilter)}>
      <div className={styles.help}>Cascading: only offer items whose column matches another field on this form (e.g. Cities in the chosen Country).</div>
      <Dropdown
        label="Column in the lookup list"
        options={[{ key: '', text: '(no cascading filter)' }].concat(sourceFields.filter((f) => f.kind !== 'Attachments').map((f) => ({ key: f.internalName, text: `${f.title} (${f.kind})` })))}
        selectedKey={cascade.sourceField}
        onChange={(_e, o) => {
          const sourceField = o ? String(o.key) : '';
          props.update({ lookupCascade: sourceField ? { ...cascade, sourceField } : undefined });
        }}
      />
      {cascade.sourceField && (
        <Dropdown
          label="must equal this form field"
          options={props.fields.filter((f) => f.internalName !== field.internalName).map((f) => ({ key: f.internalName, text: f.title }))}
          selectedKey={cascade.formField || null}
          onChange={(_e, o) => o && props.update({ lookupCascade: { ...cascade, formField: String(o.key) } })}
        />
      )}
      <ExpressionField
        label="Custom OData filter"
        template
        value={control.lookupFilter}
        fields={props.fields}
        placeholder="e.g. Active eq 1 and Region eq '{{[Region]}}'"
        description="Added to the lookup query with 'and'. {{ }} inserts form values, with quotes escaped."
        onChange={(v) => props.update({ lookupFilter: v || undefined })}
      />
    </Group>
  );
}

function RelatedItemsSettings(props: { control: IControl; listId: string; lists: IListSummary[]; listService: ListService; update: (patch: Partial<IControl>) => void }): React.ReactElement {
  const config = props.control.relatedItems || { listId: '', parentLookupField: '', columns: [], allowAdd: true, allowEdit: true, allowDelete: true };
  const [childFields, setChildFields] = React.useState<IFieldInfo[]>([]);
  React.useEffect(() => {
    if (config.listId) {
      props.listService.getFields(config.listId).then(setChildFields).catch(() => setChildFields([]));
    } else {
      setChildFields([]);
    }
  }, [config.listId]);
  const set = (patch: Partial<typeof config>): void => props.update({ relatedItems: { ...config, ...patch } });
  const parentId = props.listId.replace(/[{}]/g, '').toLowerCase();
  const lookups = childFields.filter((f) => (f.kind === 'Lookup' || f.kind === 'LookupMulti') && (f.lookupListId || '').toLowerCase() === parentId);
  const columnCandidates = childFields.filter((f) => f.internalName !== config.parentLookupField && f.kind !== 'Attachments');
  return (
    <Group title="Related list">
      <Dropdown
        label="Child list"
        options={props.lists.filter((l) => l.id.toLowerCase() !== parentId && !l.isLibrary).map((l) => ({ key: l.id, text: l.title }))}
        selectedKey={config.listId || null}
        onChange={(_e, o) => {
          if (o) {
            const list = props.lists.filter((l) => l.id === o.key)[0];
            set({ listId: String(o.key), listTitle: list ? list.title : '', parentLookupField: '', columns: [] });
          }
        }}
      />
      {config.listId && !lookups.length && childFields.length > 0 && (
        <MessageBar>This list has no lookup column pointing at the list you are designing. Add one (e.g. &quot;Parent&quot;) in list settings first.</MessageBar>
      )}
      {lookups.length > 0 && (
        <Dropdown
          label="Lookup column that links rows to this item"
          options={lookups.map((f) => ({ key: f.internalName, text: f.title }))}
          selectedKey={config.parentLookupField || null}
          onChange={(_e, o) => o && set({ parentLookupField: String(o.key) })}
        />
      )}
      {config.parentLookupField && (
        <>
          <Label>Columns</Label>
          <div className={styles.checkList}>
            {columnCandidates.map((f) => (
              <Checkbox
                key={f.internalName}
                label={f.title}
                checked={config.columns.indexOf(f.internalName) >= 0}
                onChange={(_e, checked) =>
                  set({ columns: checked ? config.columns.concat([f.internalName]) : config.columns.filter((c) => c !== f.internalName) })
                }
              />
            ))}
          </div>
          <div className={styles.help}>Text, number, choice, yes/no and date columns are editable inline; others are shown read-only.</div>
        </>
      )}
      <Toggle label="Allow adding rows" checked={config.allowAdd} onChange={(_e, c) => set({ allowAdd: !!c })} />
      <Toggle label="Allow editing rows" checked={config.allowEdit} onChange={(_e, c) => set({ allowEdit: !!c })} />
      <Toggle label="Allow deleting rows" checked={config.allowDelete} onChange={(_e, c) => set({ allowDelete: !!c })} />
    </Group>
  );
}

function ControlProperties(props: IPropertiesPanelProps & { control: IControl; section: ISection }): React.ReactElement {
  const { control, fields, form, section } = props;
  const update = (patch: Partial<IControl>): void => props.onChange(updateControl(form, control.id, patch));
  const field = control.type === 'field' ? fields.filter((f) => f.internalName === control.field)[0] : undefined;
  const width = <WidthSlider control={control} columns={section.columns} onChange={(span) => update({ span })} />;
  const visible = (
    <ExpressionField label="Visible when" value={control.visibleWhen} fields={fields} placeholder='Always. e.g. [Status] = "Rejected"' onChange={(v) => update({ visibleWhen: v || undefined })} />
  );

  if (control.type === 'field') {
    if (!field) {
      return <MessageBar>The column &quot;{control.field}&quot; no longer exists in the list. Remove this control.</MessageBar>;
    }
    const variants = variantOptions(field);
    const editable = !field.readOnly;
    return (
      <>
        <Group title="Field">
          <div className={styles.fieldInfo}>
            <div>
              <strong>{field.title}</strong>
            </div>
            <div className={styles.help}>
              [{field.internalName}] · {field.typeAsString}
              {field.required ? ' · required in the list' : ''}
              {field.readOnly ? ' · read-only' : ''}
            </div>
          </div>
          <TextField label="Label" placeholder={field.title} value={control.label || ''} onChange={(_e, v) => update({ label: v || undefined })} />
          <TextField label="Help text" multiline autoAdjustHeight placeholder={field.description} value={control.description || ''} onChange={(_e, v) => update({ description: v || undefined })} />
          {editable && <TextField label="Placeholder" value={control.placeholder || ''} onChange={(_e, v) => update({ placeholder: v || undefined })} />}
          {variants && (
            <ChoiceGroup
              label="Show as"
              options={variants.map((v) => ({ key: v.key, text: v.text }))}
              selectedKey={control.variant && variants.some((v) => v.key === control.variant) ? control.variant : variants[0].key}
              onChange={(_e, o) => o && update({ variant: o.key as FieldVariant })}
            />
          )}
          {field.kind === 'Note' && (
            <SpinButton label="Visible lines" min={2} max={30} value={String(control.rows || (field.richText ? 6 : 4))} onChange={(_e, v) => update({ rows: Number(v) || undefined })} />
          )}
          {width}
        </Group>
        <Group title="Rules">
          {visible}
          {editable && (
            <>
              <ExpressionField
                label="Required when"
                value={control.requiredWhen}
                fields={fields}
                placeholder={field.required ? 'Always (required in the list)' : 'Never. e.g. [Amount] > 1000'}
                onChange={(v) => update({ requiredWhen: v || undefined })}
              />
              <ExpressionField label="Read-only when" value={control.readOnlyWhen} fields={fields} placeholder='Never. e.g. Not(InGroup("Approvers"))' onChange={(v) => update({ readOnlyWhen: v || undefined })} />
              <ExpressionField
                label="Default value (new items)"
                value={control.defaultValue}
                fields={fields}
                placeholder={field.defaultValue ? `List default: ${field.defaultValue}` : 'e.g. Today() or UserEmail()'}
                onChange={(v) => update({ defaultValue: v || undefined })}
              />
              <ExpressionField
                label="Calculated value"
                value={control.calculatedValue}
                fields={fields}
                placeholder="e.g. [Quantity] * [UnitPrice]"
                description="Recalculated as the form changes and saved with the item. Makes the field read-only."
                onChange={(v) => update({ calculatedValue: v || undefined })}
              />
            </>
          )}
        </Group>
        {editable && (
          <Group title="Validation" defaultOpen={!!(control.validations && control.validations.length)}>
            <RuleList
              label="Rules"
              help="Checked on save when the field has a value. Required and list constraints (length, min / max) are checked automatically."
              rules={control.validations || []}
              fields={fields}
              onChange={(rules) => update({ validations: rules.length ? rules : undefined })}
            />
          </Group>
        )}
        {editable && (field.kind === 'Lookup' || field.kind === 'LookupMulti') && (
          <LookupSettings control={control} field={field} fields={fields} listService={props.listService} update={update} />
        )}
      </>
    );
  }

  switch (control.type) {
    case 'heading':
      return (
        <Group title="Heading">
          <ExpressionField label="Text" template value={control.content} fields={fields} description="{{ }} inserts form values, e.g. Request {{[Title]}}" onChange={(v) => update({ content: v })} />
          <Dropdown
            label="Size"
            options={[1, 2, 3, 4].map((l) => ({ key: l, text: ['Large', 'Medium', 'Small', 'Label'][l - 1] }))}
            selectedKey={control.level || 2}
            onChange={(_e, o) => o && update({ level: Number(o.key) })}
          />
          {width}
          {visible}
        </Group>
      );
    case 'text':
      return (
        <Group title="Text">
          <ExpressionField
            label="Content"
            template
            multiline
            value={control.content}
            fields={fields}
            description="**bold**, *italic*, [link](https://...), lines starting with '- ' become bullets. {{ }} inserts form values."
            onChange={(v) => update({ content: v })}
          />
          {width}
          {visible}
        </Group>
      );
    case 'image':
      return (
        <Group title="Image">
          <ExpressionField label="Image URL" template value={control.imageUrl} fields={fields} placeholder="/sites/x/SiteAssets/logo.png" onChange={(v) => update({ imageUrl: v })} />
          <TextField label="Alternative text" value={control.altText || ''} onChange={(_e, v) => update({ altText: v || '' })} />
          <SpinButton label="Width (px, 0 = natural)" min={0} max={2000} step={10} value={String(control.imageWidth || 0)} onChange={(_e, v) => update({ imageWidth: Number(v) || undefined })} />
          {width}
          {visible}
        </Group>
      );
    case 'spacer':
      return (
        <Group title="Spacer">
          <SpinButton label="Height (px)" min={4} max={400} step={4} value={String(control.height || 24)} onChange={(_e, v) => update({ height: Number(v) || 24 })} />
          {visible}
        </Group>
      );
    case 'divider':
      return (
        <Group title="Divider">
          {width}
          {visible}
        </Group>
      );
    case 'button':
      return (
        <>
          <Group title="Button">
            <TextField label="Caption" value={control.label || ''} onChange={(_e, v) => update({ label: v || '' })} />
            <ChoiceGroup
              label="When clicked"
              options={[
                { key: 'save', text: 'Set fields (optional) and save' },
                { key: 'link', text: 'Open a link' },
                { key: 'print', text: 'Print the form' },
                { key: 'close', text: 'Close the form' }
              ]}
              selectedKey={control.action || 'save'}
              onChange={(_e, o) => o && update({ action: o.key as IControl['action'] })}
            />
            <ChoiceGroup
              label="Style"
              options={[
                { key: 'primary', text: 'Primary' },
                { key: 'default', text: 'Standard' }
              ]}
              selectedKey={control.buttonStyle || 'primary'}
              onChange={(_e, o) => o && update({ buttonStyle: o.key as IControl['buttonStyle'] })}
            />
            {width}
            {visible}
          </Group>
          {(control.action || 'save') === 'save' && (
            <Group title="On click">
              <AssignmentList
                label="Set fields before saving"
                help='e.g. Status = "Approved", ApprovedBy = UserEmail(), ApprovedOn = Now(). Works on display forms too, for one-click approvals.'
                items={control.assignments || []}
                fields={fields}
                onChange={(items) => update({ assignments: items })}
              />
              <Toggle label="Skip validation (e.g. Save draft)" checked={!!control.skipValidation} onChange={(_e, c) => update({ skipValidation: !!c })} />
              <ExpressionField label="Go to after saving" template value={control.url} fields={fields} placeholder="Form default" onChange={(v) => update({ url: v || undefined })} />
            </Group>
          )}
          {control.action === 'link' && (
            <Group title="Link">
              <ExpressionField label="URL" template value={control.url} fields={fields} placeholder="https://..." onChange={(v) => update({ url: v || undefined })} />
            </Group>
          )}
        </>
      );
    case 'relatedItems':
      return (
        <>
          <Group title="Related items">
            <TextField label="Caption" value={control.label || ''} onChange={(_e, v) => update({ label: v || '' })} />
            {width}
            {visible}
            <ExpressionField label="Read-only when" value={control.readOnlyWhen} fields={fields} onChange={(v) => update({ readOnlyWhen: v || undefined })} />
          </Group>
          <RelatedItemsSettings control={control} listId={props.listId} lists={props.lists} listService={props.listService} update={update} />
        </>
      );
    default:
      return <div />;
  }
}

function SectionProperties(props: IPropertiesPanelProps & { section: ISection }): React.ReactElement {
  const { section, form, fields } = props;
  const update = (patch: Partial<ISection>): void => props.onChange(updateSection(form, section.id, patch));
  return (
    <Group title="Section">
      <TextField label="Title" value={section.title || ''} onChange={(_e, v) => update({ title: v || '' })} />
      <TextField label="Description" multiline autoAdjustHeight value={section.description || ''} onChange={(_e, v) => update({ description: v || undefined })} />
      <ChoiceGroup
        label="Columns"
        options={[1, 2, 3, 4].map((n) => ({ key: String(n), text: String(n) }))}
        selectedKey={String(section.columns)}
        onChange={(_e, o) => o && update({ columns: Number(o.key) })}
        styles={{ flexContainer: { display: 'flex', gap: 16 } }}
      />
      <Toggle label="Collapsible" checked={!!section.collapsible} onChange={(_e, c) => update({ collapsible: !!c })} />
      {section.collapsible && <Toggle label="Collapsed when the form opens" checked={!!section.collapsedByDefault} onChange={(_e, c) => update({ collapsedByDefault: !!c })} />}
      <ExpressionField label="Visible when" value={section.visibleWhen} fields={fields} placeholder='Always. e.g. FormMode() <> "New"' onChange={(v) => update({ visibleWhen: v || undefined })} />
    </Group>
  );
}

function PageProperties(props: IPropertiesPanelProps & { page: IFormPage }): React.ReactElement {
  const { page, form, fields } = props;
  const update = (patch: Partial<IFormPage>): void => props.onChange(updatePage(form, page.id, patch));
  return (
    <Group title={form.layout === 'wizard' ? 'Wizard step' : 'Page'}>
      <TextField label="Title" value={page.title} onChange={(_e, v) => update({ title: v || '' })} />
      <ExpressionField label="Visible when" value={page.visibleWhen} fields={fields} placeholder='Always. e.g. InGroup("Finance")' onChange={(v) => update({ visibleWhen: v || undefined })} />
      <div className={styles.buttonRow}>
        <DefaultButton iconProps={{ iconName: 'Back' }} text="Earlier" onClick={() => props.onChange(nudgePage(form, page.id, -1))} />
        <DefaultButton iconProps={{ iconName: 'Forward' }} text="Later" onClick={() => props.onChange(nudgePage(form, page.id, 1))} />
        <DefaultButton
          iconProps={{ iconName: 'Delete' }}
          text="Delete"
          disabled={form.pages.length <= 1}
          title="Its sections move to the previous page"
          onClick={() => props.onChange(removePage(form, page.id), { kind: 'form' })}
        />
      </div>
    </Group>
  );
}

function FormProperties(props: IPropertiesPanelProps): React.ReactElement {
  const { form, fields } = props;
  const update = (patch: Partial<IFormDefinition>): void => props.onChange({ ...form, ...patch });
  const settings = (patch: Partial<IFormSettings>): void => update({ settings: { ...form.settings, ...patch } });
  const toggleMode = (mode: FormMode, on: boolean): void => {
    const modes = on ? form.modes.concat([mode]) : form.modes.filter((m) => m !== mode);
    update({ modes: ALL_MODES.filter((m) => modes.indexOf(m) >= 0) });
  };
  return (
    <>
      <Group title="Form">
        <TextField label="Name" value={form.name} onChange={(_e, v) => update({ name: v || '' })} />
        <ChoiceGroup
          label="Layout"
          options={[
            { key: 'sections', text: 'Single page', iconProps: { iconName: 'Page' } },
            { key: 'tabs', text: 'Tabs', iconProps: { iconName: 'Tab' } },
            { key: 'wizard', text: 'Wizard', iconProps: { iconName: 'Flow' } }
          ]}
          selectedKey={form.layout}
          onChange={(_e, o) => o && update({ layout: o.key as IFormDefinition['layout'] })}
        />
        <ExpressionField label="Title" template value={form.settings.title} fields={fields} placeholder="e.g. Purchase request {{[Title]}}" onChange={(v) => settings({ title: v || undefined })} />
        <ExpressionField label="Subtitle" template value={form.settings.subtitle} fields={fields} onChange={(v) => settings({ subtitle: v || undefined })} />
      </Group>
      <Group title="Who and when">
        <Label>Use this form for</Label>
        <div className={styles.checkRow}>
          {ALL_MODES.map((m) => (
            <Checkbox key={m} label={m === 'New' ? 'New items' : m === 'Edit' ? 'Editing' : 'Viewing'} checked={form.modes.indexOf(m) >= 0} onChange={(_e, c) => toggleMode(m, !!c)} />
          ))}
        </div>
        <Dropdown
          label="Only for members of"
          placeholder="Everyone"
          multiSelect
          options={props.siteGroups.map((g) => ({ key: g, text: g }))}
          selectedKeys={form.audience}
          onChange={(_e, o) => o && update({ audience: o.selected ? form.audience.concat([String(o.key)]) : form.audience.filter((a) => a !== o.key) })}
        />
        {props.contentTypes.length > 1 && (
          <Dropdown
            label="Only for content types"
            placeholder="All content types"
            multiSelect
            options={props.contentTypes.map((ct) => ({ key: ct.id, text: ct.name }))}
            selectedKeys={form.contentTypes}
            onChange={(_e, o) => o && update({ contentTypes: o.selected ? form.contentTypes.concat([String(o.key)]) : form.contentTypes.filter((c) => c !== o.key) })}
          />
        )}
        <div className={styles.help}>When a list has several forms, the first one in the list on the left that matches is used.</div>
      </Group>
      <Group title="Buttons and saving" defaultOpen={false}>
        <Toggle label="Show Save / Cancel buttons" checked={form.settings.showDefaultButtons} onChange={(_e, c) => settings({ showDefaultButtons: !!c })} />
        <TextField label="Save button caption" value={form.settings.submitLabel} onChange={(_e, v) => settings({ submitLabel: v || 'Save' })} />
        <TextField label="Cancel button caption" value={form.settings.cancelLabel} onChange={(_e, v) => settings({ cancelLabel: v || 'Cancel' })} />
        <ExpressionField label="Message after saving (page web part)" template value={form.settings.successMessage} fields={fields} onChange={(v) => settings({ successMessage: v })} />
        <ExpressionField
          label="Go to after saving"
          template
          value={form.settings.redirectUrl}
          fields={fields}
          placeholder="Back to the list"
          description="e.g. /sites/x/SitePages/Thanks.aspx?id={{[ID]}}"
          onChange={(v) => settings({ redirectUrl: v || undefined })}
        />
        <Toggle
          label="Stop people overwriting each other's changes"
          checked={form.settings.detectConflicts}
          onChange={(_e, c) => settings({ detectConflicts: !!c })}
        />
        <AssignmentList
          label="Set fields on every save"
          help='e.g. Status = If(FormMode() = "New", "Submitted", [Status])'
          items={form.settings.onSave}
          fields={fields}
          onChange={(items) => settings({ onSave: items })}
        />
        <RuleList
          label="Form rules"
          help="Checked on save after field validation, e.g. [EndDate] >= [StartDate]."
          rules={form.settings.validations}
          fields={fields}
          onChange={(rules) => settings({ validations: rules })}
        />
      </Group>
      <Group title="Appearance" defaultOpen={false}>
        <ChoiceGroup
          label="Labels"
          options={[
            { key: 'top', text: 'Above fields' },
            { key: 'left', text: 'Beside fields' }
          ]}
          selectedKey={form.settings.labelPosition}
          onChange={(_e, o) => o && settings({ labelPosition: o.key as IFormSettings['labelPosition'] })}
        />
        <SpinButton label="Maximum width (px, 0 = full)" min={0} max={3000} step={40} value={String(form.settings.maxWidth || 0)} onChange={(_e, v) => settings({ maxWidth: Number(v) || 0 })} />
        <TextField
          label="Accent colour"
          placeholder="Site theme"
          value={form.settings.accentColor || ''}
          onChange={(_e, v) => settings({ accentColor: v || undefined })}
          description="A hex colour such as #8764b8. Leave empty to use the site theme."
        />
      </Group>
    </>
  );
}

const ELEMENT_TITLES: { [type: string]: string } = {
  heading: 'Heading',
  text: 'Text',
  image: 'Image',
  divider: 'Divider',
  spacer: 'Spacer',
  button: 'Button',
  relatedItems: 'Related items'
};

export const PropertiesPanel: React.FC<IPropertiesPanelProps> = (props) => {
  const { form, selection } = props;
  let body: React.ReactNode;
  let title = 'Form settings';
  if (selection.kind === 'control') {
    const loc = findControl(form, selection.id);
    if (loc) {
      const field = props.fields.filter((f) => f.internalName === loc.control.field)[0];
      title = loc.control.type === 'field' ? loc.control.label || (field ? field.title : 'Field') : ELEMENT_TITLES[loc.control.type] || 'Element';
      body = <ControlProperties {...props} control={loc.control} section={loc.section} />;
    }
  } else if (selection.kind === 'section') {
    const loc = findSection(form, selection.id);
    if (loc) {
      title = 'Section';
      body = <SectionProperties {...props} section={loc.section} />;
    }
  } else if (selection.kind === 'page') {
    const page = form.pages.filter((p) => p.id === selection.id)[0];
    if (page) {
      title = form.layout === 'wizard' ? 'Step' : 'Page';
      body = <PageProperties {...props} page={page} />;
    }
  }
  if (!body) {
    body = <FormProperties {...props} />;
    title = 'Form settings';
  }
  return (
    <div className={styles.properties}>
      <div className={styles.propertiesHeader}>
        <span>{title}</span>
        {selection.kind !== 'form' && (
          <IconButton iconProps={{ iconName: 'Settings' }} title="Form settings" ariaLabel="Form settings" onClick={() => props.onChange(form, { kind: 'form' })} />
        )}
      </div>
      {/* Remount per selection so uncontrolled editors start from the selected element's values. */}
      <div key={selection.kind === 'form' ? 'form' : selection.id}>{body}</div>
    </div>
  );
};
