import * as React from 'react';
import {
  Checkbox,
  ChoiceGroup,
  DatePicker,
  Dropdown,
  IDropdownOption,
  Persona,
  PersonaSize,
  TextField,
  Toggle
} from '@fluentui/react';

import { parseLocalDate } from '../../form/dates';
import { displayText, IUrlValue, IUserValue } from '../../form/formValues';
import { sanitizeHtml } from '../richContent';
import { AttachmentsControl } from './controls/AttachmentsControl';
import { LookupPicker } from './controls/LookupPicker';
import { PeoplePicker } from './controls/PeoplePicker';
import { RichTextEditor } from './controls/RichTextEditor';
import { TaxonomyPicker } from './controls/TaxonomyPicker';
import { IFieldControlProps } from './controls/types';
import styles from './FormRenderer.module.scss';

const FILL_IN_KEY = '__fillin__';

function pad(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

function ReadOnlyValue(props: IFieldControlProps): React.ReactElement {
  const { field, value } = props;
  const text = displayText(field, value);
  if (field.kind === 'User' || field.kind === 'UserMulti') {
    const users = (Array.isArray(value) ? value : value ? [value] : []) as IUserValue[];
    return (
      <div className={styles.readOnly}>
        {users.length ? users.map((u) => <Persona key={u.id || u.email} text={u.title} secondaryText={u.email} size={PersonaSize.size24} />) : <span className={styles.muted}>-</span>}
      </div>
    );
  }
  if (field.kind === 'URL' && value) {
    const url = value as IUrlValue;
    if (field.urlFormat === 1) {
      return <img className={styles.picture} src={url.url} alt={url.description} />;
    }
    return (
      <div className={styles.readOnly}>
        <a href={url.url} target="_blank" rel="noopener noreferrer" data-interception="off">
          {url.description || url.url}
        </a>
      </div>
    );
  }
  if (field.kind === 'Note' && field.richText && typeof value === 'string') {
    return <div className={styles.readOnly} dangerouslySetInnerHTML={{ __html: sanitizeHtml(value) }} />;
  }
  if (field.kind === 'Attachments') {
    return <AttachmentsControl {...props} />;
  }
  return <div className={`${styles.readOnly} ${field.kind === 'Note' ? styles.preWrap : ''}`}>{text || <span className={styles.muted}>-</span>}</div>;
}

function NumberInput(props: IFieldControlProps): React.ReactElement {
  const { field, value, onChange, id, control } = props;
  const scale = field.showAsPercentage ? 100 : 1;
  const toText = (v: unknown): string => (typeof v === 'number' ? String(Math.round(v * scale * 1e8) / 1e8) : '');
  const [text, setText] = React.useState<string>(toText(value));
  // Follow outside changes (calculations, defaults) without fighting the person typing.
  React.useEffect(() => {
    const parsed = text.trim() === '' ? null : Number(text.replace(/,/g, '')) / scale;
    if (parsed !== value && !(parsed === null && (value === null || value === undefined))) {
      setText(toText(value));
    }
  }, [value]);
  return (
    <TextField
      id={id}
      ariaLabel={props.label}
      value={text}
      inputMode="decimal"
      placeholder={control.placeholder}
      suffix={field.showAsPercentage ? '%' : undefined}
      onChange={(_e, v) => {
        const next = v || '';
        setText(next);
        if (next.trim() === '') {
          onChange(null);
          return;
        }
        const n = Number(next.replace(/,/g, ''));
        onChange(isNaN(n) ? NaN : n / scale);
      }}
    />
  );
}

function DateInput(props: IFieldControlProps): React.ReactElement {
  const { field, value, onChange, id, control } = props;
  const date = value instanceof Date ? value : null;
  const time = date ? `${pad(date.getHours())}:${pad(date.getMinutes())}` : '';
  return (
    <div className={styles.dateRow}>
      <DatePicker
        textField={{ id }}
        ariaLabel={props.label}
        value={date || undefined}
        placeholder={control.placeholder || 'Select a date...'}
        allowTextInput
        formatDate={(d) => (d ? d.toLocaleDateString() : '')}
        parseDateFromString={(text) => parseLocalDate(text)}
        onSelectDate={(d) => {
          if (!d) {
            onChange(null);
            return;
          }
          const next = new Date(d.getFullYear(), d.getMonth(), d.getDate());
          if (!field.dateOnly && date) {
            next.setHours(date.getHours(), date.getMinutes());
          }
          onChange(next);
        }}
        className={styles.datePicker}
      />
      {!field.dateOnly && (
        <TextField
          type="time"
          ariaLabel={`${props.label} time`}
          value={time}
          disabled={!date}
          onChange={(_e, v) => {
            if (!date || !v) {
              return;
            }
            const [h, m] = v.split(':').map(Number);
            const next = new Date(date.getTime());
            next.setHours(h || 0, m || 0, 0, 0);
            onChange(next);
          }}
          className={styles.timeInput}
        />
      )}
      {date && (
        <button type="button" className={styles.linkButton} onClick={() => onChange(null)} aria-label={`Clear ${props.label}`}>
          Clear
        </button>
      )}
    </div>
  );
}

function ChoiceInput(props: IFieldControlProps): React.ReactElement {
  const { field, value, onChange, control, id } = props;
  const choices = field.choices || [];
  const current = typeof value === 'string' ? value : '';
  const isFillIn = !!current && choices.indexOf(current) < 0;
  const [fillInMode, setFillInMode] = React.useState<boolean>(isFillIn);
  const options: IDropdownOption[] = choices.map((c) => ({ key: c, text: c }));
  if (field.fillIn) {
    options.push({ key: FILL_IN_KEY, text: 'Other (specify)' });
  }
  if (!props.state.required && control.variant !== 'radio') {
    options.unshift({ key: '', text: '(none)' });
  }
  const selectedKey = fillInMode ? FILL_IN_KEY : current;
  const pick = (key: string): void => {
    if (key === FILL_IN_KEY) {
      setFillInMode(true);
      onChange(null);
    } else {
      setFillInMode(false);
      onChange(key || null);
    }
  };
  return (
    <>
      {control.variant === 'radio' ? (
        <ChoiceGroup
          id={id}
          ariaLabelledBy={`${id}-label`}
          options={options.map((o) => ({ key: String(o.key), text: o.text }))}
          selectedKey={selectedKey}
          onChange={(_e, o) => o && pick(String(o.key))}
        />
      ) : (
        <Dropdown
          id={id}
          ariaLabel={props.label}
          placeholder={control.placeholder || 'Select an option'}
          options={options}
          selectedKey={selectedKey || null}
          onChange={(_e, o) => o && pick(String(o.key))}
        />
      )}
      {fillInMode && (
        <TextField
          className={styles.fillIn}
          ariaLabel={`${props.label} - other value`}
          placeholder="Enter a value"
          value={current}
          onChange={(_e, v) => onChange(v || null)}
        />
      )}
    </>
  );
}

function MultiChoiceInput(props: IFieldControlProps): React.ReactElement {
  const { field, value, onChange, control, id } = props;
  const selected = (Array.isArray(value) ? value : []) as string[];
  const choices = (field.choices || []).slice();
  selected.forEach((s) => {
    if (choices.indexOf(s) < 0) {
      choices.push(s);
    }
  });
  const toggle = (choice: string, checked: boolean): void => {
    onChange(checked ? selected.concat([choice]) : selected.filter((s) => s !== choice));
  };
  if (control.variant === 'dropdown') {
    return (
      <Dropdown
        id={id}
        ariaLabel={props.label}
        multiSelect
        placeholder={control.placeholder || 'Select options'}
        options={choices.map((c) => ({ key: c, text: c }))}
        selectedKeys={selected}
        onChange={(_e, o) => o && toggle(String(o.key), !!o.selected)}
      />
    );
  }
  return (
    <div className={styles.checkboxes} role="group" aria-labelledby={`${id}-label`}>
      {choices.map((c) => (
        <Checkbox key={c} label={c} checked={selected.indexOf(c) >= 0} onChange={(_e, checked) => toggle(c, !!checked)} />
      ))}
    </div>
  );
}

function UrlInput(props: IFieldControlProps): React.ReactElement {
  const url = (props.value as IUrlValue) || { url: '', description: '' };
  const set = (patch: Partial<IUrlValue>): void => {
    const next = { ...url, ...patch };
    props.onChange(next.url || next.description ? next : null);
  };
  return (
    <div className={styles.urlRow}>
      <TextField id={props.id} ariaLabel={`${props.label} address`} placeholder="https://" value={url.url} onChange={(_e, v) => set({ url: v || '' })} />
      <TextField ariaLabel={`${props.label} description`} placeholder="Description" value={url.description} onChange={(_e, v) => set({ description: v || '' })} />
    </div>
  );
}

/** Editor for one list field, picked by the field's type and the control's variant. */
export function FieldEditor(props: IFieldControlProps): React.ReactElement {
  const { field, value, onChange, control, state, id } = props;
  if (state.readOnly) {
    return <ReadOnlyValue {...props} />;
  }
  switch (field.kind) {
    case 'Text':
      return <TextField id={id} ariaLabel={props.label} value={(value as string) || ''} placeholder={control.placeholder} maxLength={field.maxLength} onChange={(_e, v) => onChange(v || null)} />;
    case 'Note':
      if (field.richText) {
        return <RichTextEditor id={id} label={props.label} value={(value as string) || ''} rows={control.rows || 6} onChange={(v) => onChange(v || null)} />;
      }
      return (
        <TextField
          id={id}
          ariaLabel={props.label}
          multiline
          autoAdjustHeight
          rows={control.rows || 4}
          value={(value as string) || ''}
          placeholder={control.placeholder}
          onChange={(_e, v) => onChange(v || null)}
        />
      );
    case 'Number':
    case 'Currency':
    case 'Integer':
      return <NumberInput {...props} />;
    case 'Choice':
      return <ChoiceInput {...props} />;
    case 'MultiChoice':
      return <MultiChoiceInput {...props} />;
    case 'Boolean':
      return control.variant === 'checkbox' ? (
        <Checkbox id={id} ariaLabel={props.label} checked={!!value} onChange={(_e, c) => onChange(!!c)} />
      ) : (
        <Toggle id={id} ariaLabel={props.label} checked={!!value} onText="Yes" offText="No" onChange={(_e, c) => onChange(!!c)} />
      );
    case 'DateTime':
      return <DateInput {...props} />;
    case 'Lookup':
    case 'LookupMulti':
      return <LookupPicker {...props} />;
    case 'User':
    case 'UserMulti':
      return <PeoplePicker {...props} />;
    case 'URL':
      return <UrlInput {...props} />;
    case 'TaxonomyFieldType':
    case 'TaxonomyFieldTypeMulti':
      return <TaxonomyPicker {...props} />;
    case 'Attachments':
      return <AttachmentsControl {...props} />;
    default:
      return <ReadOnlyValue {...props} />;
  }
}
