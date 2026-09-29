import * as React from 'react';
import { Callout, DirectionalHint, Calendar, DatePicker, DefaultButton, PrimaryButton, Checkbox, TextField, Icon, DayOfWeek } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { IBoardColumn, CellValue, IPerson, ITimelineValue, ILabel } from '../../models/types';
import { Avatars, PeoplePicker } from '../common/Common';
import { textOn } from '../../models/colors';
import { formatDay, formatRange, relativeDay, todayIso, dateToIso, isoToDate } from '../../engine/dates';
import { labelFor } from '../../engine/viewQuery';

export interface ICellProps {
  column: IBoardColumn;
  value: CellValue;
  canEdit: boolean;
  /** Can change the column itself (edit labels). */
  canManage: boolean;
  /** Accent for timeline bars (group colour). */
  accent: string;
  /** The item is finished, so its date is not shown as overdue. */
  done?: boolean;
  onChange: (value: CellValue) => void;
  onEditLabels?: (column: IBoardColumn) => void;
  /** Add a new option to a dropdown column; resolves when saved. */
  onAddOption?: (column: IBoardColumn, text: string) => Promise<void>;
  /** Left-align content (used in the item panel). */
  alignLeft?: boolean;
}

function usePopup(): [boolean, () => void, () => void, React.RefObject<HTMLButtonElement>] {
  const [open, setOpen] = React.useState(false);
  const ref = React.useRef<HTMLButtonElement>(null);
  return [open, () => setOpen(true), () => setOpen(false), ref];
}

function toDate(iso: string | null): Date | undefined {
  return iso ? isoToDate(iso) : undefined;
}

/* ---------- Status ---------- */

function StatusCell(props: ICellProps): JSX.Element {
  const [open, show, hide, ref] = usePopup();
  const label = labelFor(props.column, props.value as string);
  const bg = label ? label.color : 'transparent';
  return (
    <>
      <button
        ref={ref}
        type="button"
        className={styles.statusFill}
        style={{ background: bg, color: label ? textOn(bg) : 'var(--wb-subtle)' }}
        disabled={!props.canEdit}
        onClick={show}
        aria-label={`${props.column.title}: ${label ? label.text : 'empty'}`}
      >
        {label ? label.text : ''}
      </button>
      {open && (
        <Callout target={ref} onDismiss={hide} directionalHint={DirectionalHint.bottomCenter} setInitialFocus isBeakVisible>
          <div className={styles.popup}>
            <div className={styles.labelList}>
              {(props.column.labels || []).map((l: ILabel) => (
                <button
                  key={l.id || l.text}
                  type="button"
                  className={styles.labelOption}
                  style={{ background: l.color, color: textOn(l.color) }}
                  onClick={() => { hide(); props.onChange(l.text); }}
                >
                  {l.text}
                </button>
              ))}
              <button type="button" className={`${styles.labelOption} ${styles.labelOptionEmpty}`} onClick={() => { hide(); props.onChange(null); }}>
                Clear
              </button>
            </div>
            {props.canManage && props.onEditLabels && (
              <button type="button" className={styles.linkBtn} onClick={() => { hide(); (props.onEditLabels as (c: IBoardColumn) => void)(props.column); }}>
                <Icon iconName="Edit" /> Edit labels
              </button>
            )}
          </div>
        </Callout>
      )}
    </>
  );
}

/* ---------- People ---------- */

function PeopleCell(props: ICellProps): JSX.Element {
  const [open, show, hide, ref] = usePopup();
  const people = (props.value as IPerson[]) || [];
  return (
    <>
      <button ref={ref} type="button" className={`${styles.cellBtn} ${props.alignLeft ? styles.cellLeft : ''}`} disabled={!props.canEdit} onClick={show}
        aria-label={`${props.column.title}: ${people.map(p => p.title).join(', ') || 'nobody'}`}>
        <Avatars people={people} />
      </button>
      {open && (
        <Callout target={ref} onDismiss={hide} directionalHint={DirectionalHint.bottomCenter} setInitialFocus>
          <div className={styles.popup} style={{ width: 300 }}>
            <PeoplePicker selected={people} onChange={p => props.onChange(p)} label={props.column.title} autoFocus />
          </div>
        </Callout>
      )}
    </>
  );
}

/* ---------- Text, link, number ---------- */

function InlineTextCell(props: ICellProps): JSX.Element {
  const [editing, setEditing] = React.useState(false);
  const [draft, setDraft] = React.useState('');
  const isNumber = props.column.type === 'number';
  const raw = props.value;

  const start = (): void => {
    if (!props.canEdit) {
      return;
    }
    setDraft(raw === null || raw === undefined ? '' : String(raw));
    setEditing(true);
  };
  const commit = (): void => {
    setEditing(false);
    const trimmed = draft.trim();
    if (isNumber) {
      const n = trimmed === '' ? null : Number(trimmed.replace(',', '.'));
      if (n !== null && isNaN(n)) {
        return;
      }
      if (n !== raw) {
        props.onChange(n);
      }
    } else if (trimmed !== (raw || '')) {
      props.onChange(trimmed === '' ? null : trimmed);
    }
  };

  if (editing) {
    return (
      <input
        className={styles.inlineInput}
        autoFocus
        value={draft}
        inputMode={isNumber ? 'decimal' : undefined}
        type={props.column.type === 'link' ? 'url' : 'text'}
        onChange={e => setDraft(e.target.value)}
        onBlur={commit}
        onKeyDown={e => {
          if (e.key === 'Enter') {
            commit();
          } else if (e.key === 'Escape') {
            setEditing(false);
          }
        }}
        aria-label={props.column.title}
      />
    );
  }

  let display: React.ReactNode = raw === null || raw === undefined ? '' : String(raw);
  if (isNumber && typeof raw === 'number') {
    const d = props.column.decimals;
    display = (d !== undefined ? raw.toFixed(d) : String(raw)) + (props.column.unit ? ' ' + props.column.unit : '');
  }
  if (props.column.type === 'link' && raw) {
    const href = /^https?:\/\//i.test(String(raw)) ? String(raw) : 'https://' + String(raw);
    display = (
      <a href={href} target="_blank" rel="noopener noreferrer" onClick={e => e.stopPropagation()} className={styles.ellipsis}>
        {String(raw).replace(/^https?:\/\//i, '')}
      </a>
    );
  }
  const align = isNumber && !props.alignLeft ? styles.cellRight : props.alignLeft || props.column.type !== 'number' ? styles.cellLeft : '';
  return (
    <button type="button" className={`${styles.cellBtn} ${align}`} disabled={!props.canEdit} onClick={start} aria-label={`${props.column.title}: ${raw || 'empty'}`}>
      <span className={styles.ellipsis}>{display}</span>
    </button>
  );
}

/* ---------- Long text ---------- */

function LongTextCell(props: ICellProps): JSX.Element {
  const [open, show, hide, ref] = usePopup();
  const [draft, setDraft] = React.useState('');
  const value = (props.value as string) || '';
  return (
    <>
      <button ref={ref} type="button" className={`${styles.cellBtn} ${styles.cellLeft}`} disabled={!props.canEdit}
        onClick={() => { setDraft(value); show(); }} title={value} aria-label={`${props.column.title}: ${value || 'empty'}`}>
        <span className={styles.ellipsis}>{value}</span>
      </button>
      {open && (
        <Callout target={ref} onDismiss={hide} directionalHint={DirectionalHint.bottomLeftEdge} setInitialFocus>
          <div className={styles.popup} style={{ width: 340 }}>
            <TextField multiline rows={6} value={draft} onChange={(_, v) => setDraft(v || '')} label={props.column.title} autoFocus />
            <div className={styles.row}>
              <span className={styles.spacer} />
              <DefaultButton text="Cancel" onClick={hide} />
              <PrimaryButton text="Save" onClick={() => { hide(); if (draft !== value) { props.onChange(draft.trim() === '' ? null : draft); } }} />
            </div>
          </div>
        </Callout>
      )}
    </>
  );
}

/* ---------- Date ---------- */

function DateCell(props: ICellProps): JSX.Element {
  const [open, show, hide, ref] = usePopup();
  const value = (props.value as string) || null;
  const overdue = !!value && !props.done && value < todayIso();
  return (
    <>
      <button ref={ref} type="button" className={`${styles.cellBtn} ${props.alignLeft ? styles.cellLeft : ''}`} disabled={!props.canEdit} onClick={show}
        aria-label={`${props.column.title}: ${value ? formatDay(value) : 'no date'}`}>
        <span className={overdue ? styles.overdue : ''}>{relativeDay(value)}</span>
        {overdue && <Icon iconName="Warning" className={styles.overdue} aria-label="Overdue" />}
      </button>
      {open && (
        <Callout target={ref} onDismiss={hide} directionalHint={DirectionalHint.bottomCenter} setInitialFocus>
          <div className={styles.popup}>
            <Calendar
              value={toDate(value)}
              firstDayOfWeek={DayOfWeek.Monday}
              showGoToToday
              onSelectDate={d => { hide(); props.onChange(dateToIso(d)); }}
            />
            {value && <button type="button" className={styles.linkBtn} onClick={() => { hide(); props.onChange(null); }}>Clear date</button>}
          </div>
        </Callout>
      )}
    </>
  );
}

/* ---------- Timeline ---------- */

function TimelineCell(props: ICellProps): JSX.Element {
  const [open, show, hide, ref] = usePopup();
  const value = (props.value as ITimelineValue) || { start: null, end: null };
  const [draft, setDraft] = React.useState<ITimelineValue>(value);
  const has = !!(value.start || value.end);
  const overdue = !!value.end && !props.done && value.end < todayIso();
  const save = (): void => {
    hide();
    let next = draft;
    if (next.start && next.end && next.end < next.start) {
      next = { start: next.end, end: next.start };
    }
    if (next.start && !next.end) {
      next = { ...next, end: next.start };
    }
    if (next.end && !next.start) {
      next = { ...next, start: next.end };
    }
    props.onChange(next);
  };
  return (
    <>
      <button ref={ref} type="button" className={styles.cellBtn} disabled={!props.canEdit} onClick={() => { setDraft(value); show(); }}
        aria-label={`${props.column.title}: ${has ? formatRange(value.start, value.end) : 'not set'}`} style={{ padding: '0 6px' }}>
        {has ? (
          <span className={styles.timelineBar} style={{ background: overdue ? '#e2445c' : props.accent, color: textOn(overdue ? '#e2445c' : props.accent) }}>
            {formatRange(value.start, value.end)}
          </span>
        ) : (
          <span className={styles.timelineBar} style={{ background: 'var(--wb-border-light)' }} />
        )}
      </button>
      {open && (
        <Callout target={ref} onDismiss={hide} directionalHint={DirectionalHint.bottomCenter} setInitialFocus>
          <div className={styles.popup} style={{ width: 280 }}>
            <DatePicker label="Start" value={toDate(draft.start)} firstDayOfWeek={DayOfWeek.Monday}
              onSelectDate={d => setDraft({ ...draft, start: d ? dateToIso(d) : null })} formatDate={d => (d ? formatDay(dateToIso(d)) : '')} />
            <DatePicker label="End" value={toDate(draft.end)} firstDayOfWeek={DayOfWeek.Monday}
              onSelectDate={d => setDraft({ ...draft, end: d ? dateToIso(d) : null })} formatDate={d => (d ? formatDay(dateToIso(d)) : '')} />
            <div className={styles.row}>
              {has && <button type="button" className={styles.linkBtn} onClick={() => { hide(); props.onChange({ start: null, end: null }); }}>Clear</button>}
              <span className={styles.spacer} />
              <DefaultButton text="Cancel" onClick={hide} />
              <PrimaryButton text="Save" onClick={save} />
            </div>
          </div>
        </Callout>
      )}
    </>
  );
}

/* ---------- Dropdown (tags) ---------- */

function DropdownCell(props: ICellProps): JSX.Element {
  const [open, show, hide, ref] = usePopup();
  const [newOption, setNewOption] = React.useState('');
  const [saving, setSaving] = React.useState(false);
  const values = (props.value as string[]) || [];
  const toggle = (text: string): void => {
    const next = values.indexOf(text) >= 0 ? values.filter(v => v !== text) : values.concat([text]);
    props.onChange(next);
  };
  const add = async (): Promise<void> => {
    const text = newOption.trim();
    if (!text || !props.onAddOption) {
      return;
    }
    setSaving(true);
    try {
      await props.onAddOption(props.column, text);
      props.onChange(values.concat([text]));
      setNewOption('');
    } finally {
      setSaving(false);
    }
  };
  return (
    <>
      <button ref={ref} type="button" className={`${styles.cellBtn} ${props.alignLeft ? styles.cellLeft : ''}`} disabled={!props.canEdit} onClick={show}
        aria-label={`${props.column.title}: ${values.join(', ') || 'empty'}`}>
        {values.slice(0, 2).map(v => {
          const l = labelFor(props.column, v) as ILabel;
          return <span key={v} className={styles.chip} style={{ background: l.color, color: textOn(l.color) }}>{v}</span>;
        })}
        {values.length > 2 && <span className={styles.pill}>+{values.length - 2}</span>}
      </button>
      {open && (
        <Callout target={ref} onDismiss={hide} directionalHint={DirectionalHint.bottomCenter} setInitialFocus>
          <div className={styles.popup}>
            {(props.column.labels || []).length === 0 && <span className={styles.muted}>No options yet.</span>}
            {(props.column.labels || []).map(l => (
              <Checkbox key={l.id || l.text} label={l.text} checked={values.indexOf(l.text) >= 0} onChange={() => toggle(l.text)} />
            ))}
            {props.canManage && props.onAddOption && (
              <div className={styles.row} style={{ flexWrap: 'nowrap' }}>
                <TextField placeholder="New option" value={newOption} onChange={(_, v) => setNewOption(v || '')}
                  onKeyDown={e => { if (e.key === 'Enter') { add().catch(() => undefined); } }} styles={{ root: { flex: 1 } }} />
                <DefaultButton text="Add" onClick={() => { add().catch(() => undefined); }} disabled={saving || !newOption.trim()} />
              </div>
            )}
            {props.canManage && props.onEditLabels && (
              <button type="button" className={styles.linkBtn} onClick={() => { hide(); (props.onEditLabels as (c: IBoardColumn) => void)(props.column); }}>
                <Icon iconName="Edit" /> Edit options
              </button>
            )}
          </div>
        </Callout>
      )}
    </>
  );
}

/* ---------- Checkbox ---------- */

function CheckboxCell(props: ICellProps): JSX.Element {
  return (
    <div className={styles.cellBtn} style={{ cursor: 'default', justifyContent: props.alignLeft ? 'flex-start' : 'center' }}>
      <Checkbox checked={props.value === true} disabled={!props.canEdit} onChange={(_, c) => props.onChange(!!c)} ariaLabel={props.column.title} />
    </div>
  );
}

export function Cell(props: ICellProps): JSX.Element {
  switch (props.column.type) {
    case 'status':
      return <StatusCell {...props} />;
    case 'people':
      return <PeopleCell {...props} />;
    case 'longtext':
      return <LongTextCell {...props} />;
    case 'date':
      return <DateCell {...props} />;
    case 'timeline':
      return <TimelineCell {...props} />;
    case 'dropdown':
      return <DropdownCell {...props} />;
    case 'checkbox':
      return <CheckboxCell {...props} />;
    default:
      return <InlineTextCell {...props} />;
  }
}

/** Default column widths in px. */
export function columnWidth(col: IBoardColumn): number {
  if (col.width) {
    return col.width;
  }
  switch (col.type) {
    case 'people':
      return 110;
    case 'status':
      return 140;
    case 'timeline':
      return 170;
    case 'longtext':
    case 'text':
    case 'link':
      return 180;
    case 'dropdown':
      return 170;
    case 'checkbox':
      return 90;
    default:
      return 120;
  }
}

export const COLUMN_TYPES: { type: IBoardColumn['type']; title: string; icon: string; description: string }[] = [
  { type: 'status', title: 'Status', icon: 'StatusCircleRing', description: 'Coloured labels, like Working on it or Done' },
  { type: 'people', title: 'People', icon: 'People', description: 'Assign one or more people' },
  { type: 'timeline', title: 'Timeline', icon: 'Timeline', description: 'Start and end dates, shown on the Timeline view' },
  { type: 'date', title: 'Date', icon: 'Calendar', description: 'A single date' },
  { type: 'text', title: 'Text', icon: 'TextField', description: 'A short line of text' },
  { type: 'longtext', title: 'Long text', icon: 'AlignLeft', description: 'Notes or a description' },
  { type: 'number', title: 'Numbers', icon: 'NumberSymbol', description: 'Budget, hours or any number, summed per group' },
  { type: 'dropdown', title: 'Dropdown', icon: 'Tag', description: 'Pick one or more options' },
  { type: 'checkbox', title: 'Checkbox', icon: 'CheckboxComposite', description: 'Yes or no' },
  { type: 'link', title: 'Link', icon: 'Link', description: 'A web address' }
];
