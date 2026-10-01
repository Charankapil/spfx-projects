import * as React from 'react';
import {
  Dialog, DialogType, DialogFooter, PrimaryButton, DefaultButton, NormalPeoplePicker,
  IPersonaProps, TooltipHost, Spinner, SpinnerSize
} from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { IPerson, ILabel } from '../../models/types';
import { initials, personColor } from '../../services/PeopleService';
import { PALETTE, textOn } from '../../models/colors';
import { useApp } from '../AppContext';

export function Avatar(props: { person: IPerson; size?: number }): JSX.Element {
  const size = props.size || 26;
  return (
    <TooltipHost content={props.person.title}>
      <span
        className={styles.avatar}
        style={{ background: personColor(props.person.id), width: size, height: size, fontSize: Math.round(size * 0.38) }}
        aria-label={props.person.title}
      >
        {initials(props.person.title)}
      </span>
    </TooltipHost>
  );
}

export function Avatars(props: { people: IPerson[]; max?: number; size?: number }): JSX.Element {
  const max = props.max || 3;
  const shown = props.people.slice(0, max);
  const extra = props.people.length - shown.length;
  if (props.people.length === 0) {
    return (
      <span className={styles.avatars}>
        <span className={`${styles.avatar} ${styles.avatarEmpty}`} aria-label="Nobody assigned">＋</span>
      </span>
    );
  }
  return (
    <span className={styles.avatars}>
      {shown.map(p => <Avatar key={p.id} person={p} size={props.size} />)}
      {extra > 0 && (
        <TooltipHost content={props.people.slice(max).map(p => p.title).join(', ')}>
          <span className={`${styles.avatar} ${styles.avatarEmpty}`}>+{extra}</span>
        </TooltipHost>
      )}
    </span>
  );
}

export function LabelChip(props: { label: ILabel | null }): JSX.Element | null {
  if (!props.label) {
    return null;
  }
  return (
    <span className={styles.chip} style={{ background: props.label.color, color: textOn(props.label.color) }}>
      {props.label.text}
    </span>
  );
}

export function Swatches(props: { value: string; onChange: (color: string) => void }): JSX.Element {
  return (
    <div className={styles.swatches} role="radiogroup" aria-label="Colour">
      {PALETTE.map(c => (
        <button
          key={c.value}
          type="button"
          role="radio"
          aria-checked={c.value === props.value}
          aria-label={c.name}
          title={c.name}
          className={`${styles.swatch} ${c.value === props.value ? styles.swatchOn : ''}`}
          style={{ background: c.value }}
          onClick={() => props.onChange(c.value)}
        />
      ))}
    </div>
  );
}

/** People picker backed by SharePoint's people search. Returns site users (with ids). */
export function PeoplePicker(props: {
  selected: IPerson[];
  onChange: (people: IPerson[]) => void;
  label?: string;
  max?: number;
  autoFocus?: boolean;
}): JSX.Element {
  const { services } = useApp();
  const [error, setError] = React.useState('');
  const toPersona = (p: IPerson): IPersonaProps => ({ key: String(p.id), text: p.title, secondaryText: p.email, id: String(p.id) });

  const onResolve = async (filter: string): Promise<IPersonaProps[]> => {
    try {
      const found = await services.people.search(filter);
      return found.map(f => ({ key: f.key, text: f.title, secondaryText: f.email, id: 'key:' + f.key }));
    } catch (e) {
      setError((e as Error).message);
      return [];
    }
  };

  const onItemsChange = async (items?: IPersonaProps[]): Promise<void> => {
    setError('');
    try {
      const people: IPerson[] = [];
      for (const it of items || []) {
        if (it.id && it.id.indexOf('key:') === 0) {
          people.push(await services.people.ensure(it.id.substring(4)));
        } else {
          const existing = props.selected.filter(p => String(p.id) === it.id)[0];
          if (existing) {
            people.push(existing);
          }
        }
      }
      props.onChange(people);
    } catch (e) {
      setError('That person could not be added: ' + (e as Error).message);
    }
  };

  return (
    <div>
      {props.label && <label className={styles.small}>{props.label}</label>}
      <NormalPeoplePicker
        onResolveSuggestions={onResolve}
        selectedItems={props.selected.map(toPersona)}
        onChange={items => { onItemsChange(items).catch(() => undefined); }}
        itemLimit={props.max}
        resolveDelay={300}
        pickerSuggestionsProps={{ noResultsFoundText: 'No people found', loadingText: 'Searching…', suggestionsHeaderText: 'People' }}
        inputProps={{ 'aria-label': props.label || 'People', placeholder: props.selected.length ? '' : 'Type a name' }}
        getTextFromItem={p => p.text || ''}
        componentRef={props.autoFocus ? r => { if (r) { setTimeout(() => r.focusInput(), 0); } } : undefined}
      />
      {error && <div className={styles.small} style={{ color: 'var(--wb-danger)' }}>{error}</div>}
    </div>
  );
}

export interface IConfirmOptions {
  title: string;
  message: string;
  confirmText: string;
  danger?: boolean;
}

/** Hook that shows a confirm dialog and resolves with the user's choice. */
export function useConfirm(): [JSX.Element | null, (o: IConfirmOptions) => Promise<boolean>] {
  const [state, setState] = React.useState<(IConfirmOptions & { resolve: (ok: boolean) => void }) | null>(null);
  const ask = React.useCallback((o: IConfirmOptions) => new Promise<boolean>(resolve => setState({ ...o, resolve })), []);
  const close = (ok: boolean): void => {
    if (state) {
      state.resolve(ok);
    }
    setState(null);
  };
  const dialog = state ? (
    <Dialog
      hidden={false}
      onDismiss={() => close(false)}
      dialogContentProps={{ type: DialogType.normal, title: state.title, subText: state.message }}
      modalProps={{ isBlocking: true }}
    >
      <DialogFooter>
        <PrimaryButton
          text={state.confirmText}
          onClick={() => close(true)}
          styles={state.danger ? { root: { background: '#d83a52', borderColor: '#d83a52' }, rootHovered: { background: '#b52e44', borderColor: '#b52e44' } } : undefined}
        />
        <DefaultButton text="Cancel" onClick={() => close(false)} />
      </DialogFooter>
    </Dialog>
  ) : null;
  return [dialog, ask];
}

export function Loading(props: { label?: string }): JSX.Element {
  return (
    <div className={styles.center}>
      <Spinner size={SpinnerSize.large} label={props.label || 'Loading…'} />
    </div>
  );
}

export function download(fileName: string, content: string, mime: string): void {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/** People for site user ids, resolved once per change of the id list. */
export function usePeople(ids: number[]): IPerson[] {
  const { services } = useApp();
  const key = ids.join(',');
  const [people, setPeople] = React.useState<IPerson[]>([]);
  React.useEffect(() => {
    let cancelled = false;
    if (ids.length === 0) {
      setPeople([]);
      return undefined;
    }
    services.people.resolve(ids).then(m => {
      if (!cancelled) {
        setPeople(ids.map(id => m.get(id) as IPerson));
      }
    }).catch(() => undefined);
    return () => { cancelled = true; };
  }, [key]);
  return people;
}

/** Pickers for the optional project roles (head, lead, sponsor). */
export function RolesEditor(props: {
  value: { head: IPerson[]; lead: IPerson[]; sponsor: IPerson[] };
  onChange: (value: { head: IPerson[]; lead: IPerson[]; sponsor: IPerson[] }) => void;
  disabled?: boolean;
}): JSX.Element {
  const { value } = props;
  if (props.disabled) {
    return (
      <div style={{ display: 'grid', gap: 8 }}>
        {(['head', 'lead', 'sponsor'] as const).map(r => (
          <div key={r} className={styles.row}>
            <span className={styles.muted} style={{ width: 130 }}>{r === 'head' ? 'Project head' : r === 'lead' ? 'Project lead' : 'Project sponsor'}</span>
            {value[r].length ? <Avatars people={value[r]} max={5} /> : <span className={styles.muted}>Not set</span>}
            <span>{value[r].map(p => p.title).join(', ')}</span>
          </div>
        ))}
      </div>
    );
  }
  return (
    <div style={{ display: 'grid', gap: 8 }}>
      <PeoplePicker label="Project head (optional)" selected={value.head} onChange={p => props.onChange({ ...value, head: p })} />
      <PeoplePicker label="Project lead (optional)" selected={value.lead} onChange={p => props.onChange({ ...value, lead: p })} />
      <PeoplePicker label="Project sponsor (optional)" selected={value.sponsor} onChange={p => props.onChange({ ...value, sponsor: p })} />
    </div>
  );
}
