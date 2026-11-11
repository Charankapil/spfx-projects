import * as React from 'react';
import { useCallback, useEffect, useRef, useState } from 'react';
import { TextField } from '@fluentui/react';

import { CancelToken, isCancelled } from '../RequestQueue';
import { SpClient } from '../SpClient';
import styles from './Toolkit.module.scss';

export interface IPerson {
  /** Login name, as SharePoint's ensureuser expects it. */
  key: string;
  name: string;
  email: string;
}

interface IEntity {
  Key: string;
  DisplayText: string;
  EntityData?: { Email?: string };
  IsResolved?: boolean;
}

interface IPickerResponse {
  value?: string;
  d?: { ClientPeoplePickerSearchUser?: string };
}

/** Wait this long after the last keystroke before asking SharePoint. */
const DEBOUNCE_MS = 600;
const MIN_CHARS = 3;

/**
 * Finds people with SharePoint's own people-picker service, so no Microsoft
 * Graph permission is needed. One request per pause in typing, never one
 * per keystroke.
 */
export const PeoplePicker: React.FC<{
  sp: SpClient;
  label: string;
  selected: IPerson | undefined;
  onChange: (person: IPerson | undefined) => void;
  disabled?: boolean;
}> = (props) => {
  const { sp, onChange } = props;
  const [text, setText] = useState('');
  const [suggestions, setSuggestions] = useState<IPerson[]>([]);
  const [searching, setSearching] = useState(false);
  const timer = useRef<number | undefined>(undefined);
  const token = useRef<CancelToken | undefined>(undefined);
  const mounted = useRef(true);

  useEffect(
    () => () => {
      mounted.current = false;
      window.clearTimeout(timer.current);
      if (token.current) {
        token.current.cancel();
      }
    },
    []
  );

  const lookup = useCallback(
    (query: string) => {
      if (token.current) {
        token.current.cancel();
      }
      const mine = new CancelToken();
      token.current = mine;
      setSearching(true);
      sp.post<IPickerResponse>(
        `${sp.webUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`,
        {
          queryParams: {
            QueryString: query,
            MaximumEntitySuggestions: 8,
            AllowEmailAddresses: false,
            AllowMultipleEntities: false,
            PrincipalType: 1,
            PrincipalSource: 15,
            SharePointGroupID: 0
          }
        },
        { token: mine }
      )
        .then((json) => {
          if (!mounted.current || mine.cancelled) {
            return;
          }
          const raw = json ? json.value || (json.d && json.d.ClientPeoplePickerSearchUser) : undefined;
          let entities: IEntity[] = [];
          try {
            entities = raw ? (JSON.parse(raw) as IEntity[]) : [];
          } catch {
            entities = [];
          }
          setSuggestions(
            entities
              .filter((e) => !!e.Key)
              .map((e) => ({ key: e.Key, name: e.DisplayText, email: (e.EntityData && e.EntityData.Email) || '' }))
          );
          setSearching(false);
        })
        .catch((err) => {
          if (mounted.current && !isCancelled(err)) {
            setSuggestions([]);
            setSearching(false);
          }
        });
    },
    [sp]
  );

  const onType = useCallback(
    (value: string) => {
      setText(value);
      window.clearTimeout(timer.current);
      if (value.trim().length < MIN_CHARS) {
        // Drop any lookup still on its way, so old matches cannot pop up under an emptied box.
        if (token.current) {
          token.current.cancel();
        }
        setSearching(false);
        setSuggestions([]);
        return;
      }
      timer.current = window.setTimeout(() => lookup(value.trim()), DEBOUNCE_MS);
    },
    [lookup]
  );

  if (props.selected) {
    return (
      <div>
        <div style={{ fontWeight: 600, padding: '5px 0' }}>{props.label}</div>
        <div className={styles.row}>
          <span className={styles.pill}>{props.selected.name}</span>
          <button type="button" className={styles.linkButton} disabled={props.disabled} onClick={() => onChange(undefined)}>
            Change
          </button>
        </div>
      </div>
    );
  }

  return (
    <div>
      <TextField
        label={props.label}
        placeholder="Type at least 3 letters of a name"
        value={text}
        onChange={(_, value) => onType(value || '')}
        disabled={props.disabled}
        autoComplete="off"
      />
      {searching && <div className={styles.small}>Searching...</div>}
      {!searching && text.trim().length >= MIN_CHARS && suggestions.length === 0 && <div className={styles.small}>No matching people yet.</div>}
      {suggestions.length > 0 && (
        <div className={styles.stack} style={{ marginTop: 6 }} role="listbox" aria-label="Matching people">
          {suggestions.map((person) => (
            <button
              key={person.key}
              type="button"
              role="option"
              aria-selected={false}
              className={styles.option}
              onClick={() => {
                setText('');
                setSuggestions([]);
                onChange(person);
              }}
            >
              <div style={{ fontWeight: 600 }}>{person.name}</div>
              {person.email && <div className={styles.small}>{person.email}</div>}
            </button>
          ))}
        </div>
      )}
    </div>
  );
};
