import * as React from 'react';
import { Callout, DirectionalHint, Persona, PersonaSize, TextField } from '@fluentui/react';
import { ESignService, IPersonSuggestion } from '../services/ESignService';
import styles from './ESign.module.scss';

export interface IPersonInputProps {
  service: ESignService;
  label: string;
  name: string;
  email: string;
  onChange: (person: { name: string; email: string }) => void;
  required?: boolean;
}

/**
 * Name + email entry that suggests people from the tenant directory as you
 * type (SharePoint's people picker service, no Graph permission needed). Any
 * email address can also be typed directly.
 */
export function PersonInput(props: IPersonInputProps): JSX.Element {
  const [query, setQuery] = React.useState(props.name);
  const [suggestions, setSuggestions] = React.useState<IPersonSuggestion[]>([]);
  const [open, setOpen] = React.useState(false);
  const anchor = React.useRef<HTMLDivElement>(null);
  const latest = React.useRef(0);

  React.useEffect(() => setQuery(props.name), [props.name]);

  const search = (value: string): void => {
    setQuery(value);
    props.onChange({ name: value, email: /@/.test(value) && !props.email ? value.trim() : props.email });
    const ticket = ++latest.current;
    setTimeout(() => {
      if (ticket !== latest.current) {
        return;
      }
      props.service
        .searchPeople(value)
        .then(found => {
          if (ticket === latest.current) {
            setSuggestions(found);
            setOpen(found.length > 0);
          }
        })
        .catch(() => setSuggestions([]));
    }, 250);
  };

  return (
    <>
      <div ref={anchor}>
        <TextField
          label={props.label}
          required={props.required}
          value={query}
          placeholder="Search people or type a name"
          onChange={(_, v) => search(v || '')}
          onFocus={() => suggestions.length && setOpen(true)}
          autoComplete="off"
        />
      </div>
      {open && (
        <Callout
          target={anchor}
          isBeakVisible={false}
          directionalHint={DirectionalHint.bottomLeftEdge}
          onDismiss={() => setOpen(false)}
          setInitialFocus={false}
        >
          <div style={{ minWidth: 280 }}>
            {suggestions.map(s => (
              <button
                key={s.email}
                className={styles.browserRow}
                onClick={() => {
                  props.onChange({ name: s.name, email: s.email });
                  setQuery(s.name);
                  setOpen(false);
                }}
              >
                <Persona text={s.name} secondaryText={s.email} size={PersonaSize.size32} />
              </button>
            ))}
          </div>
        </Callout>
      )}
    </>
  );
}
