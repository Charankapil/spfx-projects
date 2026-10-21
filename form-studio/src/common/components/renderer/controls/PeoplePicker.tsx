import * as React from 'react';
import { IPersonaProps, NormalPeoplePicker, ValidationState } from '@fluentui/react';

import { IUserValue } from '../../../form/formValues';
import { IPersonSuggestion } from '../../../services/ListService';
import { IFieldControlProps } from './types';

interface IPersona extends IPersonaProps {
  user?: IUserValue;
  suggestion?: IPersonSuggestion;
}

function toPersona(u: IUserValue): IPersona {
  return { key: String(u.id || u.email), text: u.title || u.email, secondaryText: u.email, user: u };
}

export const PeoplePicker: React.FC<IFieldControlProps> = (props) => {
  const { field, value, onChange, listService, id } = props;
  const multi = field.kind === 'UserMulti';
  const users = (Array.isArray(value) ? value : value ? [value] : []) as IUserValue[];
  const [error, setError] = React.useState<string | undefined>();

  const resolve = async (filter: string): Promise<IPersona[]> => {
    if (!listService || filter.length < 2) {
      return [];
    }
    try {
      const results = await listService.searchPeople(filter, field);
      return results
        .filter((s) => !users.some((u) => (u.email && u.email === s.email) || u.title === s.title))
        .map((s) => ({ key: s.key, text: s.title, secondaryText: s.email || s.description, suggestion: s }));
    } catch {
      return [];
    }
  };

  const change = async (items?: IPersona[]): Promise<void> => {
    setError(undefined);
    const list = items || [];
    try {
      const resolved: IUserValue[] = [];
      for (const item of list) {
        if (item.user) {
          resolved.push(item.user);
        } else if (item.suggestion && listService) {
          resolved.push(await listService.resolveSuggestion(item.suggestion));
        }
      }
      onChange(multi ? resolved : resolved[0] || null);
    } catch (e) {
      setError(`Could not add that person: ${(e as Error).message}`);
    }
  };

  return (
    <>
      <NormalPeoplePicker
        inputProps={{ id, 'aria-label': props.label, placeholder: props.control.placeholder || (users.length ? '' : 'Enter a name or e-mail address') }}
        selectedItems={users.map(toPersona)}
        onResolveSuggestions={resolve}
        onChange={(items) => {
          change(items as IPersona[]).catch(() => undefined);
        }}
        itemLimit={multi ? undefined : 1}
        resolveDelay={300}
        onValidateInput={() => ValidationState.invalid}
        pickerSuggestionsProps={{ suggestionsHeaderText: 'Suggested people', noResultsFoundText: 'No one found', loadingText: 'Searching...' }}
      />
      {error && <div role="alert" style={{ color: '#a4262c', fontSize: 12 }}>{error}</div>}
    </>
  );
};
