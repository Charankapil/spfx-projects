import * as React from 'react';
import { ITag, TagPicker } from '@fluentui/react';

import { ITermValue } from '../../../form/formValues';
import { IFieldControlProps } from './types';

export const TaxonomyPicker: React.FC<IFieldControlProps> = (props) => {
  const { field, value, onChange, listService, control, id } = props;
  const multi = field.kind === 'TaxonomyFieldTypeMulti';
  const selected = (Array.isArray(value) ? value : value ? [value] : []) as ITermValue[];

  const resolve = async (filter: string): Promise<ITag[]> => {
    if (!listService) {
      return [];
    }
    try {
      const terms = await listService.getTerms(field);
      const q = filter.toLowerCase();
      return terms
        .filter((t) => !selected.some((s) => s.termGuid === t.termGuid))
        .filter((t) => !q || (t.path || t.label).toLowerCase().indexOf(q) >= 0)
        .slice(0, 50)
        .map((t) => ({ key: t.termGuid, name: t.path || t.label, data: t } as ITag));
    } catch {
      return [];
    }
  };

  return (
    <TagPicker
      inputProps={{ id, 'aria-label': props.label, placeholder: control.placeholder || (selected.length ? '' : 'Search terms...') }}
      selectedItems={selected.map((t) => ({ key: t.termGuid, name: t.label }))}
      onResolveSuggestions={resolve}
      onEmptyResolveSuggestions={() => resolve('')}
      onChange={(items) => {
        const next = (items || []).map((t) => {
          const data = (t as ITag & { data?: ITermValue }).data;
          return { termGuid: String(t.key), label: data ? data.label : t.name };
        });
        onChange(multi ? next : next[0] || null);
      }}
      itemLimit={multi ? undefined : 1}
      pickerSuggestionsProps={{ noResultsFoundText: 'No matching terms', loadingText: 'Loading terms...' }}
    />
  );
};
