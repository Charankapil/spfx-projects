import * as React from 'react';
import { ITag, TagPicker } from '@fluentui/react';

import { ILookupValue } from '../../../form/formValues';
import { odataString } from '../../../services/http';
import { IFieldControlProps } from './types';

/**
 * Searchable lookup picker. Options can be narrowed by a cascading filter
 * (source column = value of another field on this form) and by a custom
 * OData filter with {{ expression }} placeholders.
 */
export const LookupPicker: React.FC<IFieldControlProps> = (props) => {
  const { field, value, onChange, listService, control, engine, values, id } = props;
  const multi = field.kind === 'LookupMulti';
  const selected = (Array.isArray(value) ? value : value ? [value] : []) as ILookupValue[];

  const buildFilter = async (): Promise<string | undefined | null> => {
    const parts: string[] = [];
    if (control.lookupCascade && control.lookupCascade.sourceField && control.lookupCascade.formField && listService) {
      const parentValue = values[control.lookupCascade.formField];
      const cascade = await listService.cascadeFilter(field, control.lookupCascade.sourceField, parentValue);
      if (!cascade) {
        // Nothing chosen in the parent field yet: no options.
        return null;
      }
      parts.push(cascade);
    }
    if (control.lookupFilter && control.lookupFilter.trim()) {
      parts.push(engine.template(control.lookupFilter, values, odataString));
    }
    return parts.length ? parts.join(' and ') : undefined;
  };

  const resolve = async (filter: string): Promise<ITag[]> => {
    if (!listService) {
      return [];
    }
    try {
      const extra = await buildFilter();
      if (extra === null) {
        return [];
      }
      const options = await listService.searchLookup(field, filter, extra);
      return options
        .filter((o) => !selected.some((s) => s.id === o.id))
        .map((o) => ({ key: o.id, name: o.title || `#${o.id}` }));
    } catch {
      return [];
    }
  };

  // Clear a cascading lookup when the field it depends on changes to something else.
  const parentKey = control.lookupCascade ? JSON.stringify(values[control.lookupCascade.formField] || null) : '';
  const firstRun = React.useRef(true);
  React.useEffect(() => {
    if (firstRun.current) {
      firstRun.current = false;
      return;
    }
    if (control.lookupCascade && selected.length) {
      onChange(multi ? [] : null);
    }
  }, [parentKey]);

  return (
    <TagPicker
      inputProps={{ id, 'aria-label': props.label, placeholder: control.placeholder || (selected.length ? '' : 'Search...') }}
      selectedItems={selected.map((s) => ({ key: s.id, name: s.title || `#${s.id}` }))}
      onResolveSuggestions={resolve}
      onEmptyResolveSuggestions={() => resolve('')}
      onChange={(items) => {
        const next = (items || []).map((t) => ({ id: Number(t.key), title: t.name }));
        onChange(multi ? next : next[0] || null);
      }}
      itemLimit={multi ? undefined : 1}
      resolveDelay={300}
      pickerSuggestionsProps={{ noResultsFoundText: 'No matching items', loadingText: 'Loading...' }}
    />
  );
};
