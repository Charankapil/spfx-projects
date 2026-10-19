import * as React from 'react';
import { Checkbox } from '@fluentui/react/lib/Checkbox';

import { IDocumentRecord, IMetadataField } from '../models/IMetadata';
import { facetValues, FilterState } from '../services/grouping';
import styles from './MetadataExplorer.module.scss';

export interface IFacetPanelProps {
  fields: IMetadataField[];
  /** Documents after the other fields' filters, so counts show what a tick would leave. */
  documentsFor: (fieldKey: string) => IDocumentRecord[];
  filters: FilterState;
  onToggle: (fieldKey: string, value: string) => void;
}

const COLLAPSED_VALUES = 8;

const Facet = (props: {
  field: IMetadataField;
  documents: IDocumentRecord[];
  selected: string[];
  onToggle: (value: string) => void;
}): JSX.Element | null => {
  const [expanded, setExpanded] = React.useState(false);
  const [open, setOpen] = React.useState(true);
  const values = React.useMemo(() => facetValues(props.documents, props.field), [props.documents, props.field]);
  // Keep ticked values visible even when they fall outside the top N.
  const shown = expanded
    ? values
    : values.filter((v, i) => i < COLLAPSED_VALUES || props.selected.indexOf(v.value) >= 0);

  return (
    <div className={styles.facet}>
      <button type="button" className={styles.facetHeader} aria-expanded={open} onClick={() => setOpen(!open)}>
        <span>{props.field.title}</span>
        <span className={styles.facetCount}>{values.length}</span>
      </button>
      {open && (
        <div className={styles.facetValues}>
          {shown.map((v) => (
            <div key={v.value} className={styles.facetValue}>
              <Checkbox
                label={v.value}
                checked={props.selected.indexOf(v.value) >= 0}
                onChange={() => props.onToggle(v.value)}
              />
              <span className={styles.facetValueCount}>{v.count}</span>
            </div>
          ))}
          {values.length > COLLAPSED_VALUES && (
            <button type="button" className={styles.linkButton} onClick={() => setExpanded(!expanded)}>
              {expanded ? 'Show fewer' : `Show all ${values.length}`}
            </button>
          )}
        </div>
      )}
    </div>
  );
};

export const FacetPanel = (props: IFacetPanelProps): JSX.Element => (
  <aside className={styles.facets} aria-label="Filter by metadata">
    {props.fields.map((field) => (
      <Facet
        key={field.key}
        field={field}
        documents={props.documentsFor(field.key)}
        selected={props.filters[field.key] || []}
        onToggle={(value) => props.onToggle(field.key, value)}
      />
    ))}
  </aside>
);
