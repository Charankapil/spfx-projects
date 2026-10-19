import * as React from 'react';

import { IMetadataField, NO_VALUE } from '../models/IMetadata';
import { IGroup } from '../services/grouping';
import { DocumentTable } from './DocumentTable';
import styles from './MetadataExplorer.module.scss';

export interface IGroupListProps {
  groups: IGroup[];
  field: IMetadataField;
  chipFields: IMetadataField[];
  locationTitle: string;
  depth?: number;
}

const GroupSection = (props: IGroupListProps & { group: IGroup; startOpen: boolean }): JSX.Element => {
  const [open, setOpen] = React.useState(props.startOpen);
  const { group } = props;
  const missing = group.value === NO_VALUE;
  return (
    <section className={`${styles.group} ${props.depth ? styles.subgroup : ''}`}>
      <button type="button" className={styles.groupHeader} aria-expanded={open} onClick={() => setOpen(!open)}>
        <span className={`${styles.chevron} ${open ? styles.chevronOpen : ''}`} aria-hidden="true">
          ›
        </span>
        <span className={styles.groupField}>{props.field.title}</span>
        <span className={`${styles.groupValue} ${missing ? styles.groupMissing : ''}`}>{group.value}</span>
        <span className={styles.groupCount}>
          {group.count} {group.count === 1 ? 'document' : 'documents'}
        </span>
      </button>
      {open &&
        (group.subgroups ? (
          <GroupList
            groups={group.subgroups}
            field={props.chipFields.length ? props.chipFields[0] : props.field}
            chipFields={props.chipFields.slice(1)}
            locationTitle={props.locationTitle}
            depth={(props.depth || 0) + 1}
          />
        ) : (
          <DocumentTable documents={group.documents} chipFields={props.chipFields} locationTitle={props.locationTitle} />
        ))}
    </section>
  );
};

/**
 * Renders groups (and, for a second "then by" field, nested groups). For
 * nesting, chipFields[0] is the then-by field and the rest become row chips.
 */
export const GroupList = (props: IGroupListProps): JSX.Element => (
  <div className={styles.groups}>
    {props.groups.map((group, i) => (
      <GroupSection key={group.value} {...props} group={group} startOpen={i < 3 && !props.depth} />
    ))}
  </div>
);
