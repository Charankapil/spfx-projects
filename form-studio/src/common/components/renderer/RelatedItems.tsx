import * as React from 'react';
import { ActionButton, Checkbox, Dropdown, IconButton, MessageBar, MessageBarType, Spinner, SpinnerSize } from '@fluentui/react';

import { buildPayload } from '../../form/valueCodec';
import { displayText, emptyValue, IFormValues } from '../../form/formValues';
import { IFieldInfo } from '../../model/IFieldInfo';
import { IRelatedItemsConfig } from '../../model/IFormDefinition';
import { errorMessage } from '../../services/http';
import { ListService } from '../../services/ListService';
import styles from './FormRenderer.module.scss';

export interface IRelatedItemsHandle {
  /** Returns an error message when a row is incomplete. */
  validate(): string | undefined;
  /** Writes buffered row changes once the parent item has an id. */
  commit(parentId: number): Promise<void>;
}

interface IRow {
  key: string;
  id?: number;
  values: IFormValues;
  status: 'clean' | 'new' | 'changed' | 'deleted';
}

export interface IRelatedItemsProps {
  config: IRelatedItemsConfig;
  label: string;
  parentId?: number;
  readOnly: boolean;
  listService?: ListService;
  register: (handle: IRelatedItemsHandle | undefined) => void;
  onDirty: () => void;
}

const INLINE_KINDS = ['Text', 'Note', 'Number', 'Currency', 'Integer', 'Choice', 'Boolean', 'DateTime'];

function toDateInput(value: unknown): string {
  if (!(value instanceof Date)) {
    return '';
  }
  const p = (n: number): string => (n < 10 ? `0${n}` : String(n));
  return `${value.getFullYear()}-${p(value.getMonth() + 1)}-${p(value.getDate())}`;
}

let rowSeed = 0;

/**
 * Child items of another list, edited inline - order lines on an order,
 * expense lines on a claim. Rows are linked to this item by a lookup column
 * in the child list. Changes are buffered and written after the parent item
 * saves, so a new parent and its rows are created in one go.
 */
export const RelatedItems: React.FC<IRelatedItemsProps> = (props) => {
  const { config, listService, parentId, readOnly } = props;
  const [fields, setFields] = React.useState<IFieldInfo[]>([]);
  const [rows, setRows] = React.useState<IRow[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | undefined>();
  const rowsRef = React.useRef<IRow[]>(rows);
  rowsRef.current = rows;

  const columns = React.useMemo(
    () => config.columns.map((c) => fields.filter((f) => f.internalName === c)[0]).filter((f): f is IFieldInfo => !!f),
    [fields, config.columns.join('|')]
  );
  const parentField = fields.filter((f) => f.internalName === config.parentLookupField)[0];

  React.useEffect(() => {
    let cancelled = false;
    const load = async (): Promise<void> => {
      if (!listService || !config.listId || !config.parentLookupField) {
        setLoading(false);
        return;
      }
      try {
        const childFields = await listService.getFields(config.listId);
        if (cancelled) {
          return;
        }
        setFields(childFields);
        const lookup = childFields.filter((f) => f.internalName === config.parentLookupField)[0];
        if (parentId && lookup) {
          const shown = config.columns.map((c) => childFields.filter((f) => f.internalName === c)[0]).filter((f) => !!f);
          const items = await listService.getChildItems(config.listId, lookup, parentId, shown);
          if (!cancelled) {
            setRows(items.map((i) => ({ key: `r${rowSeed++}`, id: i.id, values: i.values, status: 'clean' })));
          }
        }
      } catch (e) {
        if (!cancelled) {
          setError(errorMessage(e));
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    };
    load().catch(() => undefined);
    return () => {
      cancelled = true;
    };
  }, [config.listId, config.parentLookupField, parentId]);

  React.useEffect(() => {
    props.register({
      validate: () => {
        for (const row of rowsRef.current) {
          if (row.status === 'deleted') {
            continue;
          }
          const missing = columns.filter((c) => c.required && (row.values[c.internalName] === null || row.values[c.internalName] === undefined || row.values[c.internalName] === ''))[0];
          if (missing) {
            return `${props.label}: fill in ${missing.title} on every row.`;
          }
        }
        return undefined;
      },
      commit: async (id: number) => {
        if (!listService || !parentField) {
          return;
        }
        for (const row of rowsRef.current) {
          if (row.status === 'deleted' && row.id) {
            await listService.deleteItem(config.listId, row.id);
          } else if (row.status === 'new') {
            const payload = buildPayload(columns, row.values);
            payload.body[`${parentField.entityPropertyName}Id`] = parentField.kind === 'LookupMulti' ? [id] : id;
            await listService.createItem(config.listId, payload);
          } else if (row.status === 'changed' && row.id) {
            await listService.updateItem(config.listId, row.id, buildPayload(columns, row.values));
          }
        }
      }
    });
    return () => props.register(undefined);
  }, [columns, parentField]);

  const update = (key: string, name: string, value: unknown): void => {
    setRows(rows.map((r) => (r.key === key ? { ...r, values: { ...r.values, [name]: value }, status: r.status === 'clean' ? 'changed' : r.status } : r)));
    props.onDirty();
  };

  const addRow = (): void => {
    const values: IFormValues = {};
    columns.forEach((c) => {
      values[c.internalName] = emptyValue(c);
    });
    setRows(rows.concat([{ key: `r${rowSeed++}`, values, status: 'new' }]));
    props.onDirty();
  };

  const removeRow = (key: string): void => {
    setRows(rows.filter((r) => !(r.key === key && r.status === 'new')).map((r) => (r.key === key ? { ...r, status: 'deleted' } : r)));
    props.onDirty();
  };

  const canEditRow = (row: IRow): boolean => !readOnly && (row.status === 'new' || config.allowEdit);

  const cell = (row: IRow, f: IFieldInfo): React.ReactNode => {
    const value = row.values[f.internalName];
    if (!canEditRow(row) || f.readOnly || INLINE_KINDS.indexOf(f.kind) < 0) {
      return <span>{displayText(f, value)}</span>;
    }
    const label = `${f.title}, row ${rows.filter((r) => r.status !== 'deleted').indexOf(row) + 1}`;
    switch (f.kind) {
      case 'Boolean':
        return <Checkbox ariaLabel={label} checked={!!value} onChange={(_e, c) => update(row.key, f.internalName, !!c)} />;
      case 'Choice':
        return (
          <Dropdown
            ariaLabel={label}
            options={[{ key: '', text: '' }].concat((f.choices || []).map((c) => ({ key: c, text: c })))}
            selectedKey={(value as string) || ''}
            onChange={(_e, o) => update(row.key, f.internalName, o && o.key ? String(o.key) : null)}
          />
        );
      case 'DateTime':
        return (
          <input
            className={styles.cellInput}
            type="date"
            aria-label={label}
            value={toDateInput(value)}
            onChange={(e) => {
              const parts = e.target.value.split('-').map(Number);
              update(row.key, f.internalName, e.target.value ? new Date(parts[0], parts[1] - 1, parts[2]) : null);
            }}
          />
        );
      case 'Number':
      case 'Currency':
      case 'Integer':
        return (
          <input
            className={styles.cellInput}
            type="number"
            aria-label={label}
            value={typeof value === 'number' ? value : ''}
            onChange={(e) => update(row.key, f.internalName, e.target.value === '' ? null : Number(e.target.value))}
          />
        );
      default:
        return (
          <input
            className={styles.cellInput}
            type="text"
            aria-label={label}
            value={(value as string) || ''}
            onChange={(e) => update(row.key, f.internalName, e.target.value || null)}
          />
        );
    }
  };

  if (!config.listId || !config.parentLookupField) {
    return <MessageBar>Choose the related list and its lookup column in the designer.</MessageBar>;
  }
  if (loading) {
    return <Spinner size={SpinnerSize.small} label={`Loading ${props.label.toLowerCase()}...`} />;
  }
  if (error) {
    return <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>;
  }

  const visible = rows.filter((r) => r.status !== 'deleted');
  const numeric = columns.filter((c) => c.kind === 'Number' || c.kind === 'Currency' || c.kind === 'Integer');

  return (
    <div className={styles.related}>
      <table className={styles.relatedTable}>
        <thead>
          <tr>
            {columns.map((c) => (
              <th key={c.internalName} scope="col">
                {c.title}
                {c.required && !readOnly && <span className={styles.required}> *</span>}
              </th>
            ))}
            {!readOnly && config.allowDelete && <th aria-label="Actions" />}
          </tr>
        </thead>
        <tbody>
          {visible.map((row) => (
            <tr key={row.key}>
              {columns.map((c) => (
                <td key={c.internalName}>{cell(row, c)}</td>
              ))}
              {!readOnly && config.allowDelete && (
                <td className={styles.rowActions}>
                  <IconButton iconProps={{ iconName: 'Delete' }} title="Remove row" ariaLabel="Remove row" onClick={() => removeRow(row.key)} />
                </td>
              )}
            </tr>
          ))}
          {!visible.length && (
            <tr>
              <td colSpan={columns.length + 1} className={styles.muted}>
                No items yet.
              </td>
            </tr>
          )}
        </tbody>
        {numeric.length > 0 && visible.length > 1 && (
          <tfoot>
            <tr>
              {columns.map((c) => (
                <td key={c.internalName}>
                  {numeric.indexOf(c) >= 0 ? displayText(c, visible.reduce((t, r) => t + ((r.values[c.internalName] as number) || 0), 0)) : ''}
                </td>
              ))}
              {!readOnly && config.allowDelete && <td />}
            </tr>
          </tfoot>
        )}
      </table>
      {!readOnly && config.allowAdd && (
        <ActionButton iconProps={{ iconName: 'Add' }} onClick={addRow}>
          Add row
        </ActionButton>
      )}
    </div>
  );
};
