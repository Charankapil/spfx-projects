import * as React from 'react';
import {
  CheckboxVisibility, ConstrainMode, DetailsList, DetailsListLayoutMode, IColumn, IconButton, SearchBox, SelectionMode, Text
} from '@fluentui/react';
import { useApp } from '../AppContext';
import { CellValue, downloadCsv, IExportColumn, IExportTable } from '../../export/tabular';
import { fmtNumber } from '../../logic/format';

export interface IColumnSpec<T> {
  key: string;
  header: string;
  /** Raw value used for sorting, searching and export. */
  value: (item: T) => CellValue;
  /** Optional custom cell rendering; defaults to the formatted value. */
  render?: (item: T) => React.ReactNode;
  minWidth?: number;
  maxWidth?: number;
  isDate?: boolean;
  isNumeric?: boolean;
  /** Excluded from the grid but kept in exports. */
  exportOnly?: boolean;
  /** Excel column width in characters. */
  width?: number;
}

export function toExportTable<T>(name: string, columns: IColumnSpec<T>[], items: T[]): IExportTable {
  const cols: IExportColumn[] = columns.map((c) => ({ header: c.header, isDate: c.isDate, width: c.width, pdfHidden: !!c.exportOnly }));
  return { name, columns: cols, rows: items.map((i) => columns.map((c) => c.value(i))) };
}

function defaultRender(v: CellValue, isDate?: boolean): string {
  if (v === undefined || v === null || v === '') return '—';
  if (typeof v === 'boolean') return v ? 'Yes' : 'No';
  if (typeof v === 'number') return fmtNumber(v);
  if (isDate) return String(v).slice(0, 10);
  return String(v);
}

export interface IDataGridProps<T> {
  title: string;
  items: T[];
  columns: IColumnSpec<T>[];
  /** Filters rendered in the grid toolbar (dropdowns, toggles). */
  filters?: React.ReactNode;
  onItemInvoked?: (item: T) => void;
  initialSort?: { key: string; descending: boolean };
  height?: number;
  emptyText?: string;
}

/** Sortable, searchable, virtualised grid with CSV export of the current view. */
export function DataGrid<T>(props: IDataGridProps<T>): React.ReactElement {
  const { theme } = useApp();
  const [query, setQuery] = React.useState('');
  const [sort, setSort] = React.useState<{ key: string; descending: boolean } | undefined>(props.initialSort);
  const visibleColumns = props.columns.filter((c) => !c.exportOnly);

  const view = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    let rows = props.items;
    if (q) {
      rows = rows.filter((item) =>
        props.columns.some((c) => {
          const v = c.value(item);
          return v !== undefined && v !== null && String(v).toLowerCase().indexOf(q) >= 0;
        })
      );
    }
    if (sort) {
      const col = props.columns.filter((c) => c.key === sort.key)[0];
      if (col) {
        const dir = sort.descending ? -1 : 1;
        rows = rows.slice().sort((a, b) => {
          const va = col.value(a);
          const vb = col.value(b);
          if (va === vb) return 0;
          if (va === undefined || va === '' || va === null) return 1; // blanks last
          if (vb === undefined || vb === '' || vb === null) return -1;
          if (typeof va === 'number' && typeof vb === 'number') return (va - vb) * dir;
          return String(va).localeCompare(String(vb), undefined, { numeric: true, sensitivity: 'base' }) * dir;
        });
      }
    }
    return rows;
  }, [props.items, props.columns, query, sort]);

  const columns: IColumn[] = visibleColumns.map((c) => ({
    key: c.key,
    name: c.header,
    fieldName: c.key,
    minWidth: c.minWidth || (c.isNumeric ? 70 : 110),
    maxWidth: c.maxWidth || (c.isNumeric ? 110 : c.isDate ? 110 : 320),
    isResizable: true,
    isSorted: !!sort && sort.key === c.key,
    isSortedDescending: !!sort && sort.key === c.key && sort.descending,
    onColumnClick: () =>
      setSort((s) => (s && s.key === c.key ? { key: c.key, descending: !s.descending } : { key: c.key, descending: !!(c.isNumeric || c.isDate) })),
    onRender: (item: T) => (
      <div style={{ textAlign: c.isNumeric ? 'right' : 'left', fontVariantNumeric: c.isNumeric ? 'tabular-nums' : undefined, whiteSpace: 'normal', wordBreak: 'break-word' }}>
        {c.render ? c.render(item) : defaultRender(c.value(item), c.isDate)}
      </div>
    ),
    styles: c.isNumeric ? { cellTitle: { justifyContent: 'flex-end' } } : undefined
  }));

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 8, minWidth: 0 }}>
      <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 8 }}>
        <SearchBox
          placeholder="Search"
          value={query}
          onChange={(_, v) => setQuery(v || '')}
          styles={{ root: { width: 220 } }}
          ariaLabel={`Search ${props.title}`}
        />
        {props.filters}
        <span style={{ flex: 1 }} />
        <Text variant="small" style={{ color: theme.semanticColors.bodySubtext, fontVariantNumeric: 'tabular-nums' }}>
          {view.length === props.items.length ? `${fmtNumber(view.length)} rows` : `${fmtNumber(view.length)} of ${fmtNumber(props.items.length)} rows`}
        </Text>
        <IconButton
          iconProps={{ iconName: 'Download' }}
          title="Download this view as CSV"
          ariaLabel="Download this view as CSV"
          disabled={!view.length}
          onClick={() => downloadCsv(toExportTable(props.title, props.columns, view))}
        />
      </div>
      <div
        data-is-scrollable="true"
        style={{ maxHeight: props.height || 520, overflow: 'auto', position: 'relative', border: `1px solid ${theme.palette.neutralLighter}`, borderRadius: 4 }}
      >
        {view.length ? (
          <DetailsList
            items={view}
            columns={columns}
            selectionMode={SelectionMode.none}
            checkboxVisibility={CheckboxVisibility.hidden}
            layoutMode={DetailsListLayoutMode.justified}
            constrainMode={ConstrainMode.unconstrained}
            onItemInvoked={props.onItemInvoked}
            onRenderRow={(rowProps, defaultRender) => {
              if (!rowProps || !defaultRender) return null;
              const invoke = props.onItemInvoked;
              return invoke ? <div onClick={() => invoke(rowProps.item as T)}>{defaultRender(rowProps)}</div> : defaultRender(rowProps);
            }}
            compact
            ariaLabelForGrid={props.title}
            styles={{ root: { cursor: props.onItemInvoked ? 'pointer' : 'default' } }}
          />
        ) : (
          <Text variant="small" style={{ padding: 16, display: 'block', color: theme.semanticColors.bodySubtext }}>
            {props.items.length ? 'No rows match the search.' : props.emptyText || 'Nothing to show.'}
          </Text>
        )}
      </div>
    </div>
  );
}
