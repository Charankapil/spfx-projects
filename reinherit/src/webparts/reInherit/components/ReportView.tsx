import * as React from 'react';
import { useEffect, useMemo, useRef, useState } from 'react';
import {
  CheckboxVisibility,
  ConstrainMode,
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  Dropdown,
  IColumn,
  Icon,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  SearchBox,
  Selection,
  SelectionMode,
  Stack
} from '@fluentui/react';

import { IRunReport, IRunStats, KIND_LABELS } from '../models/IRunReport';
import { IUniqueObject, ObjectKind, ObjectStatus } from '../models/IUniqueObject';
import { downloadCsv, STATUS_LABELS } from '../services/ExportService';
import { formatCompact, formatDate, formatDuration, formatPercent } from './format';
import styles from './ReInherit.module.scss';

export const KIND_ICONS: { [kind in ObjectKind]: string } = {
  web: 'Globe',
  library: 'DocLibrary',
  list: 'BulletedList',
  folder: 'FabricFolder',
  file: 'Page',
  item: 'CheckList'
};

const STATUS_SHORT: { [status in ObjectStatus]: string } = {
  found: 'Unique',
  restored: 'Restored',
  failed: 'Failed',
  excluded: 'Excluded',
  skipped: 'Not processed'
};

const KIND_ORDER: ObjectKind[] = ['web', 'library', 'list', 'folder', 'file', 'item'];
const ERRORS_SHOWN = 15;

export function computeStats(report: IRunReport): IRunStats {
  const stats: IRunStats = { ...report.stats, uniqueFound: report.objects.length, restored: 0, failed: 0, excluded: 0, skipped: 0 };
  for (const o of report.objects) {
    if (o.status === 'restored') {
      stats.restored++;
    } else if (o.status === 'failed') {
      stats.failed++;
    } else if (o.status === 'excluded') {
      stats.excluded++;
    } else if (o.status === 'skipped') {
      stats.skipped++;
    }
  }
  return stats;
}

function plural(n: number, noun: string): string {
  return `${formatCompact(n)} ${noun}${n === 1 ? '' : 's'}`;
}

function objectHref(o: IUniqueObject): string {
  if (o.kind === 'web') {
    return o.webUrl;
  }
  // Encode each segment: a "#" or "?" in a file name would otherwise cut the link short.
  return `${new URL(o.webUrl).origin}${o.path.split('/').map(encodeURIComponent).join('/')}`;
}

interface IReportViewProps {
  report: IRunReport;
  /** Review: before restoring, rows can be excluded / included. */
  selectable?: boolean;
  onSetExcluded?: (objects: IUniqueObject[], excluded: boolean) => void;
  actions?: React.ReactNode;
}

const Tile: React.FC<{ label: string; value: number; icon: string; tone?: string; sub?: string }> = ({
  label,
  value,
  icon,
  tone,
  sub
}) => (
  <div className={`${styles.tile} ${tone ? styles[tone] : ''}`}>
    <div className={styles.tileHead}>
      <span className={styles.tileIcon}>
        <Icon iconName={icon} />
      </span>
      {label}
    </div>
    <div className={styles.tileValue} title={value.toLocaleString()}>
      {formatCompact(value)}
    </div>
    {sub && <div className={styles.tileSub}>{sub}</div>}
  </div>
);

export const ReportView: React.FC<IReportViewProps> = ({ report, selectable, onSetExcluded, actions }) => {
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [kindFilter, setKindFilter] = useState<string>('all');
  const [selectedCount, setSelectedCount] = useState(0);
  const [showAllErrors, setShowAllErrors] = useState(false);

  const selectionRef = useRef<Selection>();
  if (!selectionRef.current) {
    selectionRef.current = new Selection({
      onSelectionChanged: () => setSelectedCount(selectionRef.current ? selectionRef.current.getSelectedCount() : 0),
      getKey: (item) => (item as unknown as IUniqueObject).key
    });
  }

  const stats = useMemo(() => computeStats(report), [report]);
  const hasRestore = report.mode === 'restore';

  const byKind = useMemo(() => {
    const map: { [kind: string]: { total: number; restored: number } } = {};
    for (const o of report.objects) {
      const entry = map[o.kind] || (map[o.kind] = { total: 0, restored: 0 });
      entry.total++;
      if (o.status === 'restored') {
        entry.restored++;
      }
    }
    return KIND_ORDER.filter((k) => map[k]).map((k) => ({ kind: k, ...map[k] }));
  }, [report]);

  const filtered = useMemo(() => {
    const term = search.trim().toLowerCase();
    return report.objects.filter(
      (o) =>
        (statusFilter === 'all' || o.status === statusFilter) &&
        (kindFilter === 'all' || o.kind === kindFilter) &&
        (!term || o.path.toLowerCase().indexOf(term) >= 0 || (o.previousPermissions || '').toLowerCase().indexOf(term) >= 0)
    );
  }, [report, search, statusFilter, kindFilter]);

  // Filters change the rows, so a stale selection would point at the wrong ones.
  useEffect(() => {
    if (selectionRef.current) {
      selectionRef.current.setAllSelected(false);
    }
  }, [search, statusFilter, kindFilter]);

  const statusOptions: IDropdownOption[] = useMemo(() => {
    const present = new Set<ObjectStatus>();
    report.objects.forEach((o) => present.add(o.status));
    return [{ key: 'all', text: 'All statuses' }].concat(
      (Object.keys(STATUS_SHORT) as ObjectStatus[])
        .filter((s) => present.has(s))
        .map((s) => ({ key: s, text: STATUS_LABELS[s] }))
    );
  }, [report]);

  const kindOptions: IDropdownOption[] = useMemo(
    () => [{ key: 'all', text: 'All types' }].concat(byKind.map((k) => ({ key: k.kind, text: KIND_LABELS[k.kind] }))),
    [byKind]
  );

  const columns: IColumn[] = useMemo(() => {
    const cols: IColumn[] = [
      {
        key: 'status',
        name: 'Status',
        minWidth: 92,
        maxWidth: 110,
        onRender: (o: IUniqueObject) => (
          <span className={`${styles.statusPill} ${styles[`status_${o.status}`]}`} title={STATUS_LABELS[o.status]}>
            {STATUS_SHORT[o.status]}
          </span>
        )
      },
      {
        key: 'name',
        name: 'Name',
        minWidth: 180,
        maxWidth: 320,
        isResizable: true,
        onRender: (o: IUniqueObject) => (
          <span className={styles.nameCell}>
            <Icon iconName={KIND_ICONS[o.kind]} className={styles.nameIcon} title={KIND_LABELS[o.kind]} />
            <a href={objectHref(o)} target="_blank" rel="noopener noreferrer" data-interception="off" title={o.path}>
              {o.name}
            </a>
          </span>
        )
      },
      {
        key: 'path',
        name: 'Location',
        minWidth: 200,
        maxWidth: 420,
        isResizable: true,
        onRender: (o: IUniqueObject) => (
          <span className={styles.pathCell} title={o.path}>
            {o.path}
          </span>
        )
      },
      {
        key: 'depth',
        name: 'Depth',
        minWidth: 44,
        maxWidth: 52,
        onRender: (o: IUniqueObject) => <span className={styles.numCell}>{o.depth || ''}</span>
      }
    ];
    if (report.options.backupPermissions && hasRestore) {
      cols.push({
        key: 'before',
        name: 'Permissions before',
        minWidth: 180,
        maxWidth: 360,
        isResizable: true,
        onRender: (o: IUniqueObject) => (
          <span className={styles.pathCell} title={o.previousPermissions}>
            {o.previousPermissions || ''}
          </span>
        )
      });
    }
    cols.push({
      key: 'message',
      name: hasRestore ? 'Details' : 'Notes',
      minWidth: 140,
      isResizable: true,
      onRender: (o: IUniqueObject) => (
        <span className={o.status === 'failed' ? `${styles.pathCell} ${styles.errorText}` : styles.pathCell} title={o.message}>
          {o.message || (o.restoredAt ? `Restored ${formatDate(new Date(o.restoredAt))}` : '')}
        </span>
      )
    });
    return cols;
  }, [report, hasRestore]);

  const selectedObjects = (): IUniqueObject[] =>
    selectionRef.current ? (selectionRef.current.getSelection() as unknown as IUniqueObject[]) : [];

  const started = new Date(report.startedAt);
  const ended = report.completedAt ? new Date(report.completedAt) : report.scanCompletedAt ? new Date(report.scanCompletedAt) : undefined;
  const pending = stats.uniqueFound - stats.restored - stats.failed - stats.excluded - stats.skipped;
  const checked = stats.itemsChecked + stats.listsChecked + stats.websChecked;

  return (
    <div className={styles.report}>
      <div className={styles.reportHeader}>
      <div className={styles.reportMeta}>
        {hasRestore ? 'Restore run' : 'Scan'} started {formatDate(started)} by {report.runBy}
        {ended && ` · took ${formatDuration(ended.getTime() - started.getTime())}`}
        {report.cancelled && ' · cancelled before it finished'}
        <div className={styles.scopeChips}>
          {report.scopes.map((s) => (
            <span key={s.url} className={styles.scopeChip} title={s.url}>
              <Icon iconName={KIND_ICONS[s.kind as ObjectKind] || 'Globe'} /> {s.title}
            </span>
          ))}
          <span className={styles.scopeChipMuted}>
            {report.options.recursive ? 'Recursive' : 'Selected objects only'}
            {report.options.recursive && report.options.maxDepth > 0 ? ` · ${report.options.maxDepth} folder level(s)` : ''}
          </span>
        </div>
      </div>
        {actions && (
          <Stack horizontal wrap tokens={{ childrenGap: 8 }} className={styles.reportActions}>
            {actions}
          </Stack>
        )}
      </div>

      <div className={styles.tiles}>
        <Tile
          label="Checked"
          icon="SearchAndApps"
          value={checked}
          sub={`${plural(stats.itemsChecked, 'item')} · ${plural(stats.listsChecked, 'list')} · ${plural(stats.websChecked, 'site')}`}
        />
        <Tile label="Unique permissions" icon="Permissions" value={stats.uniqueFound} tone="toneAccent" sub={`${formatPercent(stats.uniqueFound, checked)} of what was checked`} />
        {hasRestore ? (
          <>
            <Tile label="Restored" icon="CompletedSolid" value={stats.restored} tone="toneGood" sub={`${formatPercent(stats.restored, stats.uniqueFound)} of unique`} />
            <Tile label="Failed" icon="ErrorBadge" value={stats.failed} tone={stats.failed > 0 ? 'toneBad' : undefined} />
            <Tile label="Excluded / not processed" icon="Blocked2" value={stats.excluded + stats.skipped} />
          </>
        ) : (
          <>
            <Tile label="To restore" icon="Sync" value={pending} tone="toneGood" />
            <Tile label="Excluded" icon="Blocked2" value={stats.excluded} />
          </>
        )}
      </div>

      {byKind.length > 0 && (
        <div className={styles.kindBreakdown}>
          {byKind.map((k) => (
            <div key={k.kind} className={styles.kindRow}>
              <span className={styles.kindLabel}>
                <Icon iconName={KIND_ICONS[k.kind]} /> {KIND_LABELS[k.kind]}
              </span>
              <span className={styles.kindBar} aria-hidden="true">
                <span
                  className={styles.kindBarFill}
                  style={{ width: `${Math.max(2, (k.total / Math.max(1, stats.uniqueFound)) * 100)}%` }}
                >
                  {hasRestore && (
                    <span className={styles.kindBarDone} style={{ width: `${(k.restored / Math.max(1, k.total)) * 100}%` }} />
                  )}
                </span>
              </span>
              <span className={styles.kindCount}>
                {hasRestore ? `${k.restored.toLocaleString()} / ${k.total.toLocaleString()} restored` : `${k.total.toLocaleString()} unique`}
              </span>
            </div>
          ))}
        </div>
      )}

      {report.scanErrors.length > 0 && (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline>
          <strong>{report.scanErrors.length.toLocaleString()} part(s) of the scope could not be read</strong> and were not
          checked. They are listed in the CSV as well.
          <ul className={styles.errorList}>
            {(showAllErrors ? report.scanErrors : report.scanErrors.slice(0, ERRORS_SHOWN)).map((e, i) => (
              <li key={i}>{e}</li>
            ))}
          </ul>
          {report.scanErrors.length > ERRORS_SHOWN && (
            <button type="button" className={styles.linkButton} onClick={() => setShowAllErrors(!showAllErrors)}>
              {showAllErrors ? 'Show fewer' : `Show all ${report.scanErrors.length.toLocaleString()}`}
            </button>
          )}
        </MessageBar>
      )}

      <div className={styles.tableCard}>
        <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 10 }} className={styles.filterBar}>
          <SearchBox
            placeholder="Filter by path or principal"
            className={styles.searchBox}
            value={search}
            onChange={(_, v) => setSearch(v || '')}
          />
          <Dropdown
            className={styles.filterDropdown}
            selectedKey={statusFilter}
            options={statusOptions}
            onChange={(_, o) => o && setStatusFilter(String(o.key))}
            ariaLabel="Filter by status"
          />
          <Dropdown
            className={styles.filterDropdown}
            selectedKey={kindFilter}
            options={kindOptions}
            onChange={(_, o) => o && setKindFilter(String(o.key))}
            ariaLabel="Filter by type"
          />
          <span className={styles.spacer} />
          {selectable && onSetExcluded && (
            <>
              <DefaultButton
                text={`Exclude${selectedCount ? ` ${selectedCount.toLocaleString()}` : ''}`}
                iconProps={{ iconName: 'Blocked2' }}
                disabled={selectedCount === 0}
                onClick={() => onSetExcluded(selectedObjects(), true)}
              />
              <DefaultButton
                text="Include again"
                iconProps={{ iconName: 'Undo' }}
                disabled={selectedCount === 0}
                onClick={() => onSetExcluded(selectedObjects(), false)}
              />
            </>
          )}
          <DefaultButton text="Export CSV" iconProps={{ iconName: 'ExcelDocument' }} onClick={() => downloadCsv(report)} />
        </Stack>
        <div className={styles.tableCount}>
          {filtered.length === report.objects.length
            ? `${report.objects.length.toLocaleString()} objects with unique permissions`
            : `${filtered.length.toLocaleString()} of ${report.objects.length.toLocaleString()} objects`}
        </div>
        {report.objects.length === 0 ? (
          <div className={styles.emptyTable}>
            <Icon iconName="CompletedSolid" className={styles.emptyTableIcon} />
            Nothing in this scope has unique permissions. Everything already inherits.
          </div>
        ) : (
          <div className={styles.tableScroll} data-is-scrollable="true">
            <DetailsList
              items={filtered}
              columns={columns}
              compact
              selection={selectionRef.current}
              selectionMode={selectable ? SelectionMode.multiple : SelectionMode.none}
              checkboxVisibility={selectable ? CheckboxVisibility.always : CheckboxVisibility.hidden}
              layoutMode={DetailsListLayoutMode.justified}
              constrainMode={ConstrainMode.unconstrained}
              getKey={(item: IUniqueObject) => item.key}
              setKey="objects"
            />
          </div>
        )}
      </div>
    </div>
  );
};
