import * as React from 'react';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';

import { IExplorerData, ILoadProgress, IMetadataField, LOCATION_KEY } from '../models/IMetadata';
import { IMetadataSource } from '../services/IMetadataSource';
import { coverage, filterDocuments, FilterState, groupDocuments, toCsv } from '../services/grouping';
import { CoverageTable } from './CoverageTable';
import { FacetPanel } from './FacetPanel';
import { GroupList } from './GroupList';
import styles from './MetadataExplorer.module.scss';

export interface IMetadataExplorerProps {
  title: string;
  source: IMetadataSource;
  /** Field title or key to group by on first load. */
  defaultGroupBy: string;
  defaultThenBy: string;
}

type View = 'groups' | 'coverage';

const NONE = '__none';

function findField(fields: IMetadataField[], wanted: string): IMetadataField | undefined {
  const w = (wanted || '').trim().toLowerCase();
  return w ? fields.filter((f) => f.key.toLowerCase() === w || f.title.toLowerCase() === w)[0] : undefined;
}

/** Without a configured default, group by the first real column (the ECM "Document Type" in a seeded site). */
function initialGroupBy(fields: IMetadataField[], wanted: string): string {
  const configured = findField(fields, wanted);
  if (configured) {
    return configured.key;
  }
  const firstColumn = fields.filter((f) => f.kind !== 'builtin' && f.key !== 'ContentType' && f.key !== '_ComplianceTag')[0];
  return (firstColumn || findField(fields, 'ContentType') || fields[0]).key;
}

function downloadCsv(csv: string, fileName: string): void {
  const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' });
  const nav = window.navigator as Navigator & { msSaveOrOpenBlob?: (b: Blob, n: string) => void };
  if (nav.msSaveOrOpenBlob) {
    nav.msSaveOrOpenBlob(blob, fileName);
    return;
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export const MetadataExplorer = (props: IMetadataExplorerProps): JSX.Element => {
  const [data, setData] = React.useState<IExplorerData | undefined>(undefined);
  const [progress, setProgress] = React.useState<ILoadProgress | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [groupBy, setGroupBy] = React.useState<string>('');
  const [thenBy, setThenBy] = React.useState<string>(NONE);
  const [filters, setFilters] = React.useState<FilterState>({});
  const [text, setText] = React.useState('');
  const [view, setView] = React.useState<View>('groups');
  const [reloadToken, setReloadToken] = React.useState(0);
  const loadedSource = React.useRef<IMetadataSource | undefined>(undefined);

  React.useEffect(() => {
    let cancelled = false;
    if (loadedSource.current !== props.source) {
      loadedSource.current = props.source;
      setFilters({});
      setText('');
    }
    setError(undefined);
    setProgress({ message: 'Starting', loaded: 0 });
    props.source
      .load((p) => !cancelled && setProgress(p))
      .then((result) => {
        if (cancelled) {
          return;
        }
        setData(result);
        setGroupBy((current) => (findField(result.fields, current) ? current : initialGroupBy(result.fields, props.defaultGroupBy)));
        const then = findField(result.fields, props.defaultThenBy);
        setThenBy((current) => (current !== NONE && findField(result.fields, current) ? current : then ? then.key : NONE));
        setProgress(undefined);
      })
      .catch((err: Error) => {
        if (!cancelled) {
          setError(err.message);
          setProgress(undefined);
        }
      });
    return () => {
      cancelled = true;
    };
  }, [props.source, reloadToken]);

  const fields = data?.fields || [];
  const documents = data?.documents || [];
  const filtered = React.useMemo(() => filterDocuments(documents, filters, text), [documents, filters, text]);
  const groupField = findField(fields, groupBy);
  const thenField = thenBy !== NONE ? findField(fields, thenBy) : undefined;
  const groups = React.useMemo(
    () => (groupField ? groupDocuments(filtered, groupField, thenField) : []),
    [filtered, groupField, thenField]
  );
  const coverageRows = React.useMemo(
    () => coverage(filtered, fields).sort((a, b) => a.percent - b.percent || a.field.title.localeCompare(b.field.title)),
    [filtered, fields]
  );

  // Each facet counts against every filter except its own, so ticking more values in one field stays possible.
  const documentsFor = React.useCallback(
    (fieldKey: string) => {
      if (!filters[fieldKey] || !filters[fieldKey].length) {
        return filtered;
      }
      const others: FilterState = { ...filters, [fieldKey]: [] };
      return filterDocuments(documents, others, text);
    },
    [documents, filters, filtered, text]
  );

  const toggleFilter = (fieldKey: string, value: string): void => {
    const current = filters[fieldKey] || [];
    const next = current.indexOf(value) >= 0 ? current.filter((v) => v !== value) : current.concat([value]);
    setFilters({ ...filters, [fieldKey]: next });
  };

  const activeFilters: { field: IMetadataField; value: string }[] = [];
  Object.keys(filters).forEach((key) => {
    const field = findField(fields, key);
    if (field) {
      filters[key].forEach((value) => activeFilters.push({ field, value }));
    }
  });

  const fieldOptions: IDropdownOption[] = fields.map((f) => ({ key: f.key, text: f.title }));
  const chipFields = fields.filter(
    (f) => f.key !== groupField?.key && f.key !== thenField?.key && f.kind !== 'builtin'
  );
  const locationTitle = findField(fields, LOCATION_KEY)?.title || 'Location';

  return (
    <div className={styles.explorer}>
      <header className={styles.header}>
        <div className={styles.headerText}>
          <h2 className={styles.title}>{props.title}</h2>
          {data && (
            <p className={styles.summary}>
              {data.scopeLabel} · {documents.length.toLocaleString()} documents · {fields.length} metadata categories
            </p>
          )}
        </div>
        <div className={styles.headerActions}>
          <DefaultButton
            text="Export CSV"
            iconProps={{ iconName: 'Download' }}
            disabled={!filtered.length}
            onClick={() => downloadCsv(toCsv(filtered, fields), 'metadata-explorer.csv')}
          />
          <DefaultButton text="Refresh" iconProps={{ iconName: 'Refresh' }} disabled={!!progress} onClick={() => setReloadToken(reloadToken + 1)} />
        </div>
      </header>

      {error && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          Couldn&apos;t load documents: {error}
        </MessageBar>
      )}

      {progress && (
        <div className={styles.loading}>
          <Spinner size={SpinnerSize.medium} />
          <span>
            {progress.message}… {progress.loaded ? `${progress.loaded.toLocaleString()} documents so far` : ''}
          </span>
        </div>
      )}

      {data && data.truncated && (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline>
          Showing the first {documents.length.toLocaleString()} documents. Raise &ldquo;Maximum documents&rdquo; in the web part settings, or narrow
          the libraries or query, to see everything.
        </MessageBar>
      )}

      {data && data.sources.some((s) => s.error) && (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline>
          Some libraries could not be read: {data.sources.filter((s) => s.error).map((s) => s.title).join(', ')}.
        </MessageBar>
      )}

      {data && !progress && (
        <>
          <div className={styles.toolbar}>
            <SearchBox
              className={styles.searchBox}
              placeholder="Filter by name or metadata value"
              value={text}
              onChange={(_, v) => setText(v || '')}
            />
            <Dropdown
              className={styles.dropdown}
              label="Group by"
              selectedKey={groupField?.key}
              options={fieldOptions}
              onChange={(_, o) => o && setGroupBy(String(o.key))}
            />
            <Dropdown
              className={styles.dropdown}
              label="Then by"
              selectedKey={thenField ? thenField.key : NONE}
              options={([{ key: NONE, text: 'None' }] as IDropdownOption[]).concat(fieldOptions.filter((o) => o.key !== groupField?.key))}
              onChange={(_, o) => o && setThenBy(String(o.key))}
            />
            <div className={styles.viewSwitch} role="group" aria-label="View">
              <button type="button" aria-pressed={view === 'groups'} onClick={() => setView('groups')}>
                Grouped documents
              </button>
              <button type="button" aria-pressed={view === 'coverage'} onClick={() => setView('coverage')}>
                Metadata coverage
              </button>
            </div>
          </div>

          {activeFilters.length > 0 && (
            <div className={styles.activeFilters}>
              {activeFilters.map(({ field, value }) => (
                <button
                  type="button"
                  key={`${field.key}:${value}`}
                  className={styles.activeFilter}
                  onClick={() => toggleFilter(field.key, value)}
                  aria-label={`Remove filter ${field.title}: ${value}`}
                >
                  {field.title}: {value} <span aria-hidden="true">×</span>
                </button>
              ))}
              <button type="button" className={styles.linkButton} onClick={() => setFilters({})}>
                Clear all
              </button>
            </div>
          )}

          <div className={styles.body}>
            <FacetPanel fields={fields} documentsFor={documentsFor} filters={filters} onToggle={toggleFilter} />
            <main className={styles.results}>
              <div className={styles.resultCount}>
                {filtered.length.toLocaleString()} of {documents.length.toLocaleString()} documents
                {view === 'groups' && groupField ? ` in ${groups.length} ${groupField.title} groups` : ''}
              </div>
              {!filtered.length && <div className={styles.empty}>No documents match these filters.</div>}
              {view === 'groups' && groupField && filtered.length > 0 && (
                <GroupList
                  groups={groups}
                  field={groupField}
                  chipFields={thenField ? [thenField].concat(chipFields) : chipFields}
                  locationTitle={locationTitle}
                />
              )}
              {view === 'coverage' && filtered.length > 0 && (
                <CoverageTable
                  rows={coverageRows}
                  total={filtered.length}
                  onGroupBy={(key) => {
                    setGroupBy(key);
                    setView('groups');
                  }}
                />
              )}
            </main>
          </div>
        </>
      )}
    </div>
  );
};
