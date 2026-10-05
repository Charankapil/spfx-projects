import * as React from 'react';
import {
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  Dropdown,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
  Pivot,
  PivotItem,
  PrimaryButton,
  Selection,
  SelectionMode,
  TextField,
  Toggle
} from '@fluentui/react';
import { parseCsv } from '../../services/csv';
import { formatBytes, formatDateTime } from '../../services/format';
import { IGrowthSettings } from '../../services/GrowthEngine';
import { buildSnapshot, detectMapping, IMapping, Unit } from '../../services/StorageImport';
import { ICsvFile, IImportResult, ITenantIndex } from '../../services/TenantStore';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { ErrorBar, Loading, Pill } from '../shared/ui';

const UNITS: IDropdownOption[] = [
  { key: 'bytes', text: 'Bytes' },
  { key: 'kb', text: 'KB' },
  { key: 'mb', text: 'MB' },
  { key: 'gb', text: 'GB' },
  { key: 'tb', text: 'TB' }
];

interface ILoaded {
  text: string;
  name: string;
  folder?: string;
  /** Epoch seconds the file was last modified (the default snapshot date). */
  modified: number;
}

function dirOf(path: string): string {
  return path.substring(0, path.lastIndexOf('/'));
}

/**
 * Pick the CSV that your flow writes (from a library on this site, or from your
 * computer), confirm the column mapping against a live preview, and import it
 * as a storage snapshot for every site in the file.
 */
export const ImportPanel: React.FC<{ open: boolean; onDismiss: () => void; index?: ITenantIndex; settings: IGrowthSettings; onImported: (r: IImportResult) => void }> = ({ open, onDismiss, index, settings, onImported }) => {
  const ctx = useAdmin();
  const { tenant } = ctx;
  const origin = (/^(https:\/\/[^/]+)/i.exec(ctx.homeWebUrl) || [''])[1];

  const [libs, setLibs] = React.useState<Array<{ title: string; rootFolder: string }>>([]);
  const [lib, setLib] = React.useState<string | undefined>();
  const [folders, setFolders] = React.useState<Array<{ name: string; path: string }>>([]);
  const [folder, setFolder] = React.useState<string | undefined>();
  const [files, setFiles] = React.useState<ICsvFile[]>([]);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>();
  const [loaded, setLoaded] = React.useState<ILoaded | undefined>();
  const [mapping, setMapping] = React.useState<IMapping | undefined>();
  const [dateOverride, setDateOverride] = React.useState('');
  const [step, setStep] = React.useState('');
  const [importing, setImporting] = React.useState(false);

  const rows = React.useMemo(() => (loaded ? parseCsv(loaded.text) : []), [loaded]);
  const headers = rows[0] || [];
  const t = dateOverride ? Math.floor(Date.parse(dateOverride + 'T12:00:00') / 1000) : loaded ? loaded.modified : 0;
  // Re-built only when the file, the mapping or the date changes (17,000 rows take well under a second).
  const preview = React.useMemo(() => (loaded && mapping && mapping.url && mapping.storage ? buildSnapshot(rows, mapping, origin, t || 0) : undefined), [loaded, rows, mapping, origin, t]);

  const fileSelection = React.useMemo(
    () => new Selection({ getKey: (f) => (f as ICsvFile).serverRelativeUrl, onSelectionChanged: () => undefined }),
    []
  );

  const loadFolderFiles = async (path: string): Promise<void> => {
    setBusy(true);
    setError(undefined);
    try {
      setFiles(await tenant.listCsvFiles(path));
    } catch (e) {
      setError((e as Error).message);
      setFiles([]);
    } finally {
      setBusy(false);
    }
  };

  const chooseLibrary = async (rootFolder: string): Promise<void> => {
    setLib(rootFolder);
    setFolder(rootFolder);
    setFolders([]);
    try {
      setFolders(await tenant.listFolders(rootFolder));
    } catch {
      /* subfolders are optional */
    }
    await loadFolderFiles(rootFolder);
  };

  React.useEffect(() => {
    if (!open || libs.length > 0) {
      return;
    }
    let alive = true;
    setBusy(true);
    tenant
      .listLibraries()
      .then((l) => {
        if (!alive) {
          return;
        }
        setLibs(l);
        // Start where the last import came from, else the first library.
        const known = index && index.folder ? l.filter((x) => index.folder && index.folder.indexOf(x.rootFolder) === 0)[0] : undefined;
        const start = known || l[0];
        if (start) {
          chooseLibrary(start.rootFolder).then(() => {
            if (known && index && index.folder && index.folder !== known.rootFolder) {
              setFolder(index.folder);
              loadFolderFiles(index.folder);
            }
          });
        }
      })
      .catch((e) => alive && setError((e as Error).message))
      .then(() => alive && setBusy(false));
    return () => {
      alive = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  const adopt = (l: ILoaded): void => {
    setLoaded(l);
    setError(undefined);
    const r = parseCsv(l.text);
    const saved = index && index.mapping;
    const compatible = saved && r[0] && r[0].indexOf(saved.url) >= 0 && r[0].indexOf(saved.storage) >= 0 ? saved : undefined;
    setMapping(compatible || detectMapping(r[0] || [], r.slice(1)) || { url: '', storage: '', storageUnit: 'mb', quotaUnit: 'mb', excludeOneDrive: true });
  };

  const pickFile = async (f: ICsvFile): Promise<void> => {
    setBusy(true);
    setError(undefined);
    try {
      const text = await tenant.readCsv(f.serverRelativeUrl);
      adopt({ text, name: f.name, folder: dirOf(f.serverRelativeUrl), modified: f.modified ? Math.floor(f.modified.getTime() / 1000) : Math.floor(Date.now() / 1000) });
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const pickLocal = (file: File | undefined): void => {
    if (!file) {
      return;
    }
    const reader = new FileReader();
    reader.onload = () => adopt({ text: String(reader.result || ''), name: file.name, modified: Math.floor((file.lastModified || Date.now()) / 1000) });
    reader.onerror = () => setError('That file could not be read.');
    reader.readAsText(file);
  };

  const run = async (): Promise<void> => {
    if (!loaded || !mapping || !preview || !t) {
      return;
    }
    setImporting(true);
    setError(undefined);
    try {
      const r = await tenant.importCsv(loaded.text, mapping, { t, fileName: loaded.name, folder: loaded.folder, modified: loaded.modified }, settings, setStep);
      ctx.log('Import storage CSV', `${loaded.name}: ${r.stats.kept} sites`, true);
      ctx.notify(`Imported ${r.stats.kept.toLocaleString()} sites from ${loaded.name}.${r.replaced ? ' It replaced an earlier snapshot from the same day.' : ''}`, 'success');
      onImported(r);
      setLoaded(undefined);
      onDismiss();
    } catch (e) {
      ctx.log('Import storage CSV', loaded.name, false, (e as Error).message);
      setError((e as Error).message);
    } finally {
      setImporting(false);
      setStep('');
    }
  };

  const colOptions = (optional: boolean): IDropdownOption[] => (optional ? [{ key: '', text: '(none)' }] : []).concat(headers.map((h) => ({ key: h, text: h })));
  const set = (patch: Partial<IMapping>): void => setMapping(mapping ? { ...mapping, ...patch } : undefined);
  const sample = preview ? preview.data.u.slice(0, 5).map((u, i) => ({ u, b: preview.data.b[i], q: preview.data.q[i] })) : [];
  const total = preview ? preview.data.b.reduce((s, x) => s + x, 0) : 0;
  const biggest = preview && preview.data.b.length ? preview.data.b.reduce((m, x, i) => (x > preview.data.b[m] ? i : m), 0) : -1;

  return (
    <Panel isOpen={open} onDismiss={onDismiss} type={PanelType.large} headerText="Import storage CSV" closeButtonAriaLabel="Close" isLightDismiss={!loaded}>
      <p className={styles.muted}>
        Pick the CSV your flow saves from the SharePoint admin center site list. It needs one column with the site address and one with the storage used. Everything is read in your browser and saved as a snapshot in Site Assets.
      </p>

      {!loaded && (
        <Pivot aria-label="CSV source">
          <PivotItem headerText="From a library on this site">
            <div style={{ paddingTop: 12 }}>
              <div className={styles.filters}>
                <Dropdown className={styles.field} label="Library" selectedKey={lib} options={libs.map((l) => ({ key: l.rootFolder, text: l.title }))} onChange={(_, o) => o && chooseLibrary(String(o.key))} />
                <Dropdown
                  className={styles.field}
                  label="Folder"
                  selectedKey={folder}
                  options={(lib ? [{ key: lib, text: '(top level)' }] : []).concat(folders.map((f) => ({ key: f.path, text: f.name })))}
                  onChange={(_, o) => {
                    if (o) {
                      setFolder(String(o.key));
                      loadFolderFiles(String(o.key));
                    }
                  }}
                />
                <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={() => folder && loadFolderFiles(folder)} disabled={busy || !folder}>
                  Refresh
                </DefaultButton>
              </div>
              {busy && <Loading text="Working…" />}
              {!busy && files.length === 0 && folder && <div className={styles.empty}>No CSV files in this folder.</div>}
              {files.length > 0 && (
                <div className={styles.tableWrap} style={{ marginTop: 10 }}>
                  <DetailsList
                    items={files}
                    selection={fileSelection}
                    selectionMode={SelectionMode.none}
                    layoutMode={DetailsListLayoutMode.justified}
                    columns={[
                      { key: 'n', name: 'File', minWidth: 200, isResizable: true, onRender: (f: ICsvFile) => <strong>{f.name}</strong> },
                      { key: 'm', name: 'Modified', minWidth: 140, onRender: (f: ICsvFile) => formatDateTime(f.modified) },
                      { key: 's', name: 'Size', minWidth: 70, onRender: (f: ICsvFile) => formatBytes(f.size) },
                      { key: 'a', name: '', minWidth: 90, onRender: (f: ICsvFile) => <PrimaryButton text="Use this file" onClick={() => pickFile(f)} disabled={busy} /> }
                    ]}
                  />
                </div>
              )}
            </div>
          </PivotItem>
          <PivotItem headerText="From my computer">
            <div style={{ paddingTop: 12 }}>
              <input type="file" accept=".csv,text/csv" aria-label="Choose a CSV file" onChange={(e) => pickLocal(e.target.files ? e.target.files[0] : undefined)} />
            </div>
          </PivotItem>
        </Pivot>
      )}

      {error && <ErrorBar error={error} />}

      {loaded && mapping && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          <MessageBar messageBarType={MessageBarType.info}>
            <strong>{loaded.name}</strong> · {Math.max(0, rows.length - 1).toLocaleString()} rows · snapshot dated {formatDateTime(new Date(t * 1000))}
            <button type="button" className={styles.link} style={{ marginLeft: 12 }} onClick={() => setLoaded(undefined)}>
              Choose a different file
            </button>
          </MessageBar>

          <div className={styles.filters}>
            <Dropdown className={styles.field} label="Site address column" required selectedKey={mapping.url} options={colOptions(false)} onChange={(_, o) => o && set({ url: String(o.key) })} />
            <Dropdown className={styles.field} label="Storage used column" required selectedKey={mapping.storage} options={colOptions(false)} onChange={(_, o) => o && set({ storage: String(o.key) })} />
            <Dropdown className={styles.field} label="Storage unit" selectedKey={mapping.storageUnit} options={UNITS} onChange={(_, o) => o && set({ storageUnit: String(o.key) as Unit })} />
          </div>
          <div className={styles.filters}>
            <Dropdown className={styles.field} label="Storage quota column (optional)" selectedKey={mapping.quota || ''} options={colOptions(true)} onChange={(_, o) => o && set({ quota: String(o.key) || undefined })} />
            <Dropdown className={styles.field} label="Quota unit" selectedKey={mapping.quotaUnit} options={UNITS} onChange={(_, o) => o && set({ quotaUnit: String(o.key) as Unit })} disabled={!mapping.quota} />
            <Dropdown className={styles.field} label="Site title column (optional)" selectedKey={mapping.title || ''} options={colOptions(true)} onChange={(_, o) => o && set({ title: String(o.key) || undefined })} />
            <Dropdown className={styles.field} label="Deleted marker column (optional)" selectedKey={mapping.deleted || ''} options={colOptions(true)} onChange={(_, o) => o && set({ deleted: String(o.key) || undefined })} />
            <Dropdown className={styles.field} label="Archive status column (optional)" selectedKey={mapping.archived || ''} options={colOptions(true)} onChange={(_, o) => o && set({ archived: String(o.key) || undefined })} />
          </div>
          <div className={styles.filters}>
            <Dropdown className={styles.field} label="Template column (optional)" selectedKey={mapping.template || ''} options={colOptions(true)} onChange={(_, o) => o && set({ template: String(o.key) || undefined })} />
            <Dropdown className={styles.field} label="Teams-connected column (optional)" selectedKey={mapping.teams || ''} options={colOptions(true)} onChange={(_, o) => o && set({ teams: String(o.key) || undefined })} />
            <Dropdown className={styles.field} label="Last activity column (optional)" selectedKey={mapping.lastActivity || ''} options={colOptions(true)} onChange={(_, o) => o && set({ lastActivity: String(o.key) || undefined })} />
          </div>
          <div className={styles.filters}>
            <Toggle label="Skip OneDrive sites" inlineLabel checked={mapping.excludeOneDrive} onChange={(_, c) => set({ excludeOneDrive: !!c })} />
            <Toggle label="Keep deleted sites as their own group" inlineLabel checked={mapping.includeDeleted !== false} onChange={(_, c) => set({ includeDeleted: !!c })} disabled={!mapping.deleted} />
            <TextField className={styles.field} label="Snapshot date (leave empty to use the file's modified date)" type="date" value={dateOverride} onChange={(_, v) => setDateOverride(v || '')} />
          </div>

          {!preview || preview.stats.kept === 0 ? (
            <MessageBar messageBarType={MessageBarType.warning}>No usable site rows yet. Choose the column that holds the site address and the column that holds storage used.</MessageBar>
          ) : (
            <div className={styles.card}>
              <h3 className={styles.cardTitle}>Check this looks right</h3>
              <p style={{ marginTop: 0 }}>
                <strong>{preview.stats.kept.toLocaleString()}</strong> sites · <strong>{formatBytes(total)}</strong> in total
                {biggest >= 0 && (
                  <>
                    {' '}
                    · largest: {preview.data.u[biggest]} ({formatBytes(preview.data.b[biggest])})
                  </>
                )}
              </p>
              <p style={{ margin: '0 0 6px' }}>
                <Pill kind="good">{(preview.stats.kept - preview.stats.archived - (mapping.includeDeleted !== false ? preview.stats.deleted : 0)).toLocaleString()} active</Pill>{' '}
                <Pill kind="info">{preview.stats.archived.toLocaleString()} archived</Pill>{' '}
                <Pill kind="warning">
                  {preview.stats.deleted.toLocaleString()} deleted{mapping.includeDeleted !== false ? '' : ' (skipped)'}
                </Pill>
                {!mapping.archived && <span className={styles.muted}> · choose an archive status column to separate archived sites</span>}
              </p>
              <p className={styles.muted}>
                Skipped: {preview.stats.noUrl} without an address, {preview.stats.badNumber} without a valid number, {preview.stats.oneDrive} OneDrive, {preview.stats.duplicates} duplicates.
                If the total or the largest site looks wrong by a factor of 1,000 or 1,024, change the unit.
              </p>
              <DetailsList
                items={sample}
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.justified}
                columns={[
                  { key: 'u', name: 'Site', minWidth: 220, isResizable: true, onRender: (r: { u: string }) => r.u },
                  { key: 'b', name: 'Storage', minWidth: 90, onRender: (r: { b: number }) => formatBytes(r.b) },
                  { key: 'q', name: 'Of quota', minWidth: 80, onRender: (r: { q: number }) => (r.q ? `${(r.q / 10).toFixed(1)}%` : '-') }
                ]}
              />
            </div>
          )}

          {importing && <Loading text={step || 'Importing…'} />}
          <div className={styles.actions}>
            <PrimaryButton onClick={run} disabled={importing || !preview || preview.stats.kept === 0 || !t}>
              Import {preview ? preview.stats.kept.toLocaleString() : ''} sites
            </PrimaryButton>
            <DefaultButton onClick={onDismiss} disabled={importing}>
              Cancel
            </DefaultButton>
          </div>
          <p className={styles.muted} style={{ margin: 0 }}>
            Growth is calculated between snapshot dates, so import regularly (daily or weekly) and keep the date accurate. A new import within 6 hours of an earlier one replaces it.
          </p>
        </div>
      )}
    </Panel>
  );
};
