import * as React from 'react';
import { useCallback, useEffect, useRef, useState } from 'react';
import {
  DefaultButton,
  Dropdown,
  Icon,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
  Pivot,
  PivotItem,
  PrimaryButton,
  ProgressIndicator,
  Spinner,
  SpinnerSize
} from '@fluentui/react';

import { downloadCsv } from '../../common/csv';
import { extensionOf, fileIcon, formatBytes, formatDate, formatNumber } from '../../common/format';
import { CancelToken, errorText, isCancelled } from '../../common/RequestQueue';
import { fileNameOf, folderOf, search, underPath } from '../../common/search';
import { IToolkitComponentProps } from '../../common/ToolkitWebPart';
import { Empty, Shell } from '../../common/ui/Shell';
import styles from '../../common/ui/Toolkit.module.scss';
import { useMounted, useToolkit } from '../../common/ui/useToolkit';

export interface IStorageHogFinderProps {
  title: string;
}

interface IStorageMetrics {
  TotalSize?: string | number;
  TotalFileStreamSize?: string | number;
  TotalFileCount?: string | number;
  LastModified?: string;
}

interface IFolderRow {
  Name: string;
  ServerRelativeUrl: string;
  StorageMetrics?: IStorageMetrics;
}

interface IFolderEntry {
  name: string;
  path: string;
  totalSize: number;
  currentSize: number;
  fileCount: number;
  lastModified: string;
}

interface IFileRow {
  Name: string;
  ServerRelativeUrl: string;
  Length: string | number;
  TimeLastModified: string;
  UIVersionLabel?: string;
}

interface ILargeFile {
  name: string;
  path: string;
  folder: string;
  size: number;
  modified: string;
  modifiedBy: string;
}

interface IVersionRow {
  ID: number;
  VersionLabel: string;
  Size: string | number;
  Created: string;
}

interface ICrumb {
  name: string;
  path: string;
}

const METRICS_QUERY =
  '$select=Name,ServerRelativeUrl,StorageMetrics/TotalSize,StorageMetrics/TotalFileStreamSize,' +
  'StorageMetrics/TotalFileCount,StorageMetrics/LastModified&$expand=StorageMetrics';

function toEntry(row: IFolderRow, displayName?: string): IFolderEntry {
  const metrics = row.StorageMetrics || {};
  const totalSize = Number(metrics.TotalSize) || 0;
  return {
    name: displayName || row.Name,
    path: row.ServerRelativeUrl,
    totalSize,
    // Never report more "current" than "total", whatever the server rounds to.
    currentSize: Math.min(Number(metrics.TotalFileStreamSize) || 0, totalSize),
    fileCount: Number(metrics.TotalFileCount) || 0,
    lastModified: metrics.LastModified || ''
  };
}

const KEEP_OPTIONS = [5, 10, 25, 50].map((n) => ({ key: n, text: `Keep the newest ${n} versions` }));

export const StorageHogFinder: React.FC<IToolkitComponentProps<IStorageHogFinderProps>> = (props) => {
  const kit = useToolkit(props.context);
  const mounted = useMounted();

  const [crumbs, setCrumbs] = useState<ICrumb[]>([]);
  const [folders, setFolders] = useState<IFolderEntry[] | undefined>(undefined);
  const [files, setFiles] = useState<IFileRow[] | undefined>(undefined);
  const [filesNote, setFilesNote] = useState<string | undefined>(undefined);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);
  const [notice, setNotice] = useState<string | undefined>(undefined);

  const [largest, setLargest] = useState<ILargeFile[] | undefined>(undefined);
  const [largestLoading, setLargestLoading] = useState(false);

  const [versionFile, setVersionFile] = useState<{ name: string; path: string } | undefined>(undefined);
  const [versions, setVersions] = useState<IVersionRow[] | undefined>(undefined);
  const [keep, setKeep] = useState<number>(10);
  const [confirmTrim, setConfirmTrim] = useState(false);
  const [trimProgress, setTrimProgress] = useState<{ done: number; total: number } | undefined>(undefined);
  const trimToken = useRef<CancelToken | undefined>(undefined);

  // Guards against an older, slower response overwriting a newer one.
  const loadSeq = useRef(0);

  const loadFolder = useCallback(
    (trail: ICrumb[]) => {
      const seq = ++loadSeq.current;
      setLoading(true);
      setError(undefined);
      setFiles(undefined);
      setFilesNote(undefined);

      const run = async (): Promise<IFolderEntry[]> => {
        if (trail.length === 0) {
          // Two requests cover every library: the list of libraries, and the
          // storage metrics of the web's top-level folders.
          const libraries = await kit.sp.getJson<{
            value: Array<{ Title: string; RootFolder: { ServerRelativeUrl: string } }>;
          }>(
            `${kit.sp.webUrl}/_api/web/lists?$filter=BaseTemplate eq 101 and Hidden eq false` +
              `&$select=Title,RootFolder/ServerRelativeUrl&$expand=RootFolder`
          );
          const roots = await kit.sp.getJson<{ value: IFolderRow[] }>(
            `${kit.sp.webUrl}/_api/web/RootFolder/Folders?${METRICS_QUERY}`
          );
          const byPath: Record<string, IFolderRow> = {};
          for (const row of roots.value || []) {
            byPath[row.ServerRelativeUrl.toLowerCase()] = row;
          }
          const entries: IFolderEntry[] = [];
          for (const library of libraries.value || []) {
            const row = byPath[library.RootFolder.ServerRelativeUrl.toLowerCase()];
            if (row) {
              entries.push(toEntry(row, library.Title));
            }
          }
          return entries;
        }
        const current = trail[trail.length - 1];
        const children = await kit.sp.getJson<{ value: IFolderRow[] }>(
          `${kit.sp.folderApi(current.path)}/Folders?${METRICS_QUERY}`
        );
        return (children.value || []).filter((row) => row.Name !== 'Forms' && row.Name.charAt(0) !== '_').map((row) => toEntry(row));
      };

      run()
        .then((entries) => {
          if (!mounted.current || seq !== loadSeq.current) {
            return;
          }
          entries.sort((a, b) => b.totalSize - a.totalSize);
          setFolders(entries);
          setCrumbs(trail);
          setLoading(false);
        })
        .catch((err) => {
          if (!mounted.current || seq !== loadSeq.current) {
            return;
          }
          setLoading(false);
          setError(
            `Storage figures could not be read: ${errorText(err)} ` +
              `Storage figures are only available to people who can manage this site.`
          );
        });
    },
    [kit, mounted]
  );

  useEffect(() => {
    loadFolder([]);
  }, [loadFolder]);

  const loadFiles = useCallback(() => {
    const current = crumbs[crumbs.length - 1];
    if (!current) {
      return;
    }
    const seq = loadSeq.current;
    setLoading(true);
    kit.sp
      .getJson<{ value: IFileRow[] }>(
        `${kit.sp.folderApi(current.path)}/Files?$select=Name,ServerRelativeUrl,Length,TimeLastModified,UIVersionLabel&$top=500`
      )
      .then((json) => {
        if (!mounted.current || seq !== loadSeq.current) {
          return;
        }
        const rows = (json.value || []).slice().sort((a, b) => Number(b.Length) - Number(a.Length));
        setFiles(rows.slice(0, 100));
        setFilesNote(
          rows.length >= 500
            ? 'This folder holds 500 or more files; the largest of the first 500 are shown. Use "Largest files" for the full picture.'
            : undefined
        );
        setLoading(false);
      })
      .catch((err) => {
        if (!mounted.current || seq !== loadSeq.current) {
          return;
        }
        setLoading(false);
        setFilesNote(`The files in this folder could not be listed (${errorText(err)}). Use "Largest files" instead.`);
        setFiles([]);
      });
  }, [crumbs, kit, mounted]);

  const loadLargest = useCallback(() => {
    setLargestLoading(true);
    setError(undefined);
    search(kit.sp, {
      kql: `IsDocument:1 ${underPath(kit.sp.webUrl)}`,
      select: ['Path', 'Size', 'LastModifiedTime', 'ModifiedBy'],
      rowLimit: 50,
      sort: 'Size:descending'
    })
      .then((result) => {
        if (!mounted.current) {
          return;
        }
        setLargest(
          result.rows.map((row) => ({
            name: fileNameOf(row.Path),
            path: row.Path,
            folder: kit.sp.toServerRelative(folderOf(row.Path)),
            size: Number(row.Size) || 0,
            modified: row.LastModifiedTime,
            modifiedBy: row.ModifiedBy
          }))
        );
        setLargestLoading(false);
      })
      .catch((err) => {
        if (!mounted.current) {
          return;
        }
        setLargestLoading(false);
        setError(`The largest files could not be found: ${errorText(err)}`);
      });
  }, [kit, mounted]);

  const onTabChange = useCallback(
    (item?: PivotItem) => {
      if (item && item.props.itemKey === 'largest' && !largest && !largestLoading) {
        loadLargest();
      }
    },
    [largest, largestLoading, loadLargest]
  );

  const openVersions = useCallback(
    (name: string, serverRelativePath: string) => {
      setVersionFile({ name, path: serverRelativePath });
      setVersions(undefined);
      setTrimProgress(undefined);
      kit.sp
        .getJson<{ value: IVersionRow[] }>(`${kit.sp.fileApi(serverRelativePath)}/Versions?$select=ID,VersionLabel,Size,Created`)
        .then((json) => {
          if (mounted.current) {
            // Newest first.
            setVersions((json.value || []).slice().sort((a, b) => b.ID - a.ID));
          }
        })
        .catch((err) => {
          if (mounted.current) {
            setVersions([]);
            setError(`The version history could not be read: ${errorText(err)}`);
          }
        });
    },
    [kit, mounted]
  );

  const closeVersions = useCallback(() => {
    if (trimToken.current) {
      trimToken.current.cancel();
    }
    setVersionFile(undefined);
    setConfirmTrim(false);
  }, []);

  const trim = useCallback(() => {
    setConfirmTrim(false);
    if (!versionFile || !versions) {
      return;
    }
    const doomed = versions.slice(keep);
    if (doomed.length === 0) {
      return;
    }
    const token = new CancelToken();
    trimToken.current = token;
    setTrimProgress({ done: 0, total: doomed.length });

    const run = async (): Promise<number> => {
      let done = 0;
      // One at a time through the page queue: deliberately slow, never a burst.
      for (const version of doomed) {
        await kit.sp.post(`${kit.sp.fileApi(versionFile.path)}/Versions/RecycleByID(vid=${version.ID})`, undefined, { token });
        done++;
        if (mounted.current) {
          setTrimProgress({ done, total: doomed.length });
          setVersions((current) => (current ? current.filter((v) => v.ID !== version.ID) : current));
        }
      }
      return done;
    };

    run()
      .then((done) => {
        if (mounted.current) {
          setTrimProgress(undefined);
          setNotice(`${done} older version${done === 1 ? '' : 's'} of "${versionFile.name}" moved to the recycle bin.`);
        }
      })
      .catch((err) => {
        if (mounted.current) {
          setTrimProgress(undefined);
          if (!isCancelled(err)) {
            setError(`Trimming stopped: ${errorText(err)}`);
          }
        }
      });
  }, [keep, kit, mounted, versionFile, versions]);

  const exportFolders = useCallback(() => {
    downloadCsv(
      'storage',
      ['Name', 'Path', 'Total bytes', 'Current files bytes', 'Older versions and metadata bytes', 'Files', 'Last modified'],
      (folders || []).map((f) => [f.name, f.path, f.totalSize, f.currentSize, f.totalSize - f.currentSize, f.fileCount, f.lastModified])
    );
  }, [folders]);

  const total = (folders || []).reduce((sum, f) => sum + f.totalSize, 0);
  const totalCurrent = (folders || []).reduce((sum, f) => sum + f.currentSize, 0);
  const biggest = folders && folders.length > 0 ? folders[0].totalSize : 0;
  const oldVersions = versions ? versions.slice(keep) : [];
  const oldVersionBytes = oldVersions.reduce((sum, v) => sum + (Number(v.Size) || 0), 0);

  return (
    <Shell
      icon="Database"
      title={props.properties.title || 'Storage Hog Finder'}
      subtitle="See which libraries, folders and version histories use the most storage."
      error={error}
      onDismissError={() => setError(undefined)}
      notice={notice}
      onDismissNotice={() => setNotice(undefined)}
      actions={
        <>
          <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Refresh" disabled={loading} onClick={() => loadFolder(crumbs)} />
          <DefaultButton iconProps={{ iconName: 'Download' }} text="Export CSV" disabled={!folders || folders.length === 0} onClick={exportFolders} />
        </>
      }
    >
      <Pivot onLinkClick={onTabChange}>
        <PivotItem headerText="By folder" itemKey="folders">
          <div className={styles.body} style={{ paddingTop: 12 }}>
            <nav className={styles.breadcrumb} aria-label="Folder path">
              <button type="button" className={styles.linkButton} onClick={() => loadFolder([])}>
                All libraries
              </button>
              {crumbs.map((crumb, index) => (
                <React.Fragment key={crumb.path}>
                  <Icon iconName="ChevronRight" aria-hidden="true" style={{ fontSize: 10 }} />
                  <button type="button" className={styles.linkButton} onClick={() => loadFolder(crumbs.slice(0, index + 1))}>
                    {crumb.name}
                  </button>
                </React.Fragment>
              ))}
            </nav>

            {folders && folders.length > 0 && (
              <div className={styles.kpis}>
                <div className={styles.kpi}>
                  <div className={styles.kpiValue}>{formatBytes(total)}</div>
                  <div className={styles.kpiLabel}>Total shown</div>
                </div>
                <div className={styles.kpi}>
                  <div className={styles.kpiValue}>{formatBytes(totalCurrent)}</div>
                  <div className={styles.kpiLabel}>Current files</div>
                </div>
                <div className={styles.kpi}>
                  <div className={styles.kpiValue}>{formatBytes(total - totalCurrent)}</div>
                  <div className={styles.kpiLabel}>Older versions and metadata</div>
                </div>
              </div>
            )}

            {loading && <Spinner size={SpinnerSize.medium} label="Reading storage figures..." />}

            {!loading && folders && folders.length === 0 && !files && (
              <Empty icon="FabricFolder" text={crumbs.length === 0 ? 'No document libraries were found on this site.' : 'This folder has no subfolders.'} />
            )}

            {folders && folders.length > 0 && (
              <div className={styles.tableWrap}>
                <table className={styles.table}>
                  <thead>
                    <tr>
                      <th>{crumbs.length === 0 ? 'Library' : 'Folder'}</th>
                      <th className={styles.num}>Total</th>
                      <th>Share</th>
                      <th className={styles.num}>Older versions</th>
                      <th className={styles.num}>Files</th>
                      <th>Last modified</th>
                    </tr>
                  </thead>
                  <tbody>
                    {folders.map((folder) => (
                      <tr key={folder.path}>
                        <td>
                          <div className={styles.fileCell}>
                            <Icon iconName="FabricFolder" aria-hidden="true" />
                            <button
                              type="button"
                              className={styles.linkButton}
                              onClick={() => loadFolder(crumbs.concat([{ name: folder.name, path: folder.path }]))}
                            >
                              {folder.name}
                            </button>
                          </div>
                        </td>
                        <td className={styles.num}>{formatBytes(folder.totalSize)}</td>
                        <td>
                          <div className={styles.bar} title={`${formatBytes(folder.currentSize)} current, ${formatBytes(folder.totalSize - folder.currentSize)} older versions and metadata`}>
                            <div style={{ display: 'flex', height: '100%', width: `${biggest > 0 ? (folder.totalSize / biggest) * 100 : 0}%` }}>
                              <div className={styles.barFill} style={{ width: `${folder.totalSize > 0 ? (folder.currentSize / folder.totalSize) * 100 : 0}%`, borderRadius: 0 }} />
                              <div className={styles.barFillAlt} style={{ flex: '1 1 auto' }} />
                            </div>
                          </div>
                        </td>
                        <td className={styles.num}>{formatBytes(folder.totalSize - folder.currentSize)}</td>
                        <td className={styles.num}>{formatNumber(folder.fileCount)}</td>
                        <td className={styles.nowrap}>{formatDate(folder.lastModified)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {crumbs.length > 0 && !files && !loading && (
              <div>
                <DefaultButton iconProps={{ iconName: 'Page' }} text="Show the files in this folder" onClick={loadFiles} />
              </div>
            )}
            {filesNote && <div className={styles.muted}>{filesNote}</div>}
            {files && files.length > 0 && (
              <div className={styles.tableWrap}>
                <table className={styles.table}>
                  <thead>
                    <tr>
                      <th>File</th>
                      <th className={styles.num}>Current size</th>
                      <th>Version</th>
                      <th>Modified</th>
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {files.map((file) => (
                      <tr key={file.ServerRelativeUrl}>
                        <td>
                          <div className={styles.fileCell}>
                            <Icon iconName={fileIcon(extensionOf(file.Name))} aria-hidden="true" />
                            <span className={styles.fileName}>{file.Name}</span>
                          </div>
                        </td>
                        <td className={styles.num}>{formatBytes(Number(file.Length))}</td>
                        <td>{file.UIVersionLabel || ''}</td>
                        <td className={styles.nowrap}>{formatDate(file.TimeLastModified)}</td>
                        <td>
                          <button type="button" className={styles.linkButton} onClick={() => openVersions(file.Name, file.ServerRelativeUrl)}>
                            Versions
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </PivotItem>

        <PivotItem headerText="Largest files" itemKey="largest">
          <div className={styles.body} style={{ paddingTop: 12 }}>
            {largestLoading && <Spinner size={SpinnerSize.medium} label="Finding the largest files..." />}
            {largest && largest.length === 0 && <Empty icon="Page" text="No files were found." />}
            {largest && largest.length > 0 && (
              <div className={styles.tableWrap}>
                <table className={styles.table}>
                  <thead>
                    <tr>
                      <th>File</th>
                      <th className={styles.num}>Current size</th>
                      <th>Modified</th>
                      <th>Modified by</th>
                      <th />
                    </tr>
                  </thead>
                  <tbody>
                    {largest.map((file) => (
                      <tr key={file.path}>
                        <td>
                          <div className={styles.fileCell}>
                            <Icon iconName={fileIcon(extensionOf(file.name))} aria-hidden="true" />
                            <div>
                              <a className={styles.fileName} href={file.path} target="_blank" rel="noopener noreferrer">
                                {file.name}
                              </a>
                              <div className={styles.path}>{file.folder}</div>
                            </div>
                          </div>
                        </td>
                        <td className={styles.num}>{formatBytes(file.size)}</td>
                        <td className={styles.nowrap}>{formatDate(file.modified)}</td>
                        <td>{file.modifiedBy}</td>
                        <td>
                          <button
                            type="button"
                            className={styles.linkButton}
                            onClick={() => openVersions(file.name, kit.sp.toServerRelative(file.path))}
                          >
                            Versions
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
            <div className={styles.small}>The 50 largest files the search index knows about, by the size of their current version.</div>
          </div>
        </PivotItem>
      </Pivot>

      <Panel
        isOpen={!!versionFile}
        onDismiss={closeVersions}
        type={PanelType.medium}
        headerText={versionFile ? `Versions of ${versionFile.name}` : ''}
        closeButtonAriaLabel="Close"
      >
        {!versions && <Spinner size={SpinnerSize.medium} label="Reading version history..." />}
        {versions && (
          <div className={styles.body}>
            <div className={styles.muted}>
              {versions.length} older version{versions.length === 1 ? '' : 's'} are stored besides the current one, using{' '}
              {formatBytes(versions.reduce((sum, v) => sum + (Number(v.Size) || 0), 0))}.
            </div>
            <div className={styles.toolbar}>
              <Dropdown
                label="When trimming"
                options={KEEP_OPTIONS}
                selectedKey={keep}
                onChange={(_, option) => {
                  if (option) {
                    setKeep(option.key as number);
                    setConfirmTrim(false);
                  }
                }}
                styles={{ root: { minWidth: 220 } }}
                disabled={!!trimProgress}
              />
              <PrimaryButton
                text={`Trim ${oldVersions.length} older version${oldVersions.length === 1 ? '' : 's'}`}
                disabled={oldVersions.length === 0 || !!trimProgress || confirmTrim}
                onClick={() => setConfirmTrim(true)}
              />
            </div>
            {/*
              The confirmation lives inside the panel on purpose: a separate dialog counts as a
              click outside the panel, which would close the panel before the trim could start.
            */}
            {confirmTrim && !trimProgress && (
              <MessageBar
                messageBarType={MessageBarType.severeWarning}
                isMultiline
                actions={
                  <div>
                    <PrimaryButton text="Yes, trim them" onClick={trim} />
                    <DefaultButton text="Cancel" onClick={() => setConfirmTrim(false)} />
                  </div>
                }
              >
                {oldVersions.length} older version{oldVersions.length === 1 ? '' : 's'} ({formatBytes(oldVersionBytes)}) will be moved to the
                site recycle bin. The current file and its newest {keep} versions are kept.
              </MessageBar>
            )}
            {trimProgress && (
              <>
                <ProgressIndicator
                  label={`Moving versions to the recycle bin: ${trimProgress.done} of ${trimProgress.total}`}
                  percentComplete={trimProgress.total > 0 ? trimProgress.done / trimProgress.total : 0}
                />
                <DefaultButton text="Stop" onClick={() => trimToken.current && trimToken.current.cancel()} />
              </>
            )}
            {versions.length > 0 && (
              <div className={styles.tableWrap}>
                <table className={styles.table}>
                  <thead>
                    <tr>
                      <th>Version</th>
                      <th className={styles.num}>Size</th>
                      <th>Saved</th>
                    </tr>
                  </thead>
                  <tbody>
                    {versions.slice(0, 200).map((version) => (
                      <tr key={version.ID}>
                        <td>{version.VersionLabel}</td>
                        <td className={styles.num}>{formatBytes(Number(version.Size))}</td>
                        <td className={styles.nowrap}>{formatDate(version.Created)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
            {versions.length > 200 && <div className={styles.small}>Showing the newest 200 of {versions.length} versions.</div>}
          </div>
        )}
      </Panel>

    </Shell>
  );
};
