import * as React from 'react';
import { useCallback, useEffect, useRef, useState } from 'react';
import { DefaultButton, Dropdown, Icon, IDropdownOption, PrimaryButton, ProgressIndicator } from '@fluentui/react';

import { cacheGet, cacheKey, cacheSet } from '../../common/cache';
import { downloadCsv } from '../../common/csv';
import { extensionOf, fileIcon, formatBytes, formatDate, formatDateTime, formatNumber } from '../../common/format';
import { getLibraries, ILibrary } from '../../common/libraries';
import { CancelToken, errorText, isCancelled } from '../../common/RequestQueue';
import { folderOf, search, SEARCH_PAGE_SIZE, underPath } from '../../common/search';
import { IToolkitComponentProps } from '../../common/ToolkitWebPart';
import { Empty, Shell } from '../../common/ui/Shell';
import styles from '../../common/ui/Toolkit.module.scss';
import { useMounted, useToolkit } from '../../common/ui/useToolkit';
import { addToGroups, collectDuplicates, IDuplicateGroup } from './duplicates';

export interface IDuplicateFileDetectorProps {
  title: string;
}

interface ISavedResult {
  scannedAt: string;
  scanned: number;
  complete: boolean;
  groups: IDuplicateGroup[];
  totalGroups: number;
  wasted: number;
}

/** A hard stop so one scan can never run away on a huge site. */
const MAX_FILES = 200000;
const MAX_GROUPS_KEPT = 300;
const SITE_SCOPE = '__site__';

export const DuplicateFileDetector: React.FC<IToolkitComponentProps<IDuplicateFileDetectorProps>> = (props) => {
  const kit = useToolkit(props.context);
  const mounted = useMounted();

  const [libraries, setLibraries] = useState<ILibrary[]>([]);
  const [scope, setScope] = useState<string>(SITE_SCOPE);
  const [result, setResult] = useState<ISavedResult | undefined>(undefined);
  const [scanning, setScanning] = useState(false);
  const [paused, setPaused] = useState(false);
  const [progress, setProgress] = useState<{ scanned: number; total: number }>({ scanned: 0, total: 0 });
  const [error, setError] = useState<string | undefined>(undefined);
  const [open, setOpen] = useState<string | undefined>(undefined);

  // Scan state survives a pause so that Resume carries on where it stopped.
  const groupsRef = useRef<Record<string, IDuplicateGroup>>({});
  const lastDocId = useRef<string>('0');
  const scannedRef = useRef(0);
  const totalRef = useRef(0);
  const tokenRef = useRef<CancelToken | undefined>(undefined);

  const resultKey = cacheKey(kit.scope, `duplicates:${scope}`);

  useEffect(() => {
    getLibraries(kit.sp, kit.scope)
      .then((list) => {
        if (mounted.current) {
          setLibraries(list);
        }
      })
      .catch(() => undefined);
  }, [kit, mounted]);

  useEffect(() => {
    // A different place to look starts from a clean slate.
    groupsRef.current = {};
    lastDocId.current = '0';
    scannedRef.current = 0;
    totalRef.current = 0;
    setResult(cacheGet<ISavedResult>('session', resultKey));
    setPaused(false);
    setOpen(undefined);
  }, [resultKey]);

  useEffect(
    () => () => {
      if (tokenRef.current) {
        tokenRef.current.cancel();
      }
    },
    []
  );

  const scopeUrl = useCallback((): string => {
    if (scope === SITE_SCOPE) {
      return kit.sp.webUrl;
    }
    const library = libraries.filter((l) => l.id === scope)[0];
    return library ? `${kit.sp.origin}${library.path}` : kit.sp.webUrl;
  }, [kit, libraries, scope]);

  const finish = useCallback(
    (complete: boolean) => {
      const all = collectDuplicates(groupsRef.current);
      const saved: ISavedResult = {
        scannedAt: new Date().toISOString(),
        scanned: scannedRef.current,
        complete,
        groups: all.slice(0, MAX_GROUPS_KEPT),
        totalGroups: all.length,
        wasted: all.reduce((sum, g) => sum + g.size * (g.copies.length - 1), 0)
      };
      setResult(saved);
      if (complete) {
        cacheSet('session', resultKey, saved, 60);
      }
    },
    [resultKey]
  );

  const run = useCallback(
    (resume: boolean) => {
      if (!resume) {
        groupsRef.current = {};
        lastDocId.current = '0';
        scannedRef.current = 0;
        totalRef.current = 0;
        setResult(undefined);
      }
      const token = new CancelToken();
      tokenRef.current = token;
      setScanning(true);
      setPaused(false);
      setError(undefined);
      setProgress({ scanned: scannedRef.current, total: totalRef.current });
      const base = `IsDocument:1 -FileExtension:aspx ${underPath(scopeUrl())}`;

      const loop = async (): Promise<boolean> => {
        // Pages are fetched one after another through the page queue, never in parallel.
        for (;;) {
          const page = await search(kit.sp, {
            kql: `${base} IndexDocId>${lastDocId.current}`,
            select: ['DocId', 'Path', 'Size', 'LastModifiedTime', 'ModifiedBy'],
            rowLimit: SEARCH_PAGE_SIZE,
            sort: '[DocId]:ascending',
            token
          });
          if (totalRef.current === 0) {
            totalRef.current = page.total;
          }
          for (const row of page.rows) {
            addToGroups(groupsRef.current, row.Path, Number(row.Size) || 0, row.LastModifiedTime, row.ModifiedBy);
          }
          scannedRef.current += page.rows.length;
          const last = page.rows.length > 0 ? page.rows[page.rows.length - 1].DocId : '';
          if (mounted.current) {
            setProgress({ scanned: scannedRef.current, total: Math.max(totalRef.current, scannedRef.current) });
          }
          if (page.rows.length < SEARCH_PAGE_SIZE || !last || last === lastDocId.current) {
            return true;
          }
          lastDocId.current = last;
          if (scannedRef.current >= MAX_FILES) {
            return false;
          }
        }
      };

      loop()
        .then((complete) => {
          if (!mounted.current) {
            return;
          }
          setScanning(false);
          finish(complete);
        })
        .catch((err) => {
          if (!mounted.current) {
            return;
          }
          setScanning(false);
          if (isCancelled(err)) {
            setPaused(true);
            finish(false);
          } else {
            setPaused(scannedRef.current > 0);
            finish(false);
            setError(`The scan stopped: ${errorText(err)} You can resume it.`);
          }
        });
    },
    [finish, kit, mounted, scopeUrl]
  );

  const exportCsv = useCallback(() => {
    const rows: Array<Array<string | number>> = [];
    // A scan from this session exports every set; a result restored from the cache has the kept sets only.
    const all = collectDuplicates(groupsRef.current);
    for (const group of all.length > 0 ? all : result ? result.groups : []) {
      for (const copy of group.copies) {
        rows.push([group.name, group.size, group.copies.length, kit.sp.toServerRelative(folderOf(copy.path)), copy.modified, copy.modifiedBy, copy.path]);
      }
    }
    downloadCsv('duplicate-files', ['File name', 'Size bytes', 'Copies', 'Folder', 'Modified', 'Modified by', 'Link'], rows);
  }, [kit, result]);

  const scopeOptions: IDropdownOption[] = [{ key: SITE_SCOPE, text: 'This whole site' }].concat(
    libraries.map((l) => ({ key: l.id, text: `Library: ${l.title}` }))
  );

  return (
    <Shell
      icon="Copy"
      title={props.properties.title || 'Duplicate File Detector'}
      subtitle="Finds files that have the same name and the same size."
      error={error}
      onDismissError={() => setError(undefined)}
      actions={
        <>
          {scanning ? (
            <DefaultButton iconProps={{ iconName: 'Pause' }} text="Pause" onClick={() => tokenRef.current && tokenRef.current.cancel()} />
          ) : (
            <>
              {paused && <PrimaryButton iconProps={{ iconName: 'Play' }} text="Resume" onClick={() => run(true)} />}
              <DefaultButton iconProps={{ iconName: 'Search' }} text={result || paused ? 'Scan again' : 'Scan'} primary={!paused} onClick={() => run(false)} />
            </>
          )}
          <DefaultButton iconProps={{ iconName: 'Download' }} text="Export CSV" disabled={!result || result.groups.length === 0 || scanning} onClick={exportCsv} />
        </>
      }
    >
      <div className={styles.toolbar}>
        <Dropdown
          label="Where to look"
          options={scopeOptions}
          selectedKey={scope}
          onChange={(_, option) => option && setScope(String(option.key))}
          disabled={scanning}
          styles={{ root: { minWidth: 260 } }}
        />
      </div>

      {scanning && (
        <ProgressIndicator
          label={`Scanned ${formatNumber(progress.scanned)} of about ${formatNumber(progress.total)} files`}
          description="The scan asks for 500 files at a time, one request after another, so it stays gentle on SharePoint."
          percentComplete={progress.total > 0 ? Math.min(progress.scanned / progress.total, 1) : undefined}
        />
      )}

      {!scanning && !result && <Empty icon="Copy" text="Press Scan to look for duplicate files. Nothing is changed or deleted." />}

      {result && (
        <>
          <div className={styles.kpis}>
            <div className={styles.kpi}>
              <div className={styles.kpiValue}>{formatNumber(result.totalGroups)}</div>
              <div className={styles.kpiLabel}>Sets of duplicates</div>
            </div>
            <div className={styles.kpi}>
              <div className={styles.kpiValue}>{formatBytes(result.wasted)}</div>
              <div className={styles.kpiLabel}>Used by the extra copies</div>
            </div>
            <div className={styles.kpi}>
              <div className={styles.kpiValue}>{formatNumber(result.scanned)}</div>
              <div className={styles.kpiLabel}>Files scanned</div>
            </div>
          </div>
          <div className={styles.small}>
            {result.complete ? `Scan finished ${formatDateTime(result.scannedAt)}.` : `Partial result: the scan was ${scanning ? 'still running' : 'stopped'} after ${formatNumber(result.scanned)} files.`}
            {result.scanned >= MAX_FILES ? ` Scans stop at ${formatNumber(MAX_FILES)} files; choose a single library to go deeper.` : ''}
          </div>

          {result.groups.length === 0 && result.complete && <Empty icon="Completed" text="No duplicate files were found." />}

          {result.groups.length > 0 && (
            <div className={styles.tableWrap}>
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th>File</th>
                    <th className={styles.num}>Copies</th>
                    <th className={styles.num}>Size each</th>
                    <th className={styles.num}>Extra space used</th>
                    <th />
                  </tr>
                </thead>
                <tbody>
                  {result.groups.map((group) => {
                    const key = `${group.name}|${group.size}`;
                    const isOpen = open === key;
                    return (
                      <React.Fragment key={key}>
                        <tr>
                          <td>
                            <div className={styles.fileCell}>
                              <Icon iconName={fileIcon(extensionOf(group.name))} aria-hidden="true" />
                              <span className={styles.fileName}>{group.name}</span>
                            </div>
                          </td>
                          <td className={styles.num}>{group.copies.length}</td>
                          <td className={styles.num}>{formatBytes(group.size)}</td>
                          <td className={styles.num}>{formatBytes(group.size * (group.copies.length - 1))}</td>
                          <td>
                            <button type="button" className={styles.linkButton} aria-expanded={isOpen} onClick={() => setOpen(isOpen ? undefined : key)}>
                              {isOpen ? 'Hide copies' : 'Show copies'}
                            </button>
                          </td>
                        </tr>
                        {isOpen &&
                          group.copies.map((copy) => (
                            <tr key={copy.path}>
                              <td colSpan={3}>
                                <a className={styles.link} href={copy.path} target="_blank" rel="noopener noreferrer">
                                  {kit.sp.toServerRelative(folderOf(copy.path))}
                                </a>
                              </td>
                              <td className={styles.nowrap}>{formatDate(copy.modified)}</td>
                              <td>{copy.modifiedBy}</td>
                            </tr>
                          ))}
                      </React.Fragment>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
          {result.totalGroups > result.groups.length && (
            <div className={styles.small}>
              Showing the {result.groups.length} sets that use the most space, of {formatNumber(result.totalGroups)}.
            </div>
          )}
        </>
      )}

      <div className={styles.small}>
        Files are matched on name and size, which catches straight copies. Two different files that happen to share both would also be listed,
        so open them before deleting anything.
      </div>
    </Shell>
  );
};
