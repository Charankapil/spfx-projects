import * as React from 'react';
import { useCallback, useEffect, useRef, useState } from 'react';
import { DefaultButton, Dropdown, Icon, IDropdownOption, Spinner, SpinnerSize } from '@fluentui/react';

import { downloadCsv } from '../../common/csv';
import { extensionOf, fileIcon, formatAge, formatBytes, formatDate, formatNumber } from '../../common/format';
import { errorText } from '../../common/RequestQueue';
import { fileNameOf, folderOf, IRefinerValue, ISearchResult, search, underPath } from '../../common/search';
import { IToolkitComponentProps } from '../../common/ToolkitWebPart';
import { Empty, Shell } from '../../common/ui/Shell';
import styles from '../../common/ui/Toolkit.module.scss';
import { useMounted, useToolkit } from '../../common/ui/useToolkit';

export interface IStaleContentRadarProps {
  title: string;
  months: string;
}

interface IStaleFile {
  name: string;
  path: string;
  folder: string;
  size: number;
  modified: string;
  modifiedBy: string;
}

const PAGE_SIZE = 100;

const MONTH_OPTIONS: IDropdownOption[] = [
  { key: 6, text: '6 months' },
  { key: 12, text: '1 year' },
  { key: 24, text: '2 years' },
  { key: 36, text: '3 years' },
  { key: 60, text: '5 years' }
];

function cutoffDate(months: number): string {
  const date = new Date();
  date.setMonth(date.getMonth() - months);
  return date.toISOString().substring(0, 10);
}

const Breakdown: React.FC<{ title: string; values: IRefinerValue[] }> = (props) => {
  const top = props.values.slice(0, 8);
  const max = top.length > 0 ? top[0].count : 0;
  if (top.length === 0) {
    return null;
  }
  return (
    <div className={styles.card}>
      <h3 className={styles.cardTitle}>{props.title}</h3>
      <div className={styles.stack}>
        {top.map((value) => (
          <div key={value.name}>
            <div className={styles.spread}>
              <span>{value.name}</span>
              <span className={styles.small}>{formatNumber(value.count)}</span>
            </div>
            <div className={styles.bar}>
              <div className={styles.barFill} style={{ width: `${max > 0 ? (value.count / max) * 100 : 0}%` }} />
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

export const StaleContentRadar: React.FC<IToolkitComponentProps<IStaleContentRadarProps>> = (props) => {
  const kit = useToolkit(props.context);
  const mounted = useMounted();
  const configured = parseInt(props.properties.months, 10) || 12;

  const [months, setMonths] = useState<number>(configured);
  const [files, setFiles] = useState<IStaleFile[]>([]);
  const [total, setTotal] = useState<number | undefined>(undefined);
  const [byType, setByType] = useState<IRefinerValue[]>([]);
  const [byAuthor, setByAuthor] = useState<IRefinerValue[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);
  const seq = useRef(0);

  useEffect(() => {
    setMonths(configured);
  }, [configured]);

  const query = useCallback(
    async (startRow: number, withRefiners: boolean): Promise<ISearchResult> => {
      const base = {
        kql: `IsDocument:1 ${underPath(kit.sp.webUrl)} LastModifiedTime<${cutoffDate(months)}`,
        select: ['Path', 'Size', 'LastModifiedTime', 'ModifiedBy'],
        rowLimit: PAGE_SIZE,
        startRow,
        sort: 'LastModifiedTime:ascending'
      };
      if (!withRefiners) {
        return search(kit.sp, base);
      }
      try {
        return await search(kit.sp, { ...base, refiners: 'FileType,DisplayAuthor' });
      } catch {
        // Some tenants reject a refiner; the list itself still works without the breakdowns.
        return search(kit.sp, base);
      }
    },
    [kit, months]
  );

  const load = useCallback(
    (startRow: number) => {
      const current = ++seq.current;
      setLoading(true);
      setError(undefined);
      query(startRow, startRow === 0)
        .then((result) => {
          if (!mounted.current || current !== seq.current) {
            return;
          }
          const rows: IStaleFile[] = result.rows.map((row) => ({
            name: fileNameOf(row.Path),
            path: row.Path,
            folder: kit.sp.toServerRelative(folderOf(row.Path)),
            size: Number(row.Size) || 0,
            modified: row.LastModifiedTime,
            modifiedBy: row.ModifiedBy
          }));
          if (startRow === 0) {
            setFiles(rows);
            setTotal(result.total);
            setByType(result.refiners.FileType || []);
            setByAuthor(result.refiners.DisplayAuthor || []);
          } else {
            setFiles((existing) => existing.concat(rows));
          }
          setLoading(false);
        })
        .catch((err) => {
          if (!mounted.current || current !== seq.current) {
            return;
          }
          setLoading(false);
          setError(`Stale files could not be found: ${errorText(err)}`);
        });
    },
    [kit, mounted, query]
  );

  useEffect(() => {
    load(0);
  }, [load]);

  const exportCsv = useCallback(() => {
    downloadCsv(
      'stale-content',
      ['File', 'Folder', 'Size bytes', 'Last modified', 'Last modified by', 'Link'],
      files.map((f) => [f.name, f.folder, f.size, f.modified, f.modifiedBy, f.path])
    );
  }, [files]);

  const hasMore = total !== undefined && files.length < total && files.length > 0;

  return (
    <Shell
      icon="History"
      title={props.properties.title || 'Stale Content Radar'}
      subtitle="Files in this site that nobody has changed for a long time."
      error={error}
      onDismissError={() => setError(undefined)}
      actions={
        <>
          <Dropdown
            ariaLabel="Stale after"
            options={MONTH_OPTIONS}
            selectedKey={months}
            onChange={(_, option) => option && setMonths(option.key as number)}
            styles={{ root: { minWidth: 120 } }}
            disabled={loading}
          />
          <DefaultButton iconProps={{ iconName: 'Download' }} text="Export CSV" disabled={files.length === 0} onClick={exportCsv} />
        </>
      }
    >
      {total !== undefined && (
        <div className={styles.kpis}>
          <div className={styles.kpi}>
            <div className={styles.kpiValue}>{formatNumber(total)}</div>
            <div className={styles.kpiLabel}>Files unchanged since {formatDate(`${cutoffDate(months)}T00:00:00`)}</div>
          </div>
          <div className={styles.kpi}>
            <div className={styles.kpiValue}>{formatBytes(files.reduce((sum, f) => sum + f.size, 0))}</div>
            <div className={styles.kpiLabel}>Size of the {formatNumber(files.length)} files listed below</div>
          </div>
        </div>
      )}

      {(byType.length > 0 || byAuthor.length > 0) && (
        <div className={styles.grid}>
          <Breakdown title="By file type" values={byType} />
          <Breakdown title="By author" values={byAuthor} />
        </div>
      )}

      {!loading && total === 0 && <Empty icon="Completed" text="Nothing stale here: every file has been changed within the chosen period." />}

      {files.length > 0 && (
        <div className={styles.tableWrap}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th>File</th>
                <th>Last modified</th>
                <th>Age</th>
                <th>Last modified by</th>
                <th className={styles.num}>Size</th>
              </tr>
            </thead>
            <tbody>
              {files.map((file) => (
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
                  <td className={styles.nowrap}>{formatDate(file.modified)}</td>
                  <td className={styles.nowrap}>{formatAge(file.modified)}</td>
                  <td>{file.modifiedBy}</td>
                  <td className={styles.num}>{formatBytes(file.size)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {loading && <Spinner size={SpinnerSize.medium} label="Asking the search index..." />}

      {hasMore && !loading && (
        <div className={styles.footer}>
          <DefaultButton text={`Load ${PAGE_SIZE} more`} onClick={() => load(files.length)} />
        </div>
      )}

      <div className={styles.small}>
        Oldest first. Figures come from the search index, so very recent changes can take a short while to show.
      </div>
    </Shell>
  );
};
