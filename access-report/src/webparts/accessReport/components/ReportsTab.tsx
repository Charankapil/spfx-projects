import * as React from 'react';
import { DefaultButton, Dropdown, MessageBar, MessageBarType, SearchBox, Spinner, SpinnerSize } from '@fluentui/react';
import { CHANGE_HEADERS, changeRows, compareReports, IChange } from '../services/compare';
import { downloadText, parseReport, toReportCsv } from '../services/reportCsv';
import { IReportFile } from '../services/Store';
import styles from './AccessReport.module.scss';
import { IConfig, IServices } from './AccessReport';
import { Card, ChangesGrid, dateTime, Empty, errorText, Facts, GRID_LIMIT, matches, RowsGrid } from './ui';

const size = (bytes: number): string => (bytes < 1024 ? `${bytes} B` : bytes < 1048576 ? `${(bytes / 1024).toFixed(1)} KB` : `${(bytes / 1048576).toFixed(1)} MB`);

/** The files in the report folder: open one, and see what changed since an earlier one. */
export const ReportsTab: React.FC<{ services: IServices; config: IConfig }> = ({ services, config }) => {
  const folder = config.settings.reportFolder;
  const origin = services.homeWebUrl.replace(/^(https:\/\/[^/]+).*$/, '$1');
  const [files, setFiles] = React.useState<IReportFile[] | undefined>();
  const [error, setError] = React.useState<string | undefined>();
  const [open, setOpen] = React.useState<IReportFile | undefined>();
  const [rows, setRows] = React.useState<string[][] | undefined>();
  const [against, setAgainst] = React.useState<string | undefined>();
  const [changes, setChanges] = React.useState<IChange[] | undefined>();
  const [busy, setBusy] = React.useState(false);
  const [query, setQuery] = React.useState('');
  const cache = React.useRef<{ [url: string]: string[][] }>({});
  const alive = React.useRef(true);
  React.useEffect(
    () => () => {
      alive.current = false;
    },
    []
  );

  const read = async (f: IReportFile): Promise<string[][]> => {
    if (!cache.current[f.serverRelativeUrl]) {
      cache.current[f.serverRelativeUrl] = parseReport(await services.store.readFile(f.serverRelativeUrl));
    }
    return cache.current[f.serverRelativeUrl];
  };

  const guard = async (work: () => Promise<void>): Promise<void> => {
    setBusy(true);
    setError(undefined);
    try {
      await work();
    } catch (e) {
      if (alive.current) {
        setError(errorText(e));
      }
    } finally {
      if (alive.current) {
        setBusy(false);
      }
    }
  };

  const load = React.useCallback(
    () =>
      guard(async () => {
        const list = await services.store.reports(folder);
        if (alive.current) {
          setFiles(list);
        }
      }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [services, folder]
  );
  React.useEffect(() => {
    load().catch(() => undefined);
  }, [load]);

  const reports = (files || []).filter((f) => !f.unread);
  const show = (f: IReportFile): Promise<void> =>
    guard(async () => {
      const r = await read(f);
      if (alive.current) {
        setOpen(f);
        setRows(r);
        setAgainst(undefined);
        setChanges(undefined);
        setQuery('');
      }
    });
  const compare = (url: string): Promise<void> =>
    guard(async () => {
      const earlier = reports.filter((f) => f.serverRelativeUrl === url)[0];
      if (!earlier || !rows) {
        return;
      }
      const c = compareReports(await read(earlier), rows);
      if (alive.current) {
        setAgainst(url);
        setChanges(c);
      }
    });

  const older = open ? reports.filter((f) => f.stamp < open.stamp) : [];
  const shownRows = (rows || []).filter((r) => matches(r, query));
  const shownChanges = (changes || []).filter((c) => matches([c.kind as string].concat(c.row), query));
  const count = (kind: string): string => (changes || []).filter((c) => c.kind === kind).length.toLocaleString();

  return (
    <div className={styles.tab}>
      {error && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline onDismiss={() => setError(undefined)}>
          {error}
        </MessageBar>
      )}
      <Card
        title="Reports in the folder"
        right={
          <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={load} disabled={busy}>
            Refresh
          </DefaultButton>
        }
      >
        <p className={styles.hint}>
          Every run - monthly or from this page - leaves a file in <span className={styles.code}>{folder}</span>. Times are in UTC.
        </p>
        {!files ? (
          <Spinner size={SpinnerSize.small} label="Loading…" labelPosition="right" style={{ justifyContent: 'flex-start' }} />
        ) : files.length === 0 ? (
          <Empty text="No reports yet. Run one from the first tab, or wait for the monthly flow." />
        ) : (
          <div className={styles.tableWrap}>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th>Run (UTC)</th>
                  <th>File</th>
                  <th>Saved by</th>
                  <th>Size</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {files.map((f) => (
                  <tr key={f.serverRelativeUrl}>
                    <td style={{ whiteSpace: 'nowrap' }}>{f.at.toISOString().substring(0, 16).replace('T', ' ')}</td>
                    <td>
                      <a className={styles.link} href={origin + encodeURI(f.serverRelativeUrl)} target="_blank" rel="noopener noreferrer">
                        {f.name}
                      </a>{' '}
                      {f.unread && <span className={`${styles.pill} ${styles.pillChanged}`}>not read</span>}
                    </td>
                    <td>{f.modifiedBy}</td>
                    <td style={{ whiteSpace: 'nowrap' }}>{size(f.size)}</td>
                    <td style={{ textAlign: 'right' }}>
                      {!f.unread && (
                        <DefaultButton onClick={() => show(f)} disabled={busy} primary={!!open && open.serverRelativeUrl === f.serverRelativeUrl}>
                          {open && open.serverRelativeUrl === f.serverRelativeUrl ? 'Open' : 'View'}
                        </DefaultButton>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>

      {open && rows && (
        <Card
          title={changes ? `What changed in ${open.name}` : `${open.name} · ${dateTime(open.at)}`}
          right={
            <div className={styles.actions}>
              <SearchBox placeholder="Filter" value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { width: 220 } }} />
              <Dropdown
                ariaLabel="Compare with an earlier report"
                placeholder={older.length ? 'Compare with an earlier report' : 'No earlier report to compare with'}
                selectedKey={against || null}
                options={older.map((f) => ({ key: f.serverRelativeUrl, text: f.name }))}
                onChange={(_, o) => o && compare(String(o.key))}
                disabled={busy || !older.length}
                styles={{ root: { width: 300 } }}
              />
              {changes && (
                <>
                  <DefaultButton
                    onClick={() => {
                      setChanges(undefined);
                      setAgainst(undefined);
                    }}
                  >
                    Show the report
                  </DefaultButton>
                  <DefaultButton iconProps={{ iconName: 'Download' }} onClick={() => downloadText(`access-changes-${open.stamp}.csv`, toReportCsv(CHANGE_HEADERS, changeRows(changes)))} disabled={!changes.length}>
                    Download changes
                  </DefaultButton>
                </>
              )}
            </div>
          }
        >
          {changes ? (
            <>
              <Facts
                big
                items={[
                  { label: 'New access', value: count('Added') },
                  { label: 'Access level changed', value: count('Changed') },
                  { label: 'Access removed', value: count('Removed') }
                ]}
              />
              <div style={{ marginTop: 12 }}>{shownChanges.length ? <ChangesGrid changes={shownChanges} /> : <Empty text={changes.length ? 'No change matches the filter.' : 'Nothing changed between the two reports.'} />}</div>
            </>
          ) : shownRows.length ? (
            <RowsGrid rows={shownRows} />
          ) : (
            <Empty text={rows.length ? 'No row matches the filter.' : 'This report has no rows: nobody in it had access anywhere.'} />
          )}
          {(changes ? shownChanges.length : shownRows.length) > GRID_LIMIT && <p className={`${styles.muted} ${styles.small}`}>Showing the first {GRID_LIMIT.toLocaleString()}; the file has everything.</p>}
        </Card>
      )}
    </div>
  );
};
