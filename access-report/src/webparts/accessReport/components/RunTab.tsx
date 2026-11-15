import * as React from 'react';
import { ChoiceGroup, DefaultButton, MessageBar, MessageBarType, PrimaryButton, ProgressIndicator, SearchBox } from '@fluentui/react';
import { ISite } from '../services/AccessScan';
import { downloadText, fileStamp, reportFileName, unreadFileName } from '../services/reportCsv';
import { clearJob, GENTLE_PAUSE_MS, IRunJob, IRunToken, loadJob, newJob, processed, reportCsv, requeueUnread, runJob, RunSpeed, saveJob, unreadCsv, unreadSites } from '../services/RunJob';
import styles from './AccessReport.module.scss';
import { IConfig, IServices } from './AccessReport';
import { Card, dateTime, Empty, errorText, Facts, GRID_LIMIT, matches, RowsGrid } from './ui';

function duration(sec: number): string {
  if (sec < 90) {
    return `${Math.max(1, Math.round(sec))} seconds`;
  }
  if (sec < 5400) {
    return `${Math.round(sec / 60)} minutes`;
  }
  return `${(sec / 3600).toFixed(1)} hours`;
}

/** Rough time for a run: about one request per subject and site, more with libraries. */
export function estimateSeconds(sites: number, subjects: number, libraries: boolean, speed: RunSpeed): number {
  const requests = subjects * (libraries ? 1.3 : 1) + (libraries ? 1 : 0);
  return sites * (speed === 'gentle' ? requests * 0.35 + GENTLE_PAUSE_MS / 1000 : (requests * 0.35) / 2);
}

/** Runs the report in the browser and saves it in the report folder - what the flow does on a schedule. */
export const RunTab: React.FC<{ services: IServices; config: IConfig; onOpenTab: (key: string) => void }> = ({ services, config, onOpenTab }) => {
  const saved = React.useMemo(() => loadJob(), []);
  const [job, setJob] = React.useState<IRunJob | undefined>(saved && saved.job);
  const sitesRef = React.useRef<ISite[]>(saved ? saved.sites : []);
  const [, setTick] = React.useState(0);
  const [running, setRunning] = React.useState(false);
  const [busy, setBusy] = React.useState<string | undefined>();
  const [message, setMessage] = React.useState<{ kind: MessageBarType; text: string } | undefined>();
  const [storageOk, setStorageOk] = React.useState(true);
  const [speed, setSpeed] = React.useState<RunSpeed>('gentle');
  const [query, setQuery] = React.useState('');
  const tokenRef = React.useRef<IRunToken>({ stop: true });
  const jobRef = React.useRef<IRunJob | undefined>(job);
  jobRef.current = job;
  const mounted = React.useRef(true);

  // Save on leaving the page or closing the tab, and stop the workers.
  React.useEffect(() => {
    const onUnload = (): void => {
      if (jobRef.current) {
        saveJob(jobRef.current);
      }
    };
    window.addEventListener('beforeunload', onUnload);
    return () => {
      mounted.current = false;
      tokenRef.current.stop = true;
      onUnload();
      window.removeEventListener('beforeunload', onUnload);
    };
  }, []);

  const active = config.subjects.filter((s) => s.active && s.login);
  const { settings } = config;

  /** Writes the finished report (and the unread list, if any) into the report folder. */
  const saveToFolder = async (j: IRunJob): Promise<void> => {
    const stamp = fileStamp(new Date());
    try {
      j.savedAs = await services.store.saveFile(j.folder, reportFileName(stamp), reportCsv(j));
      if (j.unread.length) {
        await services.store.saveFile(j.folder, unreadFileName(stamp), unreadCsv(j));
      }
      saveJob(j);
      if (mounted.current) {
        const u = unreadSites(j).length;
        setMessage({
          kind: u ? MessageBarType.warning : MessageBarType.success,
          text: `Finished: ${j.rows.length.toLocaleString()} rows across ${j.total.toLocaleString()} sites, saved as ${reportFileName(stamp)} in ${j.folder}.${
            u ? ` ${u.toLocaleString()} site${u === 1 ? '' : 's'} could not be read completely: the list is saved next to the report, and "Retry unread" checks them again.` : ''
          }`
        });
      }
    } catch (e) {
      if (mounted.current) {
        setMessage({ kind: MessageBarType.error, text: `The scan finished, but the report could not be saved in ${j.folder}: ${errorText(e)} Download it instead, or fix the folder under "Who and where" and press "Save to folder".` });
      }
    }
  };

  /** Runs (or resumes) the job until it is finished, stopped or paused by errors. */
  const run = async (j: IRunJob): Promise<void> => {
    const token: IRunToken = { stop: false };
    tokenRef.current = token;
    setRunning(true);
    setMessage(undefined);
    await runJob(j, sitesRef.current, services.scanner, token, (stored) => {
      if (mounted.current) {
        setStorageOk(stored);
        setTick((t) => t + 1);
      }
    });
    if (j.finished) {
      await saveToFolder(j);
    }
    if (mounted.current) {
      setRunning(false);
      if (token.reason) {
        setMessage({ kind: MessageBarType.warning, text: token.reason });
      } else if (!j.finished) {
        setMessage({ kind: MessageBarType.info, text: 'Stopped. Progress is saved in this browser: Resume carries on from here, even after closing the page.' });
      }
      setTick((t) => t + 1);
    }
  };

  const start = async (): Promise<void> => {
    setMessage(undefined);
    setBusy('Listing sites…');
    try {
      if (!(await services.store.folderExists(settings.reportFolder))) {
        setMessage({ kind: MessageBarType.error, text: `The report folder ${settings.reportFolder} does not exist. Choose an existing folder under "Who and where".` });
        return;
      }
      const r = await services.lister.list(services.homeWebUrl, settings.siteFilter, (n) => mounted.current && setBusy(`Listing sites… ${n.toLocaleString()}`));
      if (!r.sites.length) {
        setMessage({ kind: MessageBarType.warning, text: settings.siteFilter ? `No site matches "${settings.siteFilter}".` : 'No sites were found.' });
        return;
      }
      const j = newJob(
        active.map((s) => ({ label: s.label, login: s.login })),
        settings.includeLibraries,
        settings.siteFilter,
        settings.reportFolder,
        speed,
        r.sites.length
      );
      sitesRef.current = r.sites;
      setStorageOk(saveJob(j, r.sites));
      setJob(j);
      setQuery('');
      setBusy(undefined);
      await run(j);
      if (r.truncated && mounted.current) {
        setMessage({ kind: MessageBarType.warning, text: 'The site list was cut short by search (more than 50,000 results); the report covers the sites that were listed.' });
      }
    } catch (e) {
      setMessage({ kind: MessageBarType.error, text: errorText(e) });
    } finally {
      setBusy(undefined);
    }
  };

  const retryUnread = async (): Promise<void> => {
    if (job) {
      requeueUnread(job);
      await run(job);
    }
  };

  const discard = (): void => {
    clearJob();
    setJob(undefined);
    sitesRef.current = [];
    setMessage(undefined);
  };

  const rowsAll = React.useMemo(() => (job ? job.rows.map((r) => r.slice(1) as string[]) : []), [job, job ? job.rows.length : 0]);
  const rows = React.useMemo(() => rowsAll.filter((r) => matches(r, query)), [rowsAll, query]);
  const done = job ? processed(job) : 0;
  const canStart = !running && !busy;
  const unreadCount = job ? unreadSites(job).length : 0;

  return (
    <div className={styles.tab}>
      <Card title="Run the report now">
        <p className={styles.hint}>
          Checks every site you can open for the people and groups below and saves the result as a CSV in the report folder. The monthly flow does exactly this on its own; use this to try it out or
          when you need a report today.
        </p>
        <Facts
          items={[
            { label: 'People and groups', value: active.length ? `${active.length.toLocaleString()} included` : 'None yet', warn: !active.length },
            { label: 'Libraries', value: settings.includeLibraries ? 'Included' : 'Sites only' },
            { label: 'Sites', value: settings.siteFilter ? `Containing “${settings.siteFilter}”` : 'All you can open' },
            { label: 'Report folder', value: <span className={styles.code}>{settings.reportFolder}</span> }
          ]}
        />
        <div style={{ marginTop: 12 }}>
          <ChoiceGroup
            label="Speed"
            selectedKey={speed}
            onChange={(_, o) => o && setSpeed(o.key as RunSpeed)}
            options={[
              { key: 'gentle', text: 'Gentle (recommended): one site at a time, with a short pause' },
              { key: 'normal', text: 'Normal: two sites at a time' }
            ]}
            disabled={!canStart}
          />
        </div>
        <div className={styles.actions} style={{ marginTop: 12 }}>
          <PrimaryButton iconProps={{ iconName: 'Play' }} onClick={start} disabled={!canStart || !active.length}>
            {busy || (job ? 'Run a new report' : 'Run report')}
          </PrimaryButton>
          <DefaultButton onClick={() => onOpenTab('setup')} disabled={running}>
            Change who and where
          </DefaultButton>
        </div>
        {!active.length && <p className={`${styles.muted} ${styles.small}`}>Add at least one person or group under “Who and where” first.</p>}
        {job && canStart && <p className={`${styles.muted} ${styles.small}`}>Running a new report replaces the one shown below (its saved file stays in the folder).</p>}
        {!storageOk && <MessageBar messageBarType={MessageBarType.warning}>This browser would not save the progress (private window or storage full): keep this page open until the run finishes.</MessageBar>}
      </Card>

      {message && (
        <MessageBar messageBarType={message.kind} isMultiline onDismiss={() => setMessage(undefined)}>
          {message.text}
        </MessageBar>
      )}

      {job && (
        <>
          <Card
            title={job.finished ? 'Last run' : running ? 'Running' : 'Unfinished run'}
            right={
              <div className={styles.actions}>
                {running ? (
                  <DefaultButton
                    iconProps={{ iconName: 'Pause' }}
                    onClick={() => {
                      tokenRef.current.stop = true;
                    }}
                  >
                    Stop (resume later)
                  </DefaultButton>
                ) : (
                  <>
                    {!job.finished && (
                      <PrimaryButton iconProps={{ iconName: 'Play' }} onClick={() => run(job)} disabled={!!busy}>
                        Resume
                      </PrimaryButton>
                    )}
                    {job.finished && unreadCount > 0 && (
                      <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={retryUnread} disabled={!!busy}>
                        Retry {unreadCount.toLocaleString()} unread
                      </DefaultButton>
                    )}
                    {job.finished && !job.savedAs && (
                      <DefaultButton iconProps={{ iconName: 'Save' }} onClick={() => saveToFolder(job).then(() => setTick((t) => t + 1))} disabled={!!busy}>
                        Save to folder
                      </DefaultButton>
                    )}
                    <DefaultButton onClick={discard} disabled={!!busy}>
                      Discard
                    </DefaultButton>
                  </>
                )}
              </div>
            }
          >
            <ProgressIndicator
              label={`${done.toLocaleString()} of ${job.total.toLocaleString()} sites checked${running ? '' : job.finished ? ' (finished)' : ' (stopped)'}`}
              description={`Started ${dateTime(new Date(job.started))} · ${job.subjects.length} people and groups · ${job.speed} speed${job.includeLibraries ? ' · with libraries' : ''}${
                running ? ` · about ${duration(estimateSeconds(job.total - done, job.subjects.length, job.includeLibraries, job.speed))} left` : ''
              }`}
              percentComplete={job.total ? done / job.total : undefined}
            />
            <div style={{ marginTop: 10 }}>
              <Facts
                big
                items={[
                  { label: 'Sites with access', value: rowsAll.filter((r) => r[3] === 'Site').length.toLocaleString() },
                  { label: 'Libraries with access', value: rowsAll.filter((r) => r[3] === 'Library').length.toLocaleString() },
                  { label: 'Libraries with own permissions', value: rowsAll.filter((r) => r[8] === 'Yes').length.toLocaleString() },
                  { label: 'Sites not read completely', value: unreadCount.toLocaleString(), warn: unreadCount > 0 }
                ]}
              />
            </div>
            {job.savedAs && (
              <p className={styles.small} style={{ marginBottom: 0 }}>
                Saved as{' '}
                <a className={styles.link} href={services.homeWebUrl.replace(/^(https:\/\/[^/]+).*$/, '$1') + encodeURI(job.savedAs)} target="_blank" rel="noopener noreferrer">
                  {job.savedAs}
                </a>
              </p>
            )}
          </Card>

          <Card
            title={`${rows.length.toLocaleString()} ${rows.length === 1 ? 'row' : 'rows'}${query.trim() ? ' match' : ''}`}
            right={
              <div className={styles.actions}>
                <SearchBox placeholder="Filter by who, site, library or how" value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { width: 280 } }} />
                <DefaultButton iconProps={{ iconName: 'Download' }} onClick={() => downloadText(`access-report${job.finished ? '' : '-partial'}.csv`, reportCsv(job))} disabled={!rowsAll.length}>
                  Download CSV{job.finished ? '' : ' (so far)'}
                </DefaultButton>
                {job.unread.length > 0 && (
                  <DefaultButton iconProps={{ iconName: 'Download' }} onClick={() => downloadText('access-report-unread.csv', unreadCsv(job))}>
                    Download unread list
                  </DefaultButton>
                )}
              </div>
            }
          >
            {rows.length === 0 ? <Empty text={running ? 'Nothing found yet.' : rowsAll.length ? 'No row matches the filter.' : 'No access found.'} /> : <RowsGrid rows={rows} />}
            {rows.length > GRID_LIMIT && (
              <p className={`${styles.muted} ${styles.small}`}>
                Showing the first {GRID_LIMIT.toLocaleString()}; the file has all {rows.length.toLocaleString()}.
              </p>
            )}
          </Card>
        </>
      )}
    </div>
  );
};
