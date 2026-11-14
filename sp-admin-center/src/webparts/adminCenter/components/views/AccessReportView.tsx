import * as React from 'react';
import {
  Checkbox,
  ChoiceGroup,
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  Dropdown,
  IColumn,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  ProgressIndicator,
  SearchBox,
  SelectionMode,
  TextField
} from '@fluentui/react';
import { CancelledError } from '../../core/SPClient';
import {
  ACCESS_HEADERS,
  AccessScanner,
  accessRows,
  clearJob,
  IAccessRow,
  IScanJob,
  IScanPrincipal,
  loadJob,
  remaining,
  saveJob,
  scanError,
  ScanSpeed
} from '../../services/AccessScan';
import { downloadCsv } from '../../services/exportCsv';
import { formatDateTime } from '../../services/format';
import { hasScope } from '../../services/GraphApi';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { graphErrorText, useGraphScopes } from '../shared/graph';
import { getInventory, setInventory } from '../shared/inventory';
import { Card, Empty, Pill, StatTile, ViewHeader } from '../shared/ui';

/** Pause between sites in gentle mode, on top of SPClient's own pacing. */
const GENTLE_PAUSE_MS = 500;
/** This many failures in a row (network down, signed out...) pauses the scan instead of failing every site. */
const MAX_FAILS_IN_A_ROW = 8;
const SAVE_EVERY_MS = 5000;

type SiteRef = { url: string; title: string };

function secondsPerSite(p: IScanPrincipal['kind'], libs: boolean, speed: ScanSpeed): number {
  const requests = p === 'user' ? (libs ? 2.3 : 1) : 1.1;
  return speed === 'gentle' ? requests * 0.35 + GENTLE_PAUSE_MS / 1000 : (requests * 0.35) / 2;
}

function duration(sec: number): string {
  if (sec < 90) {
    return `${Math.max(1, Math.round(sec))} seconds`;
  }
  if (sec < 5400) {
    return `${Math.round(sec / 60)} minutes`;
  }
  return `${(sec / 3600).toFixed(1)} hours`;
}

/**
 * Every site and library a person or a group can reach, with the access level and how they
 * get it. Gentle on SharePoint by default (one request at a time with a pause), saves its
 * progress in this browser every few seconds and resumes where it stopped.
 */
export const AccessReportView: React.FC = () => {
  const ctx = useAdmin();
  const scanner = React.useMemo(() => new AccessScanner(ctx.client, ctx.homeWebUrl), [ctx.client, ctx.homeWebUrl]);
  const access = useGraphScopes();
  const usersOk = !!access.scopes && hasScope(access.scopes, 'User.Read.All');
  const groupsOk = !!access.scopes && hasScope(access.scopes, 'Group.Read.All');

  const [kind, setKind] = React.useState<'user' | 'group'>('user');
  const [userText, setUserText] = React.useState('');
  const [groupText, setGroupText] = React.useState('');
  const [groupFound, setGroupFound] = React.useState<Array<{ id: string; displayName: string; mail: string; kind: 'm365' | 'security' }> | undefined>();
  const [groupPick, setGroupPick] = React.useState<string | undefined>();
  const [manualType, setManualType] = React.useState<'security' | 'm365'>('security');
  const [siteFilter, setSiteFilter] = React.useState('');
  const [includeLibraries, setIncludeLibraries] = React.useState(true);
  const [speed, setSpeed] = React.useState<ScanSpeed>('gentle');

  const saved = React.useMemo(() => loadJob(), []);
  const [job, setJob] = React.useState<IScanJob | undefined>(saved && saved.job);
  const sitesRef = React.useRef<SiteRef[]>(saved ? saved.sites : []);
  const [, setTick] = React.useState(0);
  const [running, setRunning] = React.useState(false);
  const [busy, setBusy] = React.useState<string | undefined>();
  const [message, setMessage] = React.useState<{ kind: MessageBarType; text: string } | undefined>();
  const [storageOk, setStorageOk] = React.useState(true);
  const [show, setShow] = React.useState<'all' | 'Site' | 'Library'>('all');
  const [query, setQuery] = React.useState('');
  const runRef = React.useRef<{ stop: boolean; reason?: string }>({ stop: true });
  const jobRef = React.useRef<IScanJob | undefined>(job);
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
      runRef.current.stop = true;
      onUnload();
      window.removeEventListener('beforeunload', onUnload);
    };
  }, []);

  const resolvePrincipal = async (): Promise<IScanPrincipal | undefined> => {
    if (kind === 'user') {
      const v = userText.trim();
      if (!v) {
        return undefined;
      }
      if (usersOk) {
        const u = await ctx.graph.findUser(v).catch(() => undefined);
        if (!u) {
          setMessage({ kind: MessageBarType.error, text: `${v} was not found in the directory.` });
          return undefined;
        }
        return { kind: 'user', key: u.userPrincipalName || u.mail || v, label: `${u.displayName} <${u.mail || u.userPrincipalName}>` };
      }
      return { kind: 'user', key: v, label: v };
    }
    if (groupsOk && groupFound) {
      const g = groupFound.filter((x) => x.id === groupPick)[0];
      return g ? { kind: g.kind, key: g.id, label: `${g.displayName}${g.mail ? ` <${g.mail}>` : ''} (${g.kind === 'm365' ? 'Microsoft 365 group' : 'security group'})` } : undefined;
    }
    const id = groupText.trim().toLowerCase();
    if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/.test(id)) {
      setMessage({ kind: MessageBarType.error, text: 'Enter the group’s object id (a GUID from Entra ID).' });
      return undefined;
    }
    return { kind: manualType, key: id, label: `${id} (${manualType === 'm365' ? 'Microsoft 365 group' : 'security group'})` };
  };

  const findGroups = async (): Promise<void> => {
    setBusy('Searching groups…');
    try {
      const found = await ctx.graph.searchAnyGroups(groupText);
      setGroupFound(found);
      setGroupPick(found[0] ? found[0].id : undefined);
    } catch (e) {
      setMessage({ kind: MessageBarType.error, text: graphErrorText(e) });
    } finally {
      setBusy(undefined);
    }
  };

  /** Runs (or resumes) the saved job until it is finished, stopped or paused by errors. */
  const run = async (j: IScanJob): Promise<void> => {
    const sites = sitesRef.current;
    const token = { stop: false, reason: undefined as string | undefined };
    runRef.current = token;
    setRunning(true);
    setMessage(undefined);
    const queue = remaining(j);
    j.pending = [];
    let failsInARow = 0;
    let lastSave = 0;
    const persist = (force: boolean): void => {
      if (force || Date.now() - lastSave > SAVE_EVERY_MS) {
        lastSave = Date.now();
        const ok = saveJob(j);
        if (mounted.current) {
          setStorageOk(ok);
        }
      }
      if (mounted.current) {
        setTick((t) => t + 1);
      }
    };
    const worker = async (): Promise<void> => {
      while (!token.stop) {
        const i = queue.shift();
        if (i === undefined) {
          return;
        }
        j.pending.push(i);
        if (i >= j.next) {
          j.next = i + 1;
        }
        try {
          const rows = await scanner.scanSite(sites[i], j.principal, j.includeLibraries);
          rows.forEach((r) => j.rows.push(r));
          failsInARow = 0;
        } catch (e) {
          if (e instanceof CancelledError) {
            // Not a site problem: leave it pending so it is done on resume.
            token.stop = true;
            persist(true);
            return;
          }
          const err = scanError(e);
          j.failed.push({ i, error: err });
          // "No access" is about that one site; anything else in a row suggests the connection.
          failsInARow = /cannot open|not found/i.test(err) ? 0 : failsInARow + 1;
          if (failsInARow >= MAX_FAILS_IN_A_ROW) {
            token.stop = true;
            token.reason = `Paused after ${MAX_FAILS_IN_A_ROW} errors in a row (last: ${err}). Check your connection or sign-in, then Resume.`;
          }
        }
        j.pending = j.pending.filter((x) => x !== i);
        persist(false);
        if (j.speed === 'gentle' && !token.stop) {
          await new Promise((r) => setTimeout(r, GENTLE_PAUSE_MS));
        }
      }
    };
    await Promise.all(j.speed === 'gentle' ? [worker()] : [worker(), worker()]);
    j.finished = !queue.length && !j.pending.length;
    persist(true);
    ctx.log('Access report', `${j.principal.label}: ${j.next - j.pending.length} of ${j.total} sites, ${j.rows.length} places with access`, true);
    if (mounted.current) {
      setRunning(false);
      if (token.reason) {
        setMessage({ kind: MessageBarType.warning, text: token.reason });
      } else if (j.finished) {
        setMessage({ kind: MessageBarType.success, text: `Finished: ${j.rows.length.toLocaleString()} places with access across ${j.total.toLocaleString()} sites${j.failed.length ? `; ${j.failed.length} site${j.failed.length === 1 ? '' : 's'} could not be read (Retry failed)` : ''}.` });
      } else {
        setMessage({ kind: MessageBarType.info, text: 'Stopped. Progress is saved: Resume carries on from here, even after closing the page.' });
      }
    }
  };

  const start = async (): Promise<void> => {
    setMessage(undefined);
    setBusy('Preparing…');
    try {
      const principal = await resolvePrincipal();
      if (!principal) {
        return;
      }
      let inv = getInventory();
      if (!inv) {
        setBusy('Loading site collections…');
        const r = await ctx.search.listSites(ctx.homeWebUrl, 100000, (n) => setBusy(`Loading site collections… ${n.toLocaleString()}`));
        inv = { sites: r.sites, truncated: r.truncated, oneDriveHidden: r.oneDriveHidden, at: new Date() };
        setInventory(inv);
      }
      const f = siteFilter.trim().toLowerCase();
      const sites = inv.sites.filter((s) => s.manageable && (!f || (s.title + ' ' + s.url).toLowerCase().indexOf(f) >= 0)).map((s) => ({ url: s.url, title: s.title }));
      if (!sites.length) {
        setMessage({ kind: MessageBarType.warning, text: 'No sites match.' });
        return;
      }
      const est = sites.length * secondsPerSite(principal.kind, includeLibraries, speed);
      const ok = await ctx.confirm({
        title: 'Scan for access',
        message:
          `Check ${sites.length.toLocaleString()} sites for ${principal.label}${includeLibraries ? ', including their libraries' : ''}. ` +
          (principal.kind === 'user'
            ? `About ${includeLibraries ? 2 : 1} request${includeLibraries ? 's' : ''} per site, plus one per library with its own permissions. `
            : 'One request per site; a few more only where the group is found. ') +
          `${speed === 'gentle' ? 'Gentle: one site at a time with a short pause.' : 'Normal: two sites at a time.'} Roughly ${duration(est)}. ` +
          'Progress is saved every few seconds, so you can stop, close the page and resume later.',
        confirmText: 'Start scan'
      });
      if (!ok) {
        return;
      }
      const j: IScanJob = { v: 1, id: String(Date.now()), principal, includeLibraries, speed, next: 0, pending: [], rows: [], failed: [], started: Date.now(), saved: Date.now(), finished: false, total: sites.length };
      sitesRef.current = sites;
      setStorageOk(saveJob(j, sites));
      setJob(j);
      setBusy(undefined);
      await run(j);
    } catch (e) {
      setMessage({ kind: MessageBarType.error, text: graphErrorText(e) });
    } finally {
      setBusy(undefined);
    }
  };

  const retryFailed = async (): Promise<void> => {
    if (!job || !job.failed.length) {
      return;
    }
    job.pending = job.pending.concat(job.failed.map((f) => f.i));
    job.failed = [];
    job.finished = false;
    await run(job);
  };

  const discard = async (): Promise<void> => {
    const ok = await ctx.confirm({ title: 'Discard the scan', message: 'Remove the saved progress and results of this scan from this browser? Export first if you need them.', confirmText: 'Discard', danger: true });
    if (ok) {
      clearJob();
      setJob(undefined);
      sitesRef.current = [];
      setMessage(undefined);
    }
  };

  const exportCsv = (): void => {
    if (!job) {
      return;
    }
    const who = (job.principal.key.split('@')[0] || 'access').replace(/[^A-Za-z0-9_-]+/g, '-');
    downloadCsv(`access-${who}${job.finished ? '' : '-partial'}.csv`, ACCESS_HEADERS, accessRows(job.principal.label, job.rows));
  };

  const exportFailed = (): void => {
    if (!job) {
      return;
    }
    downloadCsv('access-scan-failed-sites.csv', ['Site URL', 'Site title', 'Error'], job.failed.map((f) => [sitesRef.current[f.i] ? sitesRef.current[f.i].url : '', sitesRef.current[f.i] ? sitesRef.current[f.i].title : '', f.error]));
  };

  const rowsAll = job ? job.rows : [];
  const rows = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    return rowsAll.filter((r) => (show === 'all' || r.object === show) && (!q || (r.siteTitle + ' ' + r.siteUrl + ' ' + r.title + ' ' + r.how).toLowerCase().indexOf(q) >= 0)).slice(0, 5000);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rowsAll, rowsAll.length, show, query]);

  const processed = job ? Math.max(0, job.next - job.pending.length) : 0;
  const columns: IColumn[] = [
    {
      key: 's',
      name: 'Site',
      minWidth: 200,
      isResizable: true,
      onRender: (r: IAccessRow) => (
        <span style={{ minWidth: 0 }}>
          <a href={r.siteUrl} target="_blank" rel="noopener noreferrer" className={styles.link}>
            {r.siteTitle}
          </a>
          <div className={styles.muted} style={{ fontSize: 12, overflow: 'hidden', textOverflow: 'ellipsis' }}>
            {r.siteUrl}
          </div>
        </span>
      )
    },
    {
      key: 'o',
      name: 'Where',
      minWidth: 170,
      isResizable: true,
      onRender: (r: IAccessRow) =>
        r.object === 'Site' ? (
          <Pill kind="info">Whole site</Pill>
        ) : (
          <span>
            <a href={r.url} target="_blank" rel="noopener noreferrer" className={styles.link}>
              {r.title}
            </a>{' '}
            {r.unique && <Pill kind="warning">Own permissions</Pill>}
          </span>
        )
    },
    { key: 'l', name: 'Access', minWidth: 120, isResizable: true, onRender: (r: IAccessRow) => r.level },
    { key: 'h', name: 'How', minWidth: 180, isResizable: true, onRender: (r: IAccessRow) => r.how }
  ];

  const canEdit = !running && !busy;
  return (
    <div className={styles.view}>
      <ViewHeader title="Access report" hint="Every site and library a person or a group can reach, with the access level and how they get it. Saves its progress so it can resume." />

      {job && !running && !job.finished && (
        <MessageBar
          messageBarType={MessageBarType.warning}
          isMultiline
          actions={
            <div className={styles.actions}>
              <PrimaryButton onClick={() => run(job)}>Resume</PrimaryButton>
              <DefaultButton onClick={discard}>Discard</DefaultButton>
            </div>
          }
        >
          Unfinished scan for <strong>{job.principal.label}</strong>: {processed.toLocaleString()} of {job.total.toLocaleString()} sites done, last saved {formatDateTime(new Date(job.saved))}. Resume carries on from where it stopped.
        </MessageBar>
      )}

      <Card title="Who and where">
        <ChoiceGroup
          selectedKey={kind}
          onChange={(_, o) => o && setKind(o.key as 'user' | 'group')}
          options={[
            { key: 'user', text: 'A person' },
            { key: 'group', text: 'A group (security group or Microsoft 365 group)' }
          ]}
          disabled={!canEdit}
          styles={{ flexContainer: { display: 'flex', gap: 16, flexWrap: 'wrap' } }}
        />
        {kind === 'user' ? (
          <TextField label="Person (e-mail or user name)" placeholder="name@contoso.com" value={userText} onChange={(_, v) => setUserText(v || '')} disabled={!canEdit} styles={{ root: { maxWidth: 420 } }} />
        ) : groupsOk ? (
          <div className={styles.actions} style={{ alignItems: 'flex-end' }}>
            <TextField label="Group name starts with" value={groupText} onChange={(_, v) => setGroupText(v || '')} onKeyDown={(e) => e.key === 'Enter' && groupText.trim() && findGroups()} disabled={!canEdit} styles={{ root: { width: 300 } }} />
            <DefaultButton onClick={findGroups} disabled={!canEdit || !groupText.trim()}>
              Find
            </DefaultButton>
            {groupFound && (
              <Dropdown
                ariaLabel="Group"
                placeholder={groupFound.length ? 'Choose a group' : 'No group starts with that'}
                selectedKey={groupPick || null}
                options={groupFound.map((g) => ({ key: g.id, text: `${g.displayName} (${g.kind === 'm365' ? 'Microsoft 365' : 'security'})` }))}
                onChange={(_, o) => o && setGroupPick(String(o.key))}
                disabled={!canEdit || !groupFound.length}
                styles={{ root: { width: 340 } }}
              />
            )}
          </div>
        ) : (
          <div className={styles.actions} style={{ alignItems: 'flex-end' }}>
            <TextField label="Group object id (from Entra ID)" placeholder="00000000-0000-0000-0000-000000000000" value={groupText} onChange={(_, v) => setGroupText(v || '')} disabled={!canEdit} styles={{ root: { width: 340 } }} />
            <Dropdown
              label="Type"
              selectedKey={manualType}
              options={[
                { key: 'security', text: 'Security group' },
                { key: 'm365', text: 'Microsoft 365 group' }
              ]}
              onChange={(_, o) => o && setManualType(o.key as 'security' | 'm365')}
              disabled={!canEdit}
              styles={{ root: { width: 200 } }}
            />
          </div>
        )}
        <div className={styles.actions} style={{ marginTop: 10, alignItems: 'flex-end' }}>
          <TextField label="Only sites whose title or address contains (optional)" value={siteFilter} onChange={(_, v) => setSiteFilter(v || '')} disabled={!canEdit} styles={{ root: { width: 340 } }} />
          <Checkbox label="Include libraries" checked={includeLibraries} onChange={(_, v) => setIncludeLibraries(!!v)} disabled={!canEdit} />
        </div>
        <ChoiceGroup
          label="Speed"
          selectedKey={speed}
          onChange={(_, o) => o && setSpeed(o.key as ScanSpeed)}
          options={[
            { key: 'gentle', text: 'Gentle (recommended): one site at a time, with a short pause' },
            { key: 'normal', text: 'Normal: two sites at a time' }
          ]}
          disabled={!canEdit}
        />
        <div className={styles.actions} style={{ marginTop: 10 }}>
          <PrimaryButton iconProps={{ iconName: 'Search' }} onClick={start} disabled={!canEdit || (kind === 'user' ? !userText.trim() : groupsOk ? !groupPick : !groupText.trim())}>
            {busy || 'Start scan'}
          </PrimaryButton>
        </div>
        <p className={styles.muted} style={{ marginBottom: 0, fontSize: 12 }}>
          A person&rsquo;s access is SharePoint&rsquo;s own effective permission, so it counts direct grants, SharePoint groups, AD and Microsoft 365 groups, &ldquo;Everyone&rdquo; and site admin rights. A group&rsquo;s access counts grants to the group and to SharePoint groups it is in. You need to be able to open the sites you scan (site collection administrator); sites you cannot open are listed as failed. Starting a new scan replaces the saved one.
        </p>
        {!storageOk && <MessageBar messageBarType={MessageBarType.warning}>This browser would not save the progress (private window or storage full): keep this page open until the scan finishes.</MessageBar>}
      </Card>

      {message && (
        <MessageBar messageBarType={message.kind} isMultiline onDismiss={() => setMessage(undefined)}>
          {message.text}
        </MessageBar>
      )}

      {job && (
        <>
          <Card
            title={`Scan for ${job.principal.label}`}
            right={
              <div className={styles.actions}>
                {running ? (
                  <DefaultButton
                    iconProps={{ iconName: 'Pause' }}
                    onClick={() => {
                      runRef.current.stop = true;
                    }}
                  >
                    Stop (resume later)
                  </DefaultButton>
                ) : (
                  <>
                    {!job.finished && <PrimaryButton iconProps={{ iconName: 'Play' }} onClick={() => run(job)}>Resume</PrimaryButton>}
                    {job.failed.length > 0 && <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={retryFailed}>Retry {job.failed.length} failed</DefaultButton>}
                    <DefaultButton onClick={discard}>Discard</DefaultButton>
                  </>
                )}
              </div>
            }
          >
            <ProgressIndicator
              label={`${processed.toLocaleString()} of ${job.total.toLocaleString()} sites checked${running ? '' : job.finished ? ' (finished)' : ' (stopped)'}`}
              description={`Started ${formatDateTime(new Date(job.started))} · saved ${formatDateTime(new Date(job.saved))} · ${job.speed === 'gentle' ? 'gentle' : 'normal'} speed${job.includeLibraries ? ' · with libraries' : ''}`}
              percentComplete={job.total ? processed / job.total : undefined}
            />
            <div className={styles.statGrid} style={{ marginTop: 10 }}>
              <StatTile icon="Globe" label="Sites with access" value={rowsAll.filter((r) => r.object === 'Site').length.toLocaleString()} tone="brand" onClick={() => setShow('Site')} />
              <StatTile icon="Library" label="Libraries with access" value={rowsAll.filter((r) => r.object === 'Library').length.toLocaleString()} onClick={() => setShow('Library')} />
              <StatTile icon="Permissions" label="Libraries with own permissions" value={rowsAll.filter((r) => r.object === 'Library' && r.unique).length.toLocaleString()} />
              <StatTile icon="Warning" label="Sites not read" value={job.failed.length.toLocaleString()} tone={job.failed.length ? 'warning' : 'good'} onClick={job.failed.length ? exportFailed : undefined} sub={job.failed.length ? 'click to export the list' : undefined} />
            </div>
          </Card>

          <Card
            title={`${rows.length.toLocaleString()} places with access`}
            right={
              <div className={styles.actions}>
                <Dropdown
                  ariaLabel="Show"
                  selectedKey={show}
                  onChange={(_, o) => o && setShow(o.key as 'all' | 'Site' | 'Library')}
                  options={[
                    { key: 'all', text: 'Sites and libraries' },
                    { key: 'Site', text: 'Sites' },
                    { key: 'Library', text: 'Libraries' }
                  ]}
                  styles={{ root: { width: 180 } }}
                />
                <DefaultButton iconProps={{ iconName: 'Download' }} onClick={exportCsv} disabled={!rowsAll.length}>
                  Export CSV{job.finished ? '' : ' (so far)'}
                </DefaultButton>
              </div>
            }
          >
            <SearchBox placeholder="Filter by site, library or how" value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { maxWidth: 340, marginBottom: 8 } }} />
            {rows.length === 0 ? <Empty text={running ? 'Nothing found yet.' : 'No access found.'} /> : <DetailsList items={rows} columns={columns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} />}
            {rowsAll.length > 5000 && <p className={styles.muted}>Showing 5,000; the export has all {rowsAll.length.toLocaleString()}.</p>}
          </Card>
        </>
      )}
    </div>
  );
};
