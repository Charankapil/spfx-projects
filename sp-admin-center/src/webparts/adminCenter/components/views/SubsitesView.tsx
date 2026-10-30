import * as React from 'react';
import { DefaultButton, DetailsList, DetailsListLayoutMode, Dropdown, IColumn, MessageBar, MessageBarType, Panel, PanelType, PrimaryButton, ProgressIndicator, SearchBox, SelectionMode } from '@fluentui/react';
import { downloadCsv } from '../../services/exportCsv';
import { formatBytes, formatDate, relativeTime } from '../../services/format';
import { fitForExcel } from '../../services/siteGroupsExport';
import { ISubsiteDetail, SUBSITE_HEADERS, subsiteRows } from '../../services/subsitesExport';
import { IDirectoryObject } from '../../services/GraphApi';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { Card, Empty, Pill, StatTile, ViewHeader } from '../shared/ui';

type Show = 'all' | 'unique' | 'inherits';

/**
 * Every subsite of the selected site collection with its size (SharePoint's storage metrics,
 * one request per subsite), last change, and, for subsites with their own permissions, the
 * people in their Owners / Members / Visitors groups (one more request). Runs on request.
 */
export const SubsitesView: React.FC = () => {
  const ctx = useAdmin();
  const { target } = ctx;
  const [details, setDetails] = React.useState<ISubsiteDetail[] | undefined>();
  const [truncated, setTruncated] = React.useState(false);
  const [progress, setProgress] = React.useState<{ text: string; done: number; total: number } | undefined>();
  const [error, setError] = React.useState<string | undefined>();
  const [show, setShow] = React.useState<Show>('all');
  const [query, setQuery] = React.useState('');
  const [open, setOpen] = React.useState<ISubsiteDetail | undefined>();
  const runRef = React.useRef<{ stop: boolean }>({ stop: false });
  const mounted = React.useRef(true);
  React.useEffect(
    () => () => {
      mounted.current = false;
      runRef.current.stop = true;
    },
    []
  );

  const load = async (): Promise<void> => {
    runRef.current.stop = true;
    const run = { stop: false };
    runRef.current = run;
    setError(undefined);
    const say = (text: string, done = 0, total = 0): void => {
      if (mounted.current) {
        setProgress({ text, done, total });
      }
    };
    try {
      say('Finding subsites…');
      const tree = await ctx.api.subsiteTree(target.siteUrl, 2000, (n) => say(`Finding subsites… ${n.toLocaleString()}`), () => run.stop);
      const out: ISubsiteDetail[] = tree.webs.map((web) => ({ web }));
      if (mounted.current) {
        setDetails(out.slice());
        setTruncated(tree.truncated);
      }
      // Size of every web, owners and members of the ones with their own permissions: two workers.
      const jobs: Array<() => Promise<void>> = [];
      out.forEach((d) => {
        jobs.push(async () => {
          d.size = await ctx.api.webStorage(d.web.url);
        });
        if (d.web.hasUniquePermissions) {
          jobs.push(async () => {
            d.access = await ctx.api.siteAccess(d.web.url, false);
          });
        }
      });
      const total = jobs.length;
      let done = 0;
      const worker = async (): Promise<void> => {
        for (let job = jobs.shift(); job && !run.stop; job = jobs.shift()) {
          await job();
          done++;
          say('Reading sizes and owners…', done, total);
          if (done % 20 === 0 && mounted.current) {
            setDetails(out.slice());
          }
        }
      };
      await Promise.all([worker(), worker()]);
      if (mounted.current) {
        setDetails(out.slice());
      }
      ctx.log('Read subsites', `${target.siteUrl}: ${out.length - 1} subsites`, true);
    } catch (e) {
      if (mounted.current) {
        setError((e as Error).message);
      }
    } finally {
      run.stop = true;
      if (mounted.current) {
        setProgress(undefined);
      }
    }
  };

  const all = details || [];
  const subs = all.filter((d) => d.web.depth > 0);
  const subsBytes = subs.reduce((n, d) => n + (d.size ? d.size.bytes : 0), 0);
  const rows = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    return all
      .filter((d) => (show === 'unique' ? d.web.depth > 0 && d.web.hasUniquePermissions : show === 'inherits' ? d.web.depth > 0 && !d.web.hasUniquePermissions : true))
      .filter((d) => !q || (d.web.title + ' ' + d.web.url).toLowerCase().indexOf(q) >= 0);
  }, [all, show, query]);

  const manage = (d: ISubsiteDetail): void => {
    ctx.setTarget({ webUrl: d.web.url, siteUrl: target.siteUrl, title: d.web.title });
    ctx.openView('people');
  };

  const peopleCell = (d: ISubsiteDetail, kind: 'owners' | 'members'): React.ReactNode => {
    if (d.web.depth > 0 && !d.web.hasUniquePermissions) {
      return <span className={styles.muted}>Inherits</span>;
    }
    if (!d.access) {
      return <span className={styles.muted}>…</span>;
    }
    if (d.access.error) {
      return <span className={styles.muted} title={d.access.error}>Not read</span>;
    }
    const list = d.access[kind];
    return (
      <button type="button" className={styles.link} onClick={() => setOpen(d)} title={list.map((p) => p.displayName).join(', ')}>
        {list.length.toLocaleString()} {kind === 'owners' ? 'owner' : 'member'}
        {list.length === 1 ? '' : 's'}
      </button>
    );
  };

  const columns: IColumn[] = [
    {
      key: 't',
      name: 'Site',
      minWidth: 240,
      isResizable: true,
      onRender: (d: ISubsiteDetail) => (
        <span style={{ minWidth: 0, paddingLeft: Math.min(d.web.depth, 6) * 16, display: 'block' }}>
          <a href={d.web.url} target="_blank" rel="noopener noreferrer" className={styles.link}>
            {d.web.title || d.web.url}
          </a>
          {d.web.depth === 0 && (
            <>
              {' '}
              <Pill kind="info">Top-level site</Pill>
            </>
          )}
          <div className={styles.muted} style={{ fontSize: 12, overflow: 'hidden', textOverflow: 'ellipsis' }} title={d.web.url}>
            {d.web.serverRelativeUrl}
          </div>
        </span>
      )
    },
    { key: 's', name: 'Size', minWidth: 90, onRender: (d: ISubsiteDetail) => (d.size ? formatBytes(d.size.bytes) : progress ? '…' : '—') },
    { key: 'f', name: 'Files', minWidth: 70, onRender: (d: ISubsiteDetail) => (d.size ? d.size.files.toLocaleString() : '') },
    { key: 'm', name: 'Last modified', minWidth: 120, onRender: (d: ISubsiteDetail) => <span title={formatDate(d.web.lastModified)}>{relativeTime(d.web.lastModified)}</span> },
    {
      key: 'p',
      name: 'Permissions',
      minWidth: 110,
      onRender: (d: ISubsiteDetail) => (d.web.depth === 0 ? '' : d.web.hasUniquePermissions ? <Pill kind="warning">Unique</Pill> : <Pill>Inherits</Pill>)
    },
    { key: 'o', name: 'Owners', minWidth: 90, onRender: (d: ISubsiteDetail) => peopleCell(d, 'owners') },
    { key: 'mm', name: 'Members', minWidth: 90, onRender: (d: ISubsiteDetail) => peopleCell(d, 'members') },
    { key: 'tp', name: 'Template', minWidth: 90, onRender: (d: ISubsiteDetail) => d.web.template },
    {
      key: 'a',
      name: '',
      minWidth: 70,
      onRender: (d: ISubsiteDetail) =>
        d.web.depth > 0 ? (
          <button type="button" className={styles.link} onClick={() => manage(d)}>
            Manage
          </button>
        ) : null
    }
  ];

  return (
    <div className={styles.view}>
      <ViewHeader title="Subsites" hint={`Every subsite of ${target.siteUrl}, with its size, last change and, where it has its own permissions, its owners and members.`} />
      <Card>
        <p className={styles.muted} style={{ marginTop: 0 }}>
          Size is SharePoint&rsquo;s stored storage figure for each subsite&rsquo;s own content (libraries and lists with version history; not its own subsites), so nothing walks your files. Requests: one per subsite to find the tree, one for its size, and one more for owners and members when it has unique permissions; at most 2 at a time.
        </p>
        <div className={styles.actions}>
          <PrimaryButton iconProps={{ iconName: 'SiteScan' }} onClick={load} disabled={!!progress}>
            {details ? 'Reload' : 'Load subsites'}
          </PrimaryButton>
          {progress && (
            <DefaultButton
              iconProps={{ iconName: 'Stop' }}
              onClick={() => {
                runRef.current.stop = true;
              }}
            >
              Stop
            </DefaultButton>
          )}
          {details && (
            <DefaultButton
              iconProps={{ iconName: 'Download' }}
              disabled={!!progress}
              onClick={() => {
                const fitted = fitForExcel(SUBSITE_HEADERS, subsiteRows(rows));
                downloadCsv(`subsites-${(target.siteUrl.split('/').pop() || 'site').replace(/[^A-Za-z0-9_-]+/g, '-')}.csv`, fitted.headers, fitted.rows);
              }}
            >
              Export to CSV
            </DefaultButton>
          )}
        </div>
        {progress && <ProgressIndicator label={`${progress.text}${progress.total ? ` ${progress.done.toLocaleString()} of ${progress.total.toLocaleString()}` : ''}`} percentComplete={progress.total ? progress.done / progress.total : undefined} />}
        {error && <MessageBar messageBarType={MessageBarType.error} isMultiline>{error}</MessageBar>}
        {truncated && <MessageBar messageBarType={MessageBarType.warning}>Stopped at 2,000 subsites.</MessageBar>}
      </Card>

      {details && (
        <>
          <div className={styles.statGrid}>
            <StatTile icon="Globe" label="Subsites" value={subs.length.toLocaleString()} sub={subs.length ? `${Math.max(...subs.map((d) => d.web.depth))} level${Math.max(...subs.map((d) => d.web.depth)) === 1 ? '' : 's'} deep` : undefined} tone="brand" onClick={() => setShow('all')} />
            <StatTile icon="Permissions" label="With unique permissions" value={subs.filter((d) => d.web.hasUniquePermissions).length.toLocaleString()} onClick={() => setShow('unique')} />
            <StatTile icon="Database" label="Size of all subsites" value={formatBytes(subsBytes)} sub={all[0] && all[0].size ? `top-level site: ${formatBytes(all[0].size.bytes)}` : undefined} />
            <StatTile icon="Page" label="Files in subsites" value={subs.reduce((n, d) => n + (d.size ? d.size.files : 0), 0).toLocaleString()} sub={all[0] && all[0].size ? `top-level site: ${all[0].size.files.toLocaleString()}` : undefined} />
          </div>
          <Card
            title={`${rows.length.toLocaleString()} sites`}
            right={
              <Dropdown
                ariaLabel="Show"
                selectedKey={show}
                onChange={(_, o) => o && setShow(o.key as Show)}
                options={[
                  { key: 'all', text: 'All' },
                  { key: 'unique', text: 'Unique permissions' },
                  { key: 'inherits', text: 'Inherit permissions' }
                ]}
                styles={{ root: { width: 190 } }}
              />
            }
          >
            <SearchBox placeholder="Filter by title or URL" value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { maxWidth: 320, marginBottom: 8 } }} />
            {subs.length === 0 && !progress ? <Empty text="This site collection has no subsites." /> : null}
            {rows.length > 0 && <DetailsList items={rows} columns={columns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} />}
          </Card>
        </>
      )}

      <Panel isOpen={!!open} onDismiss={() => setOpen(undefined)} type={PanelType.medium} headerText={open ? open.web.title : ''} closeButtonAriaLabel="Close" isLightDismiss>
        {open && open.access && (
          <>
            <p className={styles.muted} style={{ marginTop: 0, wordBreak: 'break-all' }}>
              {open.web.url}
            </p>
            <People title={`Owners${open.access.ownerGroup ? ` · ${open.access.ownerGroup}` : ''}`} people={open.access.owners} />
            <People title={`Members${open.access.memberGroup ? ` · ${open.access.memberGroup}` : ''}`} people={open.access.members} />
            <People title={`Visitors${open.access.visitorGroup ? ` · ${open.access.visitorGroup}` : ''}`} people={open.access.visitors} />
          </>
        )}
      </Panel>
    </div>
  );
};

const People: React.FC<{ title: string; people: IDirectoryObject[] }> = ({ title, people }) => (
  <>
    <h3 style={{ fontSize: 14, margin: '16px 0 6px' }}>
      {title} ({people.length})
    </h3>
    {people.length === 0 ? (
      <Empty text="None." />
    ) : (
      <ul className={styles.memberList}>
        {people.map((p) => (
          <li key={p.id} className={styles.memberRow}>
            <span style={{ minWidth: 0 }}>
              <div>
                {p.displayName} {(p.userType || '').toLowerCase() === 'guest' && <Pill kind="warning">Guest</Pill>} {p.type === 'group' && <Pill kind="info">Group</Pill>}
              </div>
              <div className={styles.muted} style={{ fontSize: 12 }}>
                {p.mail || p.userPrincipalName}
              </div>
            </span>
          </li>
        ))}
      </ul>
    )}
  </>
);
