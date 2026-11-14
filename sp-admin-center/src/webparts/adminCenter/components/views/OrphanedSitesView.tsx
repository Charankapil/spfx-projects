import * as React from 'react';
import { Checkbox, DefaultButton, DetailsList, DetailsListLayoutMode, Dropdown, IColumn, MessageBar, MessageBarType, PrimaryButton, ProgressIndicator, SearchBox, SelectionMode } from '@fluentui/react';
import { ISiteRow } from '../../models';
import { downloadCsv } from '../../services/exportCsv';
import { hasScope, IDirectoryObject, IUserStatus } from '../../services/GraphApi';
import { OwnerState, ownerState, personKey } from '../../services/m365Review';
import { peopleText } from '../../services/siteGroupsExport';
import { siteType } from '../../services/siteTypes';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { graphErrorText, NeedsApproval, useGraphScopes } from '../shared/graph';
import { getInventory, setInventory } from '../shared/inventory';
import { Card, Empty, Pill, StatTile, ViewHeader } from '../shared/ui';

interface IOrphanRow {
  site: ISiteRow;
  source: 'Microsoft 365 group' | 'SharePoint Owners group';
  state: OwnerState | 'Not read';
  active: IDirectoryObject[];
  left: IDirectoryObject[];
  note?: string;
}

type Show = 'problems' | OwnerState | 'all';

/**
 * Sites with nobody left in charge: no owners, or every owner's account deleted or disabled.
 * Owners are the Microsoft 365 group's owners (20 groups per Graph request) and, optionally,
 * the people in other sites' SharePoint Owners group (one SharePoint request per site).
 * Accounts are checked 20 per Graph request (User.Read.All).
 */
export const OrphanedSitesView: React.FC = () => {
  const ctx = useAdmin();
  const access = useGraphScopes();
  const graphOk = access.scopes ? hasScope(access.scopes, 'Group.Read.All') && hasScope(access.scopes, 'User.Read.All') : !!access.error;
  const [includeSp, setIncludeSp] = React.useState(false);
  const [rows, setRows] = React.useState<IOrphanRow[] | undefined>();
  const [progress, setProgress] = React.useState<{ text: string; done: number; total: number } | undefined>();
  const [error, setError] = React.useState<string | undefined>();
  const [show, setShow] = React.useState<Show>('problems');
  const [query, setQuery] = React.useState('');
  const runRef = React.useRef<{ stop: boolean }>({ stop: false });
  const mounted = React.useRef(true);
  React.useEffect(
    () => () => {
      mounted.current = false;
      runRef.current.stop = true;
    },
    []
  );

  const scan = async (): Promise<void> => {
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
      let inv = getInventory();
      if (!inv) {
        say('Loading site collections…');
        const r = await ctx.search.listSites(ctx.homeWebUrl, 100000, (n) => say(`Loading site collections… ${n.toLocaleString()}`));
        inv = { sites: r.sites, truncated: r.truncated, oneDriveHidden: r.oneDriveHidden, at: new Date() };
        setInventory(inv);
      }
      const groupSites = inv.sites.filter((s) => !!s.groupId);
      const spSites = includeSp ? inv.sites.filter((s) => !s.groupId && !s.relatedGroupId && s.manageable) : [];
      const minutes = Math.max(1, Math.round((groupSites.length / 20 + spSites.length / 4) / 60));
      const ok = await ctx.confirm({
        title: 'Find orphaned sites',
        message:
          `Check the owners of ${groupSites.length.toLocaleString()} group-connected sites (about ${Math.ceil(groupSites.length / 20).toLocaleString()} batched Microsoft Graph requests)` +
          (spSites.length ? ` and the SharePoint Owners group of ${spSites.length.toLocaleString()} other sites (${spSites.length.toLocaleString()} SharePoint requests, at most 2 at a time)` : '') +
          `, then whether those owners still have an account (20 per request). Roughly ${minutes} minute${minutes === 1 ? '' : 's'}; you can stop at any time.`,
        confirmText: 'Start'
      });
      if (!ok) {
        say('');
        setProgress(undefined);
        return;
      }
      // 1. Owners of group-connected sites.
      const groupOwners = await ctx.graph.ownersMany(
        groupSites.map((s) => s.groupId as string),
        (d, t) => say('Reading group owners…', d, t),
        () => run.stop
      );
      // 2. Owners group of other sites (two at a time, like everything else).
      const spOwners: { [url: string]: { owners: IDirectoryObject[]; error?: string } } = {};
      const queue = spSites.slice();
      let spDone = 0;
      const worker = async (): Promise<void> => {
        for (let s = queue.shift(); s && !run.stop; s = queue.shift()) {
          const a = await ctx.api.siteAccess(s.url, false);
          spOwners[s.url] = { owners: a.owners, error: a.error };
          spDone++;
          say('Reading SharePoint Owners groups…', spDone, spSites.length);
        }
      };
      await Promise.all([worker(), worker()]);
      // 3. Do the owners still have an account?
      const keys: string[] = [];
      Object.keys(groupOwners).forEach((g) => (groupOwners[g] || []).filter((p) => p.type === 'user').forEach((p) => keys.push(personKey(p, true))));
      Object.keys(spOwners).forEach((u) => spOwners[u].owners.filter((p) => p.type === 'user').forEach((p) => keys.push(personKey(p, false))));
      const status = await ctx.graph.userStatus(keys.filter((k) => !!k), (d, t) => say('Checking owner accounts…', d, t), () => run.stop);
      const lookup = (byId: boolean) => (p: IDirectoryObject): IUserStatus | undefined => status[personKey(p, byId)];

      const out: IOrphanRow[] = [];
      groupSites.forEach((s) => {
        const owners = groupOwners[(s.groupId as string).toLowerCase()];
        if (!owners) {
          out.push({ site: s, source: 'Microsoft 365 group', state: 'Not read', active: [], left: [], note: run.stop ? 'Stopped before this site' : 'Could not read the group' });
          return;
        }
        const r = ownerState(owners, lookup(true));
        out.push({ site: s, source: 'Microsoft 365 group', ...r });
      });
      spSites.forEach((s) => {
        const o = spOwners[s.url];
        if (!o || o.error) {
          out.push({ site: s, source: 'SharePoint Owners group', state: 'Not read', active: [], left: [], note: o ? o.error : 'Stopped before this site' });
          return;
        }
        const r = ownerState(o.owners, lookup(false));
        out.push({ site: s, source: 'SharePoint Owners group', ...r });
      });
      if (mounted.current) {
        setRows(out);
      }
      ctx.log('Find orphaned sites', `${out.length} sites checked`, true);
    } catch (e) {
      if (mounted.current) {
        setError(graphErrorText(e));
      }
    } finally {
      run.stop = true;
      if (mounted.current) {
        setProgress(undefined);
      }
    }
  };

  const all = rows || [];
  const count = (st: OwnerState | 'Not read'): number => all.filter((r) => r.state === st).length;
  const shown = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    return all
      .filter((r) => (show === 'all' ? true : show === 'problems' ? r.state === 'No owners' || r.state === 'All owners have left' : r.state === show))
      .filter((r) => !q || (r.site.title + ' ' + r.site.url).toLowerCase().indexOf(q) >= 0);
  }, [all, show, query]);

  const columns: IColumn[] = [
    {
      key: 's',
      name: 'Site',
      minWidth: 220,
      isResizable: true,
      onRender: (r: IOrphanRow) => (
        <span style={{ minWidth: 0 }}>
          <a href={r.site.url} target="_blank" rel="noopener noreferrer" className={styles.link}>
            {r.site.title}
          </a>
          <div className={styles.muted} style={{ fontSize: 12, overflow: 'hidden', textOverflow: 'ellipsis' }}>
            {r.site.url}
          </div>
        </span>
      )
    },
    {
      key: 'st',
      name: 'Owners',
      minWidth: 150,
      onRender: (r: IOrphanRow) => (
        <Pill kind={r.state === 'OK' ? 'good' : r.state === 'One active owner' ? 'warning' : r.state === 'Not read' ? undefined : 'critical'}>{r.state}</Pill>
      )
    },
    { key: 'l', name: 'Owners who left', minWidth: 180, isResizable: true, onRender: (r: IOrphanRow) => <span title={peopleText(r.left)}>{r.left.map((p) => p.displayName).join(', ')}</span> },
    { key: 'a', name: 'Active owners', minWidth: 160, isResizable: true, onRender: (r: IOrphanRow) => <span title={peopleText(r.active)}>{r.active.map((p) => p.displayName).join(', ') || r.note || ''}</span> },
    { key: 'src', name: 'Owners from', minWidth: 140, onRender: (r: IOrphanRow) => r.source },
    {
      key: 'm',
      name: '',
      minWidth: 80,
      onRender: (r: IOrphanRow) => (
        <button
          type="button"
          className={styles.link}
          onClick={async () => {
            try {
              const [web, site] = await Promise.all([ctx.api.getWeb(r.site.url), ctx.api.getSite(r.site.url)]);
              ctx.setTarget({ webUrl: web.url, siteUrl: site.url, title: web.title });
              ctx.openView(r.site.groupId ? 'm365groups' : 'people');
            } catch (e) {
              ctx.notify((e as Error).message, 'error');
            }
          }}
        >
          Fix
        </button>
      )
    }
  ];

  return (
    <div className={styles.view}>
      <ViewHeader title="Orphaned sites" hint="Sites with nobody left in charge: no owners, or every owner's account deleted or disabled. Fix opens the site's owners." />
      {access.scopes && !graphOk ? (
        <NeedsApproval scopes={['Group.Read.All', 'User.Read.All']} feature="Orphaned sites" />
      ) : (
        <>
          <Card>
            <p className={styles.muted} style={{ marginTop: 0 }}>
              Uses the site list from the All sites tab (loaded here if needed). Group-connected sites are checked by their Microsoft 365 group&rsquo;s owners; other sites, when included, by the people in their SharePoint Owners group.
            </p>
            <div className={styles.actions} style={{ alignItems: 'center' }}>
              <Checkbox label="Include sites that are not group-connected (1 SharePoint request per site)" checked={includeSp} onChange={(_, v) => setIncludeSp(!!v)} disabled={!!progress} />
              <PrimaryButton iconProps={{ iconName: 'Search' }} onClick={scan} disabled={!!progress}>
                {rows ? 'Scan again' : 'Find orphaned sites'}
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
              {rows && (
                <DefaultButton
                  iconProps={{ iconName: 'Download' }}
                  onClick={() =>
                    downloadCsv(
                      'orphaned-sites.csv',
                      ['Site URL', 'Site title', 'Site type', 'Owners from', 'Owners', 'Owners who left', 'Active owners', 'Note'],
                      shown.map((r) => [r.site.url, r.site.title, siteType(r.site.template, r.site.groupConnected), r.source, r.state, peopleText(r.left), peopleText(r.active), r.note || ''])
                    )
                  }
                >
                  Export shown
                </DefaultButton>
              )}
            </div>
            {progress && <ProgressIndicator label={`${progress.text}${progress.total ? ` ${progress.done.toLocaleString()} of ${progress.total.toLocaleString()}` : ''}`} percentComplete={progress.total ? progress.done / progress.total : undefined} />}
            {error && <MessageBar messageBarType={MessageBarType.error} isMultiline>{error}</MessageBar>}
          </Card>
          {rows && (
            <>
              <div className={styles.statGrid}>
                <StatTile icon="Globe" label="Sites checked" value={all.length.toLocaleString()} tone="brand" onClick={() => setShow('all')} />
                <StatTile icon="Blocked2" label="No owners" value={count('No owners').toLocaleString()} tone={count('No owners') ? 'critical' : 'good'} onClick={() => setShow('No owners')} />
                <StatTile icon="UserRemove" label="All owners have left" value={count('All owners have left').toLocaleString()} tone={count('All owners have left') ? 'critical' : 'good'} onClick={() => setShow('All owners have left')} />
                <StatTile icon="Contact" label="One active owner" value={count('One active owner').toLocaleString()} tone={count('One active owner') ? 'warning' : 'good'} onClick={() => setShow('One active owner')} />
              </div>
              <Card
                title={`${shown.length.toLocaleString()} sites`}
                right={
                  <Dropdown
                    ariaLabel="Show"
                    selectedKey={show}
                    onChange={(_, o) => o && setShow(o.key as Show)}
                    options={[
                      { key: 'problems', text: 'Orphaned' },
                      { key: 'No owners', text: 'No owners' },
                      { key: 'All owners have left', text: 'All owners have left' },
                      { key: 'One active owner', text: 'One active owner' },
                      { key: 'Not read', text: 'Not read' },
                      { key: 'all', text: 'All checked sites' }
                    ]}
                    styles={{ root: { width: 200 } }}
                  />
                }
              >
                <SearchBox placeholder="Filter by title or URL" value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { maxWidth: 320, marginBottom: 8 } }} />
                {shown.length === 0 ? <Empty text="No sites match." /> : <DetailsList items={shown.slice(0, 3000)} columns={columns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} />}
              </Card>
            </>
          )}
        </>
      )}
    </div>
  );
};
