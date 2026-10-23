import * as React from 'react';
import { Checkbox, ChoiceGroup, DefaultButton, TextField, DetailsList, DetailsListLayoutMode, Dropdown, IColumn, MessageBar, MessageBarType, Panel, PanelType, PrimaryButton, ProgressIndicator, SearchBox, SelectionMode } from '@fluentui/react';
import { ISiteRow } from '../../models';
import { downloadCsv } from '../../services/exportCsv';
import { formatDate } from '../../services/format';
import { hasScope, IDirectoryObject, IGroupMembership, IGroupSummary } from '../../services/GraphApi';
import { ListExport, listItemsFor } from '../../services/ListExport';
import { ExportLayout, fitForExcel, fromGroupMembership, ISiteAccess, PERSON_HEADERS, SITE_HEADERS, siteGroupRows } from '../../services/siteGroupsExport';
import { siteType } from '../../services/siteTypes';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { graphErrorText, NeedsApproval, useGraphScopes } from '../shared/graph';
import { Card, Empty, Pill, StatTile, ViewHeader } from '../shared/ui';

type Show = 'all' | 'group' | 'nogroup' | 'channel';

const GROUP_SCOPE = 'Group.Read.All';

/** Loaded once per page visit; kept while you move between views. */
let sessionSites: { sites: ISiteRow[]; truncated: boolean; oneDriveHidden: number; at: Date } | undefined;
let sessionGroups: { [id: string]: IGroupSummary } | undefined;

/**
 * Every active site collection (from search) and whether it is connected to a Microsoft 365
 * group. Owners and members of a site's group open next to it (Microsoft Graph), and the
 * export lists every site with its group's owners and members.
 */
export const SitesGroupsView: React.FC = () => {
  const ctx = useAdmin();
  const access = useGraphScopes();
  // If the token could not be read, let the feature try: its own errors explain what is wrong.
  const graphOk = access.scopes ? hasScope(access.scopes, GROUP_SCOPE) : !!access.error;
  const [data, setData] = React.useState(sessionSites);
  const [groups, setGroups] = React.useState(sessionGroups);
  const [loading, setLoading] = React.useState<string | undefined>();
  const [error, setError] = React.useState<string | undefined>();
  const [show, setShow] = React.useState<Show>('all');
  const [query, setQuery] = React.useState('');
  const [open, setOpen] = React.useState<ISiteRow | undefined>();
  const [layout, setLayout] = React.useState<ExportLayout>('site');
  const [includeAdmins, setIncludeAdmins] = React.useState(false);
  const [saveTo, setSaveTo] = React.useState<'csv' | 'list'>('csv');
  const [listSite, setListSite] = React.useState(ctx.homeWebUrl);
  const [listName, setListName] = React.useState(`Site owners and members ${new Date().toISOString().substring(0, 10)}`);
  const [savedList, setSavedList] = React.useState<{ url: string; title: string; added: number; failed: string[] } | undefined>();
  const [exporting, setExporting] = React.useState<{ done: number; total: number; phase?: string } | undefined>();
  const stopRef = React.useRef(false);
  const mounted = React.useRef(true);
  React.useEffect(
    () => () => {
      mounted.current = false;
      stopRef.current = true;
    },
    []
  );

  const load = async (): Promise<void> => {
    setError(undefined);
    setLoading('Loading site collections…');
    try {
      const r = await ctx.search.listSites(ctx.homeWebUrl, 100000, (n) => mounted.current && setLoading(`Loading site collections… ${n.toLocaleString()}`));
      sessionSites = { sites: r.sites, truncated: r.truncated, oneDriveHidden: r.oneDriveHidden, at: new Date() };
      if (mounted.current) {
        setData(sessionSites);
      }
      if (graphOk) {
        setLoading('Reading Microsoft 365 group names…');
        try {
          sessionGroups = await ctx.graph.groupDirectory();
          if (mounted.current) {
            setGroups(sessionGroups);
          }
        } catch (e) {
          if (mounted.current) {
            setError('Group names could not be read: ' + graphErrorText(e));
          }
        }
      }
    } catch (e) {
      if (mounted.current) {
        setError((e as Error).message);
      }
    } finally {
      if (mounted.current) {
        setLoading(undefined);
      }
    }
  };

  // Group names arrive later when the permission check finishes after the sites were loaded.
  React.useEffect(() => {
    if (!graphOk || !data || groups || loading) {
      return;
    }
    ctx.graph.groupDirectory().then(
      (g) => {
        sessionGroups = g;
        if (mounted.current) {
          setGroups(g);
        }
      },
      () => undefined
    );
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [graphOk, data]);

  const sites = data ? data.sites : [];
  const counts = React.useMemo(
    () => ({
      group: sites.filter((s) => !!s.groupId).length,
      channel: sites.filter((s) => !s.groupId && !!s.relatedGroupId).length
    }),
    [sites]
  );
  const rows = React.useMemo(() => {
    const q = query.trim().toLowerCase();
    return sites
      .filter((s) => (show === 'group' ? !!s.groupId : show === 'nogroup' ? !s.groupId && !s.relatedGroupId : show === 'channel' ? !s.groupId && !!s.relatedGroupId : true))
      .filter((s) => {
        if (!q) {
          return true;
        }
        const g = s.groupId && groups ? groups[s.groupId] : undefined;
        return (s.title + ' ' + s.url + ' ' + (g ? g.displayName + ' ' + g.mail : '')).toLowerCase().indexOf(q) >= 0;
      });
  }, [sites, show, query, groups]);

  const exportSites = (): void => {
    downloadCsv(
      'site-collections-m365-groups.csv',
      ['Site URL', 'Site title', 'Site type', 'Microsoft 365 group', 'Group id', 'Group name', 'Group e-mail', 'Created', 'Last modified'],
      rows.map((s) => {
        const g = s.groupId && groups ? groups[s.groupId] : undefined;
        return [s.url, s.title, siteType(s.template, s.groupConnected), s.groupId ? 'Yes' : s.relatedGroupId ? 'Teams channel site' : 'No', s.groupId || '', g ? g.displayName : '', g ? g.mail : '', s.created, s.lastModified];
      })
    );
  };

  const exportMembership = async (): Promise<void> => {
    const list = rows;
    if (saveTo === 'list') {
      try {
        ctx.client.assertSameOrigin(listSite.trim());
      } catch {
        ctx.notify('Enter the address of a SharePoint site on this tenant for the list.', 'error');
        return;
      }
      if (!listName.trim()) {
        ctx.notify('Enter a name for the list.', 'error');
        return;
      }
    }
    const groupSites = list.filter((s) => !!s.groupId);
    // Sites on another host (e.g. OneDrive) cannot be read from this page.
    const spSites = list.filter((s) => !s.groupId && s.manageable);
    const ids = graphOk ? groupSites.map((s) => s.groupId as string) : [];
    const adminSites = includeAdmins ? groupSites.filter((s) => s.manageable) : [];
    const graphRequests = Math.ceil(ids.length / 10);
    const spRequests = spSites.length * (includeAdmins ? 2 : 1) + adminSites.length;
    const minutes = Math.max(1, Math.round(Math.max(graphRequests * 1.2, spRequests / 4) / 60));
    const ok = await ctx.confirm({
      title: 'Export owners and members',
      message:
        `Read the owners and members of ${list.length.toLocaleString()} sites: ` +
        (groupSites.length ? `${groupSites.length.toLocaleString()} Microsoft 365 groups ${graphOk ? `(about ${graphRequests.toLocaleString()} batched Microsoft Graph requests)` : '(skipped: the Microsoft Graph permission is not approved)'} and ` : '') +
        `the SharePoint Owners, Members and Visitors groups of ${spSites.length.toLocaleString()} other sites (${spRequests.toLocaleString()} SharePoint requests, at most 2 at a time). ` +
        `Roughly ${minutes} minute${minutes === 1 ? '' : 's'}. ` +
        (saveTo === 'list' ? `Then ${list.length.toLocaleString()} items are added to the list "${listName.trim()}" on ${listSite.trim()} (created if it does not exist), up to 50 per request. ` : '') +
        'You can stop it and export what was read so far.',
      confirmText: 'Start export'
    });
    if (!ok) {
      return;
    }
    stopRef.current = false;
    const total = ids.length + spSites.length + adminSites.length;
    let graphDone = 0;
    let spDone = 0;
    const tick = (): void => {
      if (mounted.current) {
        setExporting({ done: graphDone + spDone, total });
      }
    };
    tick();
    try {
      let names = groups || {};
      if (graphOk && !groups) {
        names = sessionGroups = await ctx.graph.groupDirectory();
        if (mounted.current) {
          setGroups(names);
        }
      }
      const spAccess: { [url: string]: ISiteAccess } = {};
      const admins: { [url: string]: IDirectoryObject[] } = {};
      // SharePoint sites: two workers (SPClient also never has more than 2 requests in flight).
      const queue: Array<() => Promise<void>> = spSites
        .map((s) => async () => {
          spAccess[s.url] = await ctx.api.siteAccess(s.url, includeAdmins);
        })
        .concat(
          adminSites.map((s) => async () => {
            admins[s.url] = await ctx.api.siteAdmins(s.url).catch(() => []);
          })
        );
      const worker = async (): Promise<void> => {
        for (let job = queue.shift(); job && !stopRef.current; job = queue.shift()) {
          await job();
          spDone++;
          tick();
        }
      };
      const [mem] = await Promise.all([
        ids.length
          ? ctx.graph.membershipMany(
              ids,
              (d) => {
                graphDone = d;
                tick();
              },
              () => stopRef.current
            )
          : Promise.resolve({} as { [id: string]: IGroupMembership }),
        worker(),
        worker()
      ]);
      if (!mounted.current) {
        return;
      }
      const accessOf = (s: ISiteRow): ISiteAccess | undefined => {
        if (s.groupId) {
          if (!graphOk) {
            return { source: 'm365', owners: [], members: [], visitors: [], truncated: false, error: 'Microsoft 365 group members need the Group.Read.All Graph permission' };
          }
          const m = mem[s.groupId];
          return m ? { ...fromGroupMembership(m), admins: admins[s.url] } : undefined;
        }
        if (!s.manageable) {
          return { source: 'sharepoint', owners: [], members: [], visitors: [], truncated: false, error: 'On another host; open that site to read it' };
        }
        return spAccess[s.url];
      };
      const read = Object.keys(mem).length + Object.keys(spAccess).length;
      const stoppedEarly = stopRef.current;
      if (saveTo === 'csv') {
        const stamp = new Date().toISOString().substring(0, 10);
        // Cells longer than Excel allows are spread over extra "(part 2)" columns instead of breaking the row.
        const fitted = fitForExcel(layout === 'site' ? SITE_HEADERS : PERSON_HEADERS, siteGroupRows(list, names, accessOf, layout));
        downloadCsv(layout === 'site' ? `sites-owners-members-${stamp}.csv` : `sites-owners-members-people-${stamp}.csv`, fitted.headers, fitted.rows);
        ctx.log('Export sites with owners and members', `${list.length} sites, ${read} read`, true);
        ctx.notify(
          stoppedEarly ? `Stopped: exported with ${read.toLocaleString()} of ${(ids.length + spSites.length).toLocaleString()} sites read.` : `Exported ${list.length.toLocaleString()} sites with their owners and members.`,
          'success'
        );
      } else {
        stopRef.current = false;
        const site = listSite.trim().replace(/\/+$/, '');
        const title = listName.trim();
        setExporting({ done: 0, total: list.length, phase: 'Preparing the list…' });
        const exporter = new ListExport(ctx.client);
        const info = await exporter.ensureList(site, title);
        const items = listItemsFor(list, names, accessOf, new Date());
        const r = await exporter.addItems(
          site,
          info.id,
          items,
          (d) => mounted.current && setExporting({ done: d, total: items.length, phase: 'Saving to the SharePoint list' }),
          () => stopRef.current
        );
        ctx.log('Save sites with owners and members to a list', `${info.url} (${r.added} items)`, r.failed.length === 0);
        if (mounted.current) {
          setSavedList({ url: info.url, title, added: r.added, failed: r.failed });
        }
        ctx.notify(
          `${r.added.toLocaleString()} sites saved to "${title}"${info.created ? ' (new list)' : ''}${r.failed.length ? `; ${r.failed.length} failed` : ''}${stopRef.current ? ' (stopped early)' : ''}.`,
          r.failed.length ? 'error' : 'success'
        );
      }
    } catch (e) {
      ctx.notify(graphErrorText(e), 'error');
    } finally {
      if (mounted.current) {
        setExporting(undefined);
      }
    }
  };

  const columns: IColumn[] = [
    {
      key: 't',
      name: 'Site',
      minWidth: 220,
      isResizable: true,
      onRender: (s: ISiteRow) => (
        <span style={{ minWidth: 0 }}>
          <a href={s.url} target="_blank" rel="noopener noreferrer" className={styles.link}>
            {s.title}
          </a>
          <div className={styles.muted} style={{ fontSize: 12, overflow: 'hidden', textOverflow: 'ellipsis' }} title={s.url}>
            {s.url}
          </div>
        </span>
      )
    },
    { key: 'ty', name: 'Type', minWidth: 150, isResizable: true, onRender: (s: ISiteRow) => siteType(s.template, s.groupConnected) },
    {
      key: 'g',
      name: 'Microsoft 365 group',
      minWidth: 200,
      isResizable: true,
      onRender: (s: ISiteRow) => {
        if (!s.groupId) {
          return s.relatedGroupId ? <Pill kind="info">Teams channel site</Pill> : <span className={styles.muted}>Not connected</span>;
        }
        const g = groups ? groups[s.groupId] : undefined;
        return (
          <span>
            <Pill kind="good">Connected</Pill> {g ? g.displayName : ''}
            {g && g.isTeam && <span className={styles.muted}> · Teams</span>}
          </span>
        );
      }
    },
    {
      key: 'a',
      name: '',
      minWidth: 150,
      onRender: (s: ISiteRow) =>
        s.groupId || s.manageable ? (
          <button type="button" className={styles.link} onClick={() => setOpen(s)}>
            Owners &amp; members
          </button>
        ) : null
    }
  ];

  return (
    <div className={styles.view}>
      <ViewHeader title="Sites & Microsoft 365 groups" hint="Every active site collection, whether it is connected to a Microsoft 365 group, and who its owners and members are." />

      {access.scopes && !graphOk && (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline>
          The site list, and owners and members of sites that are not group-connected (their SharePoint groups), work now. Group names and the owners and members of Microsoft 365 groups need the Microsoft Graph permission <strong>Group.Read.All</strong> (or the Group.ReadWrite.All this solution requests) to be approved by a Global Administrator.{' '}
          <button type="button" className={styles.link} onClick={() => ctx.openView('apiaccess')}>
            How to approve
          </button>
        </MessageBar>
      )}

      <Card>
        <p className={styles.muted} style={{ marginTop: 0 }}>
          Site collections come from SharePoint search (pages of 500, about 35 requests for 17,000 sites), so they are the active SharePoint sites you can see; OneDrive sites, deleted and archived sites are not listed. Group names are one Microsoft Graph request per 999 groups. Nothing loads until you click.
        </p>
        <div className={styles.actions}>
          <PrimaryButton iconProps={{ iconName: 'Globe' }} onClick={load} disabled={!!loading || !!exporting}>
            {loading || (data ? 'Reload' : 'Load all site collections')}
          </PrimaryButton>
          {data && (
            <span className={styles.muted} style={{ alignSelf: 'center' }}>
              Loaded {formatDate(data.at)} {data.at.toLocaleTimeString()}
              {data.oneDriveHidden > 0 && ` · ${data.oneDriveHidden.toLocaleString()} OneDrive sites left out`}
            </span>
          )}
        </div>
        {error && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline>
            {error}
          </MessageBar>
        )}
        {data && data.truncated && <MessageBar messageBarType={MessageBarType.warning}>Stopped at 100,000 sites.</MessageBar>}
      </Card>

      {data && (
        <>
          <div className={styles.statGrid}>
            <StatTile icon="Globe" label="Active site collections" value={sites.length.toLocaleString()} tone="brand" onClick={() => setShow('all')} />
            <StatTile icon="Group" label="Connected to a Microsoft 365 group" value={counts.group.toLocaleString()} sub={`${Math.round((counts.group / Math.max(1, sites.length)) * 100)}% of sites`} onClick={() => setShow('group')} />
            <StatTile icon="SharepointLogo" label="Not group-connected" value={(sites.length - counts.group - counts.channel).toLocaleString()} onClick={() => setShow('nogroup')} />
            {counts.channel > 0 && <StatTile icon="TeamsLogo" label="Teams channel sites" value={counts.channel.toLocaleString()} sub="private / shared channels" onClick={() => setShow('channel')} />}
          </div>

          <Card title="Export">
            <div className={styles.actions} style={{ alignItems: 'flex-end' }}>
              <DefaultButton iconProps={{ iconName: 'Download' }} onClick={exportSites} disabled={!!exporting}>
                Export site list ({rows.length.toLocaleString()})
              </DefaultButton>
              <ChoiceGroup
                label="Save as"
                selectedKey={saveTo}
                onChange={(_, o) => o && setSaveTo(o.key as 'csv' | 'list')}
                options={[
                  { key: 'csv', text: 'CSV file' },
                  { key: 'list', text: 'SharePoint list (one item per site)' }
                ]}
                disabled={!!exporting}
                styles={{ flexContainer: { display: 'flex', gap: 16, flexWrap: 'wrap' } }}
              />
              {saveTo === 'csv' ? (
                <ChoiceGroup
                  label="Layout"
                  selectedKey={layout}
                  onChange={(_, o) => o && setLayout(o.key as ExportLayout)}
                  options={[
                    { key: 'site', text: 'One row per site (e-mail addresses in cells)' },
                    { key: 'person', text: 'One row per person (easy to filter in Excel)' }
                  ]}
                  disabled={!!exporting}
                  styles={{ flexContainer: { display: 'flex', gap: 16, flexWrap: 'wrap' } }}
                />
              ) : (
                <>
                  <TextField label="Site for the list" value={listSite} onChange={(_, v) => setListSite(v || '')} disabled={!!exporting} styles={{ root: { minWidth: 280, flex: 1 } }} />
                  <TextField label="List name" value={listName} onChange={(_, v) => setListName(v || '')} disabled={!!exporting} styles={{ root: { minWidth: 240 } }} />
                </>
              )}
              <Checkbox label="Include site collection administrators (1 more request per site)" checked={includeAdmins} onChange={(_, v) => setIncludeAdmins(!!v)} disabled={!!exporting} />
              <PrimaryButton iconProps={{ iconName: saveTo === 'list' ? 'BulletedList' : 'ExcelDocument' }} onClick={exportMembership} disabled={!!exporting || !!loading}>
                {saveTo === 'list' ? 'Save owners & members to the list' : 'Export with owners & members'}
              </PrimaryButton>
              {exporting && (
                <DefaultButton
                  iconProps={{ iconName: 'Stop' }}
                  onClick={() => {
                    stopRef.current = true;
                  }}
                >
                  {exporting.phase ? 'Stop saving' : 'Stop and export what is read'}
                </DefaultButton>
              )}
            </div>
            {exporting && (
              <ProgressIndicator
                label={`${exporting.phase || 'Reading owners and members'}… ${exporting.done.toLocaleString()} of ${exporting.total.toLocaleString()} sites`}
                percentComplete={exporting.total ? exporting.done / exporting.total : undefined}
              />
            )}
            {savedList && (
              <MessageBar messageBarType={savedList.failed.length ? MessageBarType.warning : MessageBarType.success} isMultiline onDismiss={() => setSavedList(undefined)}>
                {savedList.added.toLocaleString()} sites saved to{' '}
                <a href={savedList.url} target="_blank" rel="noopener noreferrer">
                  {savedList.title}
                </a>
                .{savedList.failed.length > 0 && ` Not saved: ${savedList.failed.slice(0, 10).join(', ')}${savedList.failed.length > 10 ? '…' : ''}`}
              </MessageBar>
            )}
            <p className={styles.muted} style={{ marginBottom: 0, fontSize: 12 }}>
              People are listed by e-mail address, separated by semicolons. In the CSV, a cell too long for Excel continues in a &ldquo;(part 2)&rdquo; column so rows never break. In the list, very long columns are cut with a count of what was left out (use the CSV for those). Exports follow the filter below. Group-connected sites: owners and members of the Microsoft 365 group (Graph, 10 groups per request). Other sites: the people in the site&rsquo;s SharePoint Owners, Members and Visitors groups (one SharePoint request per site, at most 2 at a time).
            </p>
          </Card>

          <Card
            title={`${rows.length.toLocaleString()} site collections`}
            right={
              <Dropdown
                ariaLabel="Show"
                selectedKey={show}
                onChange={(_, o) => o && setShow(o.key as Show)}
                options={[
                  { key: 'all', text: 'All sites' },
                  { key: 'group', text: 'Group-connected' },
                  { key: 'nogroup', text: 'Not group-connected' },
                  { key: 'channel', text: 'Teams channel sites' }
                ]}
                styles={{ root: { width: 190 } }}
              />
            }
          >
            <SearchBox placeholder="Filter by title, URL or group name" value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { maxWidth: 360, marginBottom: 8 } }} />
            {rows.length === 0 ? <Empty text="No sites match." /> : <DetailsList items={rows} columns={columns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified} />}
          </Card>
        </>
      )}

      <GroupPanel site={open} group={open && open.groupId && groups ? groups[open.groupId] : undefined} graphOk={graphOk} onDismiss={() => setOpen(undefined)} />
    </div>
  );
};

const GroupPanel: React.FC<{ site?: ISiteRow; group?: IGroupSummary; graphOk: boolean; onDismiss: () => void }> = ({ site, group, graphOk, onDismiss }) => {
  const ctx = useAdmin();
  const [state, setState] = React.useState<{ loading: boolean; data?: ISiteAccess; error?: string }>({ loading: false });
  const url = site && site.url;
  const groupId = site && site.groupId;
  const needsGraph = !!groupId && !graphOk;

  const load = React.useCallback(
    (fresh: boolean) => {
      if (!url || needsGraph) {
        return () => undefined;
      }
      let alive = true;
      setState({ loading: true });
      // Microsoft 365 group sites: the group's owners and members. Other sites: their SharePoint groups and admins.
      const read: Promise<ISiteAccess> = groupId ? ctx.graph.membership(groupId, fresh).then(fromGroupMembership) : ctx.api.siteAccess(url, true);
      read.then(
        (data) => alive && setState({ loading: false, data }),
        (e) => alive && setState({ loading: false, error: graphErrorText(e) })
      );
      return () => {
        alive = false;
      };
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [url, groupId, needsGraph]
  );
  React.useEffect(() => load(false), [load]);

  const manage = async (): Promise<void> => {
    if (!site) {
      return;
    }
    try {
      const [web, s] = await Promise.all([ctx.api.getWeb(site.url), ctx.api.getSite(site.url)]);
      ctx.setTarget({ webUrl: web.url, siteUrl: s.url, title: web.title });
      ctx.openView(groupId ? 'm365groups' : 'people');
    } catch (e) {
      ctx.notify((e as Error).message, 'error');
    }
  };

  const d = state.data;
  const sp = !!d && d.source === 'sharepoint';
  return (
    <Panel isOpen={!!site} onDismiss={onDismiss} type={PanelType.medium} headerText={group ? group.displayName : site ? site.title : ''} closeButtonAriaLabel="Close" isLightDismiss>
      {site && (
        <>
          <p className={styles.muted} style={{ wordBreak: 'break-all', marginTop: 0 }}>
            {site.url}
            <br />
            {group ? (
              <>
                Microsoft 365 group · {group.mail} · {group.visibility || 'privacy unknown'}
                {group.isTeam ? ' · Microsoft Teams team' : ''}
              </>
            ) : groupId ? (
              'Microsoft 365 group'
            ) : (
              'Not group-connected: owners and members come from the site\u2019s SharePoint groups'
            )}
          </p>
          {needsGraph ? (
            <NeedsApproval scopes={[GROUP_SCOPE]} feature="Showing Microsoft 365 group owners and members" />
          ) : state.loading ? (
            <ProgressIndicator label="Reading owners and members…" />
          ) : state.error || (d && d.error) ? (
            <MessageBar messageBarType={MessageBarType.error} isMultiline>
              {state.error || (d && d.error)}
            </MessageBar>
          ) : d ? (
            <>
              {d.owners.filter((o) => o.type === 'user').length === 0 && <MessageBar messageBarType={MessageBarType.severeWarning}>{sp ? 'The Owners group has no people in it.' : 'This group has no owners.'}</MessageBar>}
              <div className={styles.actions} style={{ margin: '8px 0' }}>
                <DefaultButton
                  iconProps={{ iconName: 'Download' }}
                  onClick={() =>
                    downloadCsv(
                      `${(group ? group.displayName : site.title).replace(/[^A-Za-z0-9_-]+/g, '-')}-owners-members.csv`,
                      PERSON_HEADERS,
                      siteGroupRows([site], group && site.groupId ? { [site.groupId]: group } : {}, () => d, 'person')
                    )
                  }
                >
                  Export this site
                </DefaultButton>
                <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={() => load(true)}>
                  Refresh
                </DefaultButton>
                <DefaultButton iconProps={{ iconName: 'Settings' }} onClick={manage}>
                  Manage owners &amp; members
                </DefaultButton>
              </div>
              {d.truncated && <MessageBar messageBarType={MessageBarType.info}>Very large group: the first 5,000 are shown.</MessageBar>}
              <PeopleList title={`Owners (${d.owners.length})${sp && d.ownerGroup ? ` · ${d.ownerGroup}` : ''}`} people={d.owners} />
              <PeopleList title={`Members (${d.members.length})${sp && d.memberGroup ? ` · ${d.memberGroup}` : ''}`} people={d.members} />
              {sp && <PeopleList title={`Visitors (${d.visitors.length})${d.visitorGroup ? ` · ${d.visitorGroup}` : ''}`} people={d.visitors} />}
              {d.admins && <PeopleList title={`Site collection administrators (${d.admins.length})`} people={d.admins} />}
            </>
          ) : null}
        </>
      )}
    </Panel>
  );
};

const PeopleList: React.FC<{ title: string; people: IDirectoryObject[] }> = ({ title, people }) => (
  <>
    <h3 style={{ fontSize: 14, margin: '16px 0 6px' }}>{title}</h3>
    {people.length === 0 ? (
      <Empty text="None." />
    ) : (
      <ul className={styles.memberList}>
        {people.map((p) => (
          <li key={p.id} className={styles.memberRow}>
            <span style={{ minWidth: 0 }}>
              <div>
                {p.displayName} {(p.userType || '').toLowerCase() === 'guest' && <Pill kind="warning">Guest</Pill>} {p.type === 'group' && <Pill kind="info">Group</Pill>}
                {p.type === 'servicePrincipal' && <Pill kind="info">App</Pill>}
              </div>
              <div className={styles.muted} style={{ fontSize: 12, overflow: 'hidden', textOverflow: 'ellipsis' }}>
                {p.mail || p.userPrincipalName}
              </div>
            </span>
          </li>
        ))}
      </ul>
    )}
  </>
);
