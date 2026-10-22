import * as React from 'react';
import { ChoiceGroup, DefaultButton, DetailsList, DetailsListLayoutMode, Dropdown, IColumn, MessageBar, MessageBarType, Panel, PanelType, PrimaryButton, ProgressIndicator, SearchBox, SelectionMode } from '@fluentui/react';
import { ISiteRow } from '../../models';
import { downloadCsv } from '../../services/exportCsv';
import { formatDate } from '../../services/format';
import { hasScope, IDirectoryObject, IGroupMembership, IGroupSummary } from '../../services/GraphApi';
import { ExportLayout, PERSON_HEADERS, SITE_HEADERS, siteGroupRows } from '../../services/siteGroupsExport';
import { siteType } from '../../services/siteTypes';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { graphErrorText, NeedsApproval, useGraphScopes } from '../shared/graph';
import { Card, Empty, Pill, StatTile, ViewHeader } from '../shared/ui';

type Show = 'all' | 'group' | 'nogroup' | 'channel';

const GROUP_SCOPE = 'Group.Read.All';

/** Loaded once per page visit; kept while you move between views. */
let sessionSites: { sites: ISiteRow[]; truncated: boolean; at: Date } | undefined;
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
  const [exporting, setExporting] = React.useState<{ done: number; total: number } | undefined>();
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
      sessionSites = { sites: r.sites, truncated: r.truncated, at: new Date() };
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
    const ids = list.filter((s) => !!s.groupId).map((s) => s.groupId as string);
    const ok = await ctx.confirm({
      title: 'Export owners and members',
      message: `Read the owners and members of ${ids.length.toLocaleString()} Microsoft 365 groups for ${list.length.toLocaleString()} sites? This uses about ${Math.max(1, Math.ceil(ids.length / 10)).toLocaleString()} batched Microsoft Graph requests, one at a time (roughly ${Math.max(1, Math.round(ids.length / 10 / 60))} minute${ids.length > 600 ? 's' : ''}). You can stop it and export what was read so far.`,
      confirmText: 'Start export'
    });
    if (!ok) {
      return;
    }
    stopRef.current = false;
    setExporting({ done: 0, total: ids.length });
    try {
      let names = groups;
      if (!names) {
        names = sessionGroups = await ctx.graph.groupDirectory();
        if (mounted.current) {
          setGroups(names);
        }
      }
      const mem = await ctx.graph.membershipMany(ids, (done, total) => mounted.current && setExporting({ done, total }), () => stopRef.current);
      if (!mounted.current) {
        return;
      }
      const stamp = new Date().toISOString().substring(0, 10);
      downloadCsv(
        layout === 'site' ? `sites-m365-owners-members-${stamp}.csv` : `sites-m365-people-${stamp}.csv`,
        layout === 'site' ? SITE_HEADERS : PERSON_HEADERS,
        siteGroupRows(list, names || {}, mem, layout)
      );
      const read = Object.keys(mem).length;
      ctx.log('Export sites with group owners and members', `${list.length} sites, ${read} groups`, true);
      ctx.notify(stopRef.current ? `Stopped: exported with ${read.toLocaleString()} of ${ids.length.toLocaleString()} groups read.` : `Exported ${list.length.toLocaleString()} sites with the owners and members of ${read.toLocaleString()} groups.`, 'success');
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
        s.groupId ? (
          <button type="button" className={styles.link} onClick={() => setOpen(s)}>
            Owners &amp; members
          </button>
        ) : null
    }
  ];

  return (
    <div className={styles.view}>
      <ViewHeader title="Sites & Microsoft 365 groups" hint="Every active site collection, whether it is connected to a Microsoft 365 group, and who owns and belongs to that group." />

      {access.scopes && !graphOk && (
        <MessageBar messageBarType={MessageBarType.warning} isMultiline>
          The site list works now. Showing group names, owners and members needs the Microsoft Graph permission <strong>Group.Read.All</strong> (or the Group.ReadWrite.All this solution requests) to be approved by a Global Administrator.{' '}
          <button type="button" className={styles.link} onClick={() => ctx.openView('apiaccess')}>
            How to approve
          </button>
        </MessageBar>
      )}

      <Card>
        <p className={styles.muted} style={{ marginTop: 0 }}>
          Site collections come from SharePoint search (pages of 500, about 35 requests for 17,000 sites), so they are the active sites you can see; deleted and archived sites are not listed. Group names are one Microsoft Graph request per 999 groups. Nothing loads until you click.
        </p>
        <div className={styles.actions}>
          <PrimaryButton iconProps={{ iconName: 'Globe' }} onClick={load} disabled={!!loading || !!exporting}>
            {loading || (data ? 'Reload' : 'Load all site collections')}
          </PrimaryButton>
          {data && (
            <span className={styles.muted} style={{ alignSelf: 'center' }}>
              Loaded {formatDate(data.at)} {data.at.toLocaleTimeString()}
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
                selectedKey={layout}
                onChange={(_, o) => o && setLayout(o.key as ExportLayout)}
                options={[
                  { key: 'site', text: 'One row per site (owners and members in cells)' },
                  { key: 'person', text: 'One row per person (easy to filter in Excel)' }
                ]}
                disabled={!!exporting}
                styles={{ flexContainer: { display: 'flex', gap: 16, flexWrap: 'wrap' } }}
              />
              <PrimaryButton iconProps={{ iconName: 'ExcelDocument' }} onClick={exportMembership} disabled={!graphOk || !!exporting || !!loading}>
                Export with owners &amp; members
              </PrimaryButton>
              {exporting && (
                <DefaultButton
                  iconProps={{ iconName: 'Stop' }}
                  onClick={() => {
                    stopRef.current = true;
                  }}
                >
                  Stop and export what is read
                </DefaultButton>
              )}
            </div>
            {exporting && (
              <ProgressIndicator
                label={`Reading owners and members… ${exporting.done.toLocaleString()} of ${exporting.total.toLocaleString()} groups`}
                percentComplete={exporting.total ? exporting.done / exporting.total : undefined}
              />
            )}
            <p className={styles.muted} style={{ marginBottom: 0, fontSize: 12 }}>
              Exports follow the filter below. Owners and members are read in batches of 10 groups per request, one batch at a time, and reused for 10 minutes.
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
  const [state, setState] = React.useState<{ loading: boolean; data?: IGroupMembership; error?: string }>({ loading: false });
  const groupId = site && site.groupId;

  const load = React.useCallback(
    (fresh: boolean) => {
      if (!groupId || !graphOk) {
        return () => undefined;
      }
      let alive = true;
      setState({ loading: true });
      ctx.graph.membership(groupId, fresh).then(
        (data) => alive && setState({ loading: false, data }),
        (e) => alive && setState({ loading: false, error: graphErrorText(e) })
      );
      return () => {
        alive = false;
      };
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [groupId, graphOk]
  );
  React.useEffect(() => load(false), [load]);

  const manage = async (): Promise<void> => {
    if (!site) {
      return;
    }
    try {
      const [web, s] = await Promise.all([ctx.api.getWeb(site.url), ctx.api.getSite(site.url)]);
      ctx.setTarget({ webUrl: web.url, siteUrl: s.url, title: web.title });
      ctx.openView('m365groups');
    } catch (e) {
      ctx.notify((e as Error).message, 'error');
    }
  };

  const d = state.data;
  return (
    <Panel isOpen={!!site} onDismiss={onDismiss} type={PanelType.medium} headerText={group ? group.displayName : site ? site.title : ''} closeButtonAriaLabel="Close" isLightDismiss>
      {site && (
        <>
          <p className={styles.muted} style={{ wordBreak: 'break-all', marginTop: 0 }}>
            {site.url}
            {group && (
              <>
                <br />
                {group.mail} · {group.visibility || 'privacy unknown'}
                {group.isTeam ? ' · Microsoft Teams team' : ''}
              </>
            )}
          </p>
          {!graphOk ? (
            <NeedsApproval scopes={[GROUP_SCOPE]} feature="Showing group owners and members" />
          ) : state.loading ? (
            <ProgressIndicator label="Reading owners and members…" />
          ) : state.error ? (
            <MessageBar messageBarType={MessageBarType.error} isMultiline>
              {state.error}
            </MessageBar>
          ) : d ? (
            <>
              {d.owners.filter((o) => o.type === 'user').length === 0 && <MessageBar messageBarType={MessageBarType.severeWarning}>This group has no owners.</MessageBar>}
              <div className={styles.actions} style={{ margin: '8px 0' }}>
                <DefaultButton
                  iconProps={{ iconName: 'Download' }}
                  onClick={() =>
                    downloadCsv(
                      `${(group ? group.displayName : site.title).replace(/[^A-Za-z0-9_-]+/g, '-')}-owners-members.csv`,
                      PERSON_HEADERS,
                      siteGroupRows([site], group ? { [site.groupId as string]: group } : {}, { [site.groupId as string]: d }, 'person')
                    )
                  }
                >
                  Export this group
                </DefaultButton>
                <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={() => load(true)}>
                  Refresh
                </DefaultButton>
                <DefaultButton iconProps={{ iconName: 'Settings' }} onClick={manage}>
                  Manage owners &amp; members
                </DefaultButton>
              </div>
              {d.truncated && <MessageBar messageBarType={MessageBarType.info}>Very large group: the first 5,000 are shown.</MessageBar>}
              <PeopleList title={`Owners (${d.owners.length})`} people={d.owners} />
              <PeopleList title={`Members (${d.members.length})`} people={d.members} />
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
