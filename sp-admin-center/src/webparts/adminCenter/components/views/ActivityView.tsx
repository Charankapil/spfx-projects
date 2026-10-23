import * as React from 'react';
import { DefaultButton, DetailsList, DetailsListLayoutMode, Dropdown, MessageBar, MessageBarType, SelectionMode } from '@fluentui/react';
import { IChangeRow, IGroupInfo, IListInfo, IUserInfo } from '../../models';
import { downloadCsv } from '../../services/exportCsv';
import { formatDateTime } from '../../services/format';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { useLoader } from '../shared/hooks';
import { Card, Empty, ErrorBar, Kpi, Loading, Pill, ViewHeader } from '../shared/ui';

/**
 * Site-collection change log: who was added or removed, permission changes,
 * new and deleted lists, groups and webs. One request returns up to 500
 * changes; item and file edits are excluded on purpose (that is the part that
 * would be heavy), so this view stays cheap on busy sites.
 */
export const ActivityView: React.FC = () => {
  const ctx = useAdmin();
  const { api, target } = ctx;
  const [days, setDays] = React.useState(7);
  const [filter, setFilter] = React.useState('all');

  const loader = useLoader(async () => {
    const site = await api.getSite(target.webUrl);
    const changes = await api.getChanges(target.webUrl, site.id, days);
    // Names for ids; these are cached, cheap calls and optional.
    const [users, groups, lists] = await Promise.all([
      api.getUsers(target.webUrl).then((r) => r.users).catch((): IUserInfo[] => []),
      api.getGroups(target.webUrl).catch((): IGroupInfo[] => []),
      api.getLists(target.webUrl).catch((): IListInfo[] => [])
    ]);
    const names: { [k: string]: string } = {};
    users.forEach((u) => (names['User:' + u.id] = u.title));
    groups.forEach((g) => (names['Group:' + g.id] = g.title));
    lists.forEach((l) => (names['List:' + l.id.toLowerCase()] = l.title));
    return { changes, names };
  }, [target.webUrl, days]);

  const nameOf = (c: IChangeRow, names: { [k: string]: string }): string => {
    const key = c.objectKind + ':' + (c.objectKind === 'List' ? c.objectId.toLowerCase() : c.objectId);
    return names[key] || (c.objectId ? `${c.objectKind} ${c.objectId.substring(0, 8)}` : c.objectKind);
  };

  if (loader.loading && !loader.data) {
    return <Loading text="Reading the change log…" />;
  }
  if (loader.error || !loader.data) {
    return <ErrorBar error={loader.error || 'Nothing loaded.'} onRetry={loader.reload} />;
  }
  const { changes, names } = loader.data;
  const rows = changes.filter((c) => filter === 'all' || c.severity === filter);
  const count = (s: string): number => changes.filter((c) => c.severity === s).length;

  return (
    <div className={styles.view}>
      <ViewHeader title="Activity" hint="Permission and structure changes in this site collection, newest first. SharePoint keeps about 60 days.">
        <Dropdown
          ariaLabel="Period"
          selectedKey={days}
          onChange={(_, o) => setDays(Number(o ? o.key : 7))}
          options={[
            { key: 1, text: 'Last 24 hours' },
            { key: 7, text: 'Last 7 days' },
            { key: 30, text: 'Last 30 days' },
            { key: 60, text: 'Last 60 days' }
          ]}
          styles={{ root: { minWidth: 150 } }}
        />
        <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={loader.reload} disabled={loader.loading}>
          Refresh
        </DefaultButton>
        <DefaultButton
          iconProps={{ iconName: 'Download' }}
          onClick={() => downloadCsv('activity.csv', ['Time', 'Change', 'Object', 'Name', 'Category'], rows.map((c) => [c.time, c.changeType, c.objectKind, nameOf(c, names), c.severity]))}
        >
          Export CSV
        </DefaultButton>
      </ViewHeader>

      {changes.length >= 500 && <MessageBar messageBarType={MessageBarType.info}>Showing the first 500 changes of the period. Choose a shorter period for the rest.</MessageBar>}

      <div className={styles.kpiRow}>
        <Kpi label="Permission changes" value={String(count('security'))} sub="groups, members, roles" />
        <Kpi label="Created / deleted" value={String(count('structure'))} sub="sites, lists, groups, users" />
        <Kpi label="Other updates" value={String(count('info'))} />
      </div>

      <div className={styles.filters}>
        <Dropdown
          className={styles.field}
          label="Show"
          selectedKey={filter}
          onChange={(_, o) => setFilter(String(o ? o.key : 'all'))}
          options={[
            { key: 'all', text: 'All changes' },
            { key: 'security', text: 'Permission changes' },
            { key: 'structure', text: 'Created / deleted' },
            { key: 'info', text: 'Other updates' }
          ]}
        />
      </div>

      <Card>
        {rows.length === 0 ? (
          <Empty text="No changes in this period." />
        ) : (
          <DetailsList
            items={rows.slice(0, 500)}
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
            columns={[
              { key: 'time', name: 'When', minWidth: 140, onRender: (c: IChangeRow) => formatDateTime(c.time) },
              {
                key: 'what',
                name: 'Change',
                minWidth: 160,
                onRender: (c: IChangeRow) => <Pill kind={c.severity === 'security' ? 'warning' : c.severity === 'structure' ? 'info' : undefined}>{c.changeType}</Pill>
              },
              { key: 'obj', name: 'Object', minWidth: 90, onRender: (c: IChangeRow) => c.objectKind },
              { key: 'name', name: 'Name', minWidth: 220, isResizable: true, onRender: (c: IChangeRow) => nameOf(c, names) }
            ]}
          />
        )}
      </Card>
    </div>
  );
};
