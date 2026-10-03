import * as React from 'react';
import { DefaultButton, PrimaryButton } from '@fluentui/react';
import { ISubWeb, ViewKey } from '../../models';
import { evaluateHealth } from '../../services/HealthEngine';
import { formatBytes, formatDate, formatCompact, relativeTime } from '../../services/format';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { useLoader } from '../shared/hooks';
import { GrowthAlert } from './GrowthAlert';
import { BarList, Card, Donut, ErrorBar, FindingRow, Kpi, Loading, Pill, ScoreRing, ViewHeader } from '../shared/ui';

/** Dashboard: one screen that answers "is this site OK, and what needs me?". */
export const OverviewView: React.FC = () => {
  const ctx = useAdmin();
  const { api, target } = ctx;

  const data = useLoader(async () => {
    // Small, bounded metadata calls (shared with the other views through the cache).
    const [web, site, lists, users, subwebs] = await Promise.all([
      api.getWeb(target.webUrl),
      api.getSite(target.webUrl),
      api.getLists(target.webUrl),
      api.getUsers(target.webUrl),
      api.getSubWebs(target.webUrl).catch((): ISubWeb[] => [])
    ]);
    return { web, site, lists, users: users.users, usersTruncated: users.truncated, subwebs };
  }, [target.webUrl]);

  const loaded = data.data;
  React.useEffect(() => {
    if (loaded) {
      const h = evaluateHealth({ site: loaded.site, web: loaded.web, users: loaded.users, lists: loaded.lists, subwebs: loaded.subwebs });
      ctx.setHealthBadge(h.findings.filter((f) => f.severity === 'critical' || f.severity === 'warning').length);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [loaded]);

  if (data.loading && !data.data) {
    return <Loading text="Loading the dashboard…" />;
  }
  if (data.error || !data.data) {
    return <ErrorBar error={data.error || 'Nothing loaded.'} onRetry={data.reload} />;
  }
  const { web, site, lists, users, usersTruncated, subwebs } = data.data;
  const health = evaluateHealth({ site, web, users, lists, subwebs });
  const libraries = lists.filter((l) => l.kind === 'Library' && !l.hidden && !l.isSystem);
  const plainLists = lists.filter((l) => l.kind === 'List' && !l.hidden && !l.isSystem);
  const guests = users.filter((u) => u.kind === 'Guest');
  const people = users.filter((u) => u.kind === 'Member');
  const admins = users.filter((u) => u.isSiteAdmin && u.kind === 'Member');
  const items = lists.filter((l) => !l.hidden && !l.isSystem).reduce((s, l) => s + l.itemCount, 0);
  const goTo = (v: ViewKey) => (): void => ctx.openView(v);
  const attention = health.findings.filter((f) => f.severity === 'critical' || f.severity === 'warning');
  const quota = site.storageFraction > 0 ? site.storageBytes / site.storageFraction : 0;
  const top = lists
    .filter((l) => !l.hidden && !l.isSystem)
    .sort((a, b) => b.itemCount - a.itemCount)
    .slice(0, 6);

  return (
    <div className={styles.view}>
      <ViewHeader title="Dashboard" hint={`${web.title} · ${web.url}`}>
        <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={ctx.reloadAll} disabled={data.loading}>
          Refresh
        </DefaultButton>
        <PrimaryButton iconProps={{ iconName: 'Health' }} onClick={() => ctx.openView('health')}>
          Full health check
        </PrimaryButton>
      </ViewHeader>

      <GrowthAlert />

      <div className={styles.kpiRow}>
        <Kpi label="Storage used" value={formatBytes(site.storageBytes)} sub={quota ? `of ${formatBytes(quota)} · ${Math.round(site.storageFraction * 100)}%` : 'site collection'} onClick={() => ctx.openView('storage')} />
        <Kpi label="Libraries" value={libraries.length.toLocaleString()} sub={`${plainLists.length} other lists`} onClick={() => ctx.openView('content')} />
        <Kpi label="Items" value={formatCompact(items)} sub="in visible lists" onClick={() => ctx.openView('content')} />
        <Kpi label="People" value={formatCompact(people.length)} sub={usersTruncated ? 'first 5,000 shown' : `${admins.length} admin${admins.length === 1 ? '' : 's'}`} onClick={() => ctx.openView('people')} />
        <Kpi label="Guests" value={guests.length.toLocaleString()} sub="external users" onClick={() => ctx.openView('people')} />
        <Kpi label="Subsites" value={subwebs.length.toLocaleString()} sub="direct children" />
      </div>

      <div className={styles.grid}>
        <Card title="Health">
          <div className={styles.scoreRing}>
            <ScoreRing score={health.score} />
            <div className={styles.muted}>
              {attention.length === 0 ? 'Nothing needs your attention right now.' : `${attention.length} item${attention.length === 1 ? '' : 's'} need${attention.length === 1 ? 's' : ''} attention.`}
            </div>
          </div>
        </Card>

        <Card title="Storage">
          {site.storageFraction > 0 ? (
            <Donut
              slices={[
                { label: 'Used', value: site.storageBytes, color: site.storageFraction >= 0.9 ? '#c4314b' : site.storageFraction >= 0.75 ? '#b87400' : '#2563eb' },
                { label: 'Free', value: Math.max(0, quota - site.storageBytes), color: '#d0d4da' }
              ]}
              centre={`${Math.round(site.storageFraction * 100)}%`}
              centreSub="of quota"
              format={formatBytes}
            />
          ) : (
            <p className={styles.muted}>Storage figures are not available for this site.</p>
          )}
        </Card>

        <Card title="Largest lists by item count">
          {top.length ? <BarList rows={top.map((l) => ({ label: l.title, value: l.itemCount }))} /> : <p className={styles.muted}>No lists yet.</p>}
        </Card>
      </div>

      {attention.length > 0 && (
        <Card title="Needs attention">
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
            {attention.slice(0, 5).map((f) => (
              <FindingRow key={f.id} severity={f.severity} title={f.title} detail={f.detail} onGo={f.view ? goTo(f.view) : undefined} />
            ))}
          </div>
        </Card>
      )}

      <div className={styles.grid}>
        <Card title="About this site">
          <dl style={{ margin: 0, display: 'grid', gridTemplateColumns: 'auto 1fr', gap: '6px 14px', fontSize: 13 }}>
            <dt className={styles.muted}>Template</dt>
            <dd style={{ margin: 0 }}>{web.template || '-'}</dd>
            <dt className={styles.muted}>Created</dt>
            <dd style={{ margin: 0 }}>{formatDate(web.created)}</dd>
            <dt className={styles.muted}>Last content change</dt>
            <dd style={{ margin: 0 }}>{relativeTime(web.lastModified)}</dd>
            <dt className={styles.muted}>Primary owner</dt>
            <dd style={{ margin: 0 }}>{site.ownerTitle || '-'}</dd>
            <dt className={styles.muted}>Microsoft 365 group</dt>
            <dd style={{ margin: 0 }}>{site.groupId && !/^0+(-0+)*$/.test(site.groupId) ? <Pill kind="info">Connected</Pill> : 'No'}</dd>
            <dt className={styles.muted}>Status</dt>
            <dd style={{ margin: 0 }}>{site.readOnly ? <Pill kind="warning">Read-only</Pill> : <Pill kind="good">Active</Pill>}</dd>
          </dl>
        </Card>

        <Card title="Quick actions">
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8, alignItems: 'flex-start' }}>
            <DefaultButton iconProps={{ iconName: 'AddFriend' }} onClick={() => ctx.openView('people')}>
              Add people to a group
            </DefaultButton>
            <DefaultButton iconProps={{ iconName: 'AreaChart' }} onClick={() => ctx.openView('growth')}>
              Track storage growth
            </DefaultButton>
            <DefaultButton iconProps={{ iconName: 'FileCode' }} onClick={() => ctx.openView('storage')}>
              Find the largest files
            </DefaultButton>
            <DefaultButton iconProps={{ iconName: 'RecycleBin' }} onClick={() => ctx.openView('recycle')}>
              Restore deleted items
            </DefaultButton>
            <DefaultButton iconProps={{ iconName: 'History' }} onClick={() => ctx.openView('activity')}>
              See recent permission changes
            </DefaultButton>
          </div>
        </Card>

        {subwebs.length > 0 && (
          <Card title="Subsites">
            <ul className={styles.memberList}>
              {subwebs.slice(0, 8).map((s) => (
                <li key={s.id} className={styles.memberRow}>
                  <span>{s.title}</span>
                  <button type="button" className={styles.link} onClick={() => ctx.setTarget({ webUrl: s.url, siteUrl: target.siteUrl, title: s.title })}>
                    Manage
                  </button>
                </li>
              ))}
            </ul>
          </Card>
        )}
      </div>
    </div>
  );
};
