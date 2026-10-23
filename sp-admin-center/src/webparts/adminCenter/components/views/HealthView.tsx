import * as React from 'react';
import { DefaultButton } from '@fluentui/react';
import { ISubWeb, ViewKey } from '../../models';
import { downloadCsv } from '../../services/exportCsv';
import { evaluateHealth } from '../../services/HealthEngine';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { useLoader } from '../shared/hooks';
import { Card, ErrorBar, FindingRow, Loading, ScoreRing, ViewHeader } from '../shared/ui';

const BIN_SAMPLE = 300;

/** Rule-based governance check over cached, bounded metadata calls. */
export const HealthView: React.FC = () => {
  const ctx = useAdmin();
  const { api, target } = ctx;

  const loader = useLoader(async () => {
    const [web, site, lists, users, subwebs] = await Promise.all([
      api.getWeb(target.webUrl),
      api.getSite(target.webUrl),
      api.getLists(target.webUrl),
      api.getUsers(target.webUrl),
      api.getSubWebs(target.webUrl).catch((): ISubWeb[] => [])
    ]);
    const bin = await api.getRecycleBin(target.webUrl, BIN_SAMPLE).catch(() => undefined);
    return evaluateHealth({
      web,
      site,
      lists,
      users: users.users,
      subwebs,
      recycleCount: bin ? bin.length : undefined,
      recycleCapped: bin ? bin.length >= BIN_SAMPLE : false
    });
  }, [target.webUrl]);

  const goTo = (v: ViewKey) => (): void => ctx.openView(v);
  const attention = loader.data ? loader.data.findings.filter((f) => f.severity === 'critical' || f.severity === 'warning').length : undefined;
  React.useEffect(() => {
    if (attention !== undefined) {
      ctx.setHealthBadge(attention);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [attention]);

  if (loader.loading && !loader.data) {
    return <Loading text="Checking the site…" />;
  }
  if (loader.error || !loader.data) {
    return <ErrorBar error={loader.error || 'Nothing loaded.'} onRetry={loader.reload} />;
  }
  const { score, findings } = loader.data;
  const issues = findings.filter((f) => f.severity !== 'good');
  const good = findings.filter((f) => f.severity === 'good');

  return (
    <div className={styles.view}>
      <ViewHeader title="Health check" hint="Rules that catch common governance, security and capacity problems.">
        <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={loader.reload} disabled={loader.loading}>
          Re-check
        </DefaultButton>
        <DefaultButton iconProps={{ iconName: 'Download' }} onClick={() => downloadCsv('health-check.csv', ['Severity', 'Finding', 'Detail'], findings.map((f) => [f.severity, f.title, f.detail]))}>
          Export CSV
        </DefaultButton>
      </ViewHeader>

      <Card>
        <div className={styles.scoreRing}>
          <ScoreRing score={score} />
          <div>
            <div style={{ fontSize: 18, fontWeight: 600 }}>{score >= 85 ? 'In good shape' : score >= 60 ? 'Needs some attention' : 'Needs attention'}</div>
            <div className={styles.muted}>
              {issues.length} finding{issues.length === 1 ? '' : 's'} · {good.length} check{good.length === 1 ? '' : 's'} passed · {target.title}
            </div>
          </div>
        </div>
      </Card>

      {issues.length > 0 && (
        <Card title="Findings">
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
            {issues.map((f) => (
              <FindingRow key={f.id} severity={f.severity} title={f.title} detail={f.detail} onGo={f.view && f.view !== 'health' ? goTo(f.view) : undefined} />
            ))}
          </div>
        </Card>
      )}

      <Card title="Passed checks">
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {good.length === 0 && <span className={styles.muted}>None yet.</span>}
          {good.map((f) => (
            <FindingRow key={f.id} severity="good" title={f.title} />
          ))}
        </div>
      </Card>
    </div>
  );
};
