import * as React from 'react';
import { Icon, DefaultButton } from '@fluentui/react';
import { formatBytes } from '../../services/format';
import { analyzeAll, anomaliesOf } from '../../services/GrowthEngine';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { useLoader } from '../shared/hooks';
import { Card, Sparkline, severityIcon } from '../shared/ui';
import { growthSummary } from './GrowthView';

/**
 * Dashboard pop-up for sites whose storage is growing unusually fast. It only
 * reads the history file (one request, cached) and stays silent when there is
 * no history or the file cannot be read.
 */
export const GrowthAlert: React.FC = () => {
  const ctx = useAdmin();
  const { growth } = ctx;
  const loader = useLoader(() => growth.load().catch(() => undefined), []);
  const anomalies = React.useMemo(() => (loader.data ? anomaliesOf(analyzeAll(loader.data)) : []), [loader.data]);
  const count = anomalies.length;
  React.useEffect(() => {
    if (!loader.loading) {
      ctx.setGrowthBadge(count);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [count, loader.loading]);

  if (count === 0) {
    return null;
  }
  return (
    <Card
      title={`${count} site${count === 1 ? ' has' : 's have'} grown unusually fast`}
      right={
        <DefaultButton onClick={() => ctx.openView('growth')} iconProps={{ iconName: 'AreaChart' }}>
          See all
        </DefaultButton>
      }
    >
      <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        {anomalies.slice(0, 3).map((s) => {
          const si = severityIcon(s.status === 'critical' ? 'critical' : 'warning');
          return (
            <div key={s.url} className={styles.finding}>
              <Icon iconName={si.icon} className={styles.findingIcon} style={{ color: si.color }} aria-hidden="true" />
              <div style={{ minWidth: 0 }}>
                <div className={styles.findingTitle}>{s.title}</div>
                <div className={styles.findingDetail}>
                  {formatBytes(s.latestBytes)} now · {growthSummary(s)}
                </div>
              </div>
              <div className={styles.findingGo}>
                <Sparkline values={s.spark} color={si.color} />
              </div>
            </div>
          );
        })}
      </div>
    </Card>
  );
};
