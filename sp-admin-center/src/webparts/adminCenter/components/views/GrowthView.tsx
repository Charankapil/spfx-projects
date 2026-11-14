import * as React from 'react';
import { DefaultButton, Pivot, PivotItem, TextField } from '@fluentui/react';
import { analyzeAll, anomaliesOf, DEFAULT_SETTINGS, IGrowthDoc, IGrowthSettings } from '../../services/GrowthEngine';
import { ITenantIndex } from '../../services/TenantStore';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { useLoader } from '../shared/hooks';
import { Card, ErrorBar, Loading, ViewHeader } from '../shared/ui';
import { TenantTab } from './TenantTab';
import { WatchlistTab } from './WatchlistTab';

export { growthSummary } from './growthShared';

/**
 * Storage growth. Two sources feed the same alert rules and thresholds:
 *  - Tenant import: every site, from a CSV your flow writes (one file read per view, any number of sites).
 *  - Watchlist: live per-site snapshots for up to 500 chosen sites (one small request per site).
 */
export const GrowthView: React.FC = () => {
  const ctx = useAdmin();
  const [tab, setTab] = React.useState<string | undefined>();
  const loader = useLoader(async () => {
    const doc = await ctx.growth.load();
    let index: ITenantIndex | undefined;
    let indexError: string | undefined;
    try {
      index = await ctx.tenant.loadIndex();
    } catch (e) {
      indexError = (e as Error).message;
    }
    return { doc, index, indexError };
  }, []);

  const data = loader.data;
  const count = React.useMemo(() => {
    if (!data) {
      return undefined;
    }
    return anomaliesOf(analyzeAll(data.doc)).length + (data.index && data.index.summary ? data.index.summary.anomalies.length : 0);
  }, [data]);
  React.useEffect(() => {
    if (count !== undefined) {
      ctx.setGrowthBadge(count);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [count]);

  if (loader.loading && !data) {
    return <Loading text="Loading the storage history…" />;
  }
  if (loader.error || !data) {
    return <ErrorBar error={loader.error || 'Nothing loaded.'} onRetry={loader.reload} />;
  }
  const hasTenant = !!(data.index && data.index.snapshots.length);
  const hasWatch = Object.keys(data.doc.sites).length > 0;
  const selected = tab || (hasTenant || !hasWatch ? 'tenant' : 'watchlist');

  return (
    <div className={styles.view}>
      <ViewHeader title="Storage growth" hint="Tracks how each site's storage changes over time and flags sites that grow unusually fast.">
        <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={loader.reload} disabled={loader.loading}>
          Refresh
        </DefaultButton>
      </ViewHeader>

      <Pivot selectedKey={selected} onLinkClick={(i) => setTab(i && i.props.itemKey)} aria-label="Storage growth sources">
        <PivotItem headerText="All sites (CSV import)" itemKey="tenant" itemCount={data.index && data.index.summary ? data.index.summary.anomalies.length || undefined : undefined}>
          <div style={{ paddingTop: 12 }}>
            <TenantTab doc={data.doc} index={data.index} indexError={data.indexError} reload={loader.reload} />
          </div>
        </PivotItem>
        <PivotItem headerText="Watchlist (live, up to 500)" itemKey="watchlist">
          <div style={{ paddingTop: 12 }}>
            <WatchlistTab doc={data.doc} reload={loader.reload} />
          </div>
        </PivotItem>
      </Pivot>

      <ThresholdsCard doc={data.doc} onSaved={loader.reload} />
    </div>
  );
};

const ThresholdsCard: React.FC<{ doc: IGrowthDoc; onSaved: () => void }> = ({ doc, onSaved }) => {
  const ctx = useAdmin();
  const [draft, setDraft] = React.useState<{ [k in keyof IGrowthSettings]?: string }>({});
  const num = (k: keyof IGrowthSettings): string => (draft[k] !== undefined ? (draft[k] as string) : String(doc.settings[k]));
  const dirty = Object.keys(draft).length > 0;

  const save = async (): Promise<void> => {
    const next = { ...doc.settings };
    (Object.keys(draft) as Array<keyof IGrowthSettings>).forEach((k) => {
      const v = Number(draft[k]);
      if (isFinite(v) && v > 0) {
        next[k] = v;
      }
    });
    try {
      await ctx.growth.saveSettings(next);
      ctx.log('Change growth thresholds', JSON.stringify(next), true);
      ctx.notify('Thresholds saved. Use "Recalculate with current thresholds" on the CSV import tab to apply them to imported data.', 'success');
      setDraft({});
      onSaved();
    } catch (e) {
      ctx.notify((e as Error).message, 'error');
    }
  };

  return (
    <Card
      title="When to raise an alert"
      right={
        <div className={styles.actions}>
          <DefaultButton disabled={!dirty} onClick={save}>
            Save thresholds
          </DefaultButton>
          <DefaultButton onClick={() => setDraft(Object.keys(DEFAULT_SETTINGS).reduce((a, k) => ({ ...a, [k]: String(DEFAULT_SETTINGS[k as keyof IGrowthSettings]) }), {}))}>Reset to defaults</DefaultButton>
        </div>
      }
    >
      <div className={styles.filters}>
        <TextField className={styles.field} label="Growth of at least (% per week)" type="number" min={1} value={num('pctPerWeek')} onChange={(_, v) => setDraft({ ...draft, pctPerWeek: v || '' })} />
        <TextField className={styles.field} label="…and at least (GB per week)" type="number" min={1} value={num('gbPerWeek')} onChange={(_, v) => setDraft({ ...draft, gbPerWeek: v || '' })} />
        <TextField className={styles.field} label="Always alert above (GB per week)" type="number" min={1} value={num('hugeGbPerWeek')} onChange={(_, v) => setDraft({ ...draft, hugeGbPerWeek: v || '' })} />
        <TextField className={styles.field} label="Quota full within (days)" type="number" min={1} value={num('forecastDays')} onChange={(_, v) => setDraft({ ...draft, forecastDays: v || '' })} />
      </div>
      <p className={styles.muted} style={{ marginBottom: 0 }}>
        Rates are scaled to a week from the snapshots you have, so uneven gaps are fine. A site is also flagged when its latest daily growth is far above its own usual pattern. Thresholds are shared by every admin and stored in <em>Site Assets</em> on this site.
      </p>
    </Card>
  );
};
