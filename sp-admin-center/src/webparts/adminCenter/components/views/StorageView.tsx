import * as React from 'react';
import { DefaultButton, DetailsList, DetailsListLayoutMode, Dropdown, MessageBar, MessageBarType, SelectionMode } from '@fluentui/react';
import { IFileRow } from '../../models';
import { downloadCsv } from '../../services/exportCsv';
import { formatBytes, formatCompact, formatDate } from '../../services/format';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { useLoader } from '../shared/hooks';
import { BarList, Card, Donut, Empty, ErrorBar, Kpi, Loading, ViewHeader } from '../shared/ui';

const SIZES = [
  { key: 10 * 1048576, text: '10 MB and larger' },
  { key: 50 * 1048576, text: '50 MB and larger' },
  { key: 100 * 1048576, text: '100 MB and larger' },
  { key: 500 * 1048576, text: '500 MB and larger' },
  { key: 1024 * 1048576, text: '1 GB and larger' }
];

/**
 * Storage insights for the whole site collection. Every number is one search
 * request (an aggregate from the index), never a walk over files, so it costs
 * the same on 10 files or 10 million.
 */
export const StorageView: React.FC = () => {
  const ctx = useAdmin();
  const { api, search, target } = ctx;
  const [minSize, setMinSize] = React.useState(SIZES[1].key);

  const overview = useLoader(async () => {
    const site = await api.getSite(target.webUrl);
    const [types, stale] = await Promise.all([search.fileTypes(target.siteUrl), search.staleCounts(target.siteUrl, [1, 2, 3])]);
    return { site, types, stale };
  }, [target.siteUrl]);

  const big = useLoader(() => search.largestFiles(target.siteUrl, minSize, 50), [target.siteUrl, minSize]);

  if (overview.loading && !overview.data) {
    return <Loading text="Asking the search index…" />;
  }
  if (overview.error || !overview.data) {
    return <ErrorBar error={overview.error || 'Nothing loaded.'} onRetry={overview.reload} />;
  }
  const { site, types, stale } = overview.data;
  const quota = site.storageFraction > 0 ? site.storageBytes / site.storageFraction : 0;
  const topTypes = types.types.slice(0, 7);
  const other = types.types.slice(7).reduce((s, t) => s + t.count, 0);
  const slices = topTypes.map((t) => ({ label: t.type.toUpperCase() || 'none', value: t.count }));
  if (other > 0) {
    slices.push({ label: 'Other', value: other });
  }

  return (
    <div className={styles.view}>
      <ViewHeader title="Storage & content insights" hint={`Whole site collection: ${target.siteUrl}`}>
        <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={() => { overview.reload(); big.reload(); }} disabled={overview.loading || big.loading}>
          Refresh
        </DefaultButton>
      </ViewHeader>

      <MessageBar messageBarType={MessageBarType.info}>
        File figures come from the search index: they cover indexed files you can see, and may lag behind very recent changes. Storage used is SharePoint&rsquo;s own figure and includes version history and the recycle bin.
      </MessageBar>

      <div className={styles.kpiRow}>
        <Kpi label="Storage used" value={formatBytes(site.storageBytes)} sub={quota ? `of ${formatBytes(quota)}` : undefined} />
        <Kpi label="Indexed files" value={formatCompact(types.total)} sub="in the search index" />
        <Kpi label="File types" value={String(types.types.length)} sub={types.topTenOnly ? 'top 10 only' : 'distinct extensions'} />
        {stale.map((s) => (
          <Kpi key={s.years} label={`Untouched ${s.years}+ yr`} value={formatCompact(s.count)} sub={types.total ? `${Math.round((s.count / Math.max(1, types.total)) * 100)}% of files` : undefined} />
        ))}
      </div>

      <div className={styles.grid}>
        <Card title="Files by type">{slices.length ? <Donut slices={slices} centre={formatCompact(types.total)} centreSub="files" /> : <Empty text="No indexed files found." />}</Card>
        <Card title="Most common extensions">
          {topTypes.length ? <BarList rows={types.types.slice(0, 10).map((t) => ({ label: '.' + t.type, value: t.count }))} /> : <Empty text="Nothing to show." />}
        </Card>
      </div>

      <Card
        title="Largest files"
        right={
          <div className={styles.actions}>
            <Dropdown
              ariaLabel="Minimum size"
              selectedKey={minSize}
              onChange={(_, o) => setMinSize(Number(o ? o.key : SIZES[1].key))}
              options={SIZES}
              styles={{ root: { minWidth: 170 } }}
            />
            <DefaultButton
              iconProps={{ iconName: 'Download' }}
              disabled={!big.data || big.data.length === 0}
              onClick={() => downloadCsv('largest-files.csv', ['Name', 'Size (bytes)', 'Type', 'Modified', 'Author', 'URL'], (big.data || []).map((f) => [f.title, f.size, f.extension, f.modified, f.author, f.path]))}
            >
              Export
            </DefaultButton>
          </div>
        }
      >
        {big.loading && !big.data && <Loading />}
        {big.error && <ErrorBar error={big.error} onRetry={big.reload} />}
        {big.data && big.data.length === 0 && <Empty text="No files that large. Nice." />}
        {big.data && big.data.length > 0 && (
          <DetailsList
            items={big.data}
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
            getKey={(f: IFileRow) => f.path}
            columns={[
              {
                key: 'name',
                name: 'File',
                minWidth: 220,
                isResizable: true,
                onRender: (f: IFileRow) => (
                  <a className={styles.link} href={f.path} target="_blank" rel="noopener noreferrer">
                    {f.title}
                  </a>
                )
              },
              { key: 'size', name: 'Size', minWidth: 80, onRender: (f: IFileRow) => <strong>{formatBytes(f.size)}</strong> },
              { key: 'ext', name: 'Type', minWidth: 60, onRender: (f: IFileRow) => f.extension },
              { key: 'mod', name: 'Modified', minWidth: 100, onRender: (f: IFileRow) => formatDate(f.modified) },
              { key: 'by', name: 'Author', minWidth: 120, onRender: (f: IFileRow) => f.author }
            ]}
          />
        )}
        <p className={styles.muted} style={{ marginBottom: 0 }}>
          Largest current versions only. Older versions are included in the site&rsquo;s storage figure but not listed here.
        </p>
      </Card>
    </div>
  );
};
