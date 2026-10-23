import * as React from 'react';
import { DefaultButton, DetailsList, DetailsListLayoutMode, PrimaryButton, SelectionMode, TextField } from '@fluentui/react';
import { IActionLogEntry } from '../../models';
import { downloadCsv } from '../../services/exportCsv';
import { formatDateTime } from '../../services/format';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { useLoader } from '../shared/hooks';
import { Card, Empty, ErrorBar, Loading, Pill, ViewHeader } from '../shared/ui';

export const SettingsView: React.FC<{ stats: { requests: number; retries: number; throttled: number; cacheHits: number } }> = ({ stats }) => {
  const ctx = useAdmin();
  const { api, target } = ctx;
  const web = useLoader(() => api.getWeb(target.webUrl), [target.webUrl]);
  const [title, setTitle] = React.useState<string | undefined>();
  const [description, setDescription] = React.useState<string | undefined>();
  const [saving, setSaving] = React.useState(false);

  if (web.loading && !web.data) {
    return <Loading />;
  }
  if (web.error || !web.data) {
    return <ErrorBar error={web.error || 'Nothing loaded.'} onRetry={web.reload} />;
  }
  const w = web.data;
  const t = title === undefined ? w.title : title;
  const d = description === undefined ? w.description : description;
  const dirty = t !== w.title || d !== w.description;

  const save = async (): Promise<void> => {
    if (!t.trim()) {
      ctx.notify('The title cannot be empty.', 'error');
      return;
    }
    setSaving(true);
    try {
      await api.updateWeb(target.webUrl, { Title: t.trim(), Description: d });
      ctx.log('Update site title / description', w.url, true);
      ctx.notify('Site details saved.', 'success');
      ctx.setTarget({ ...target, title: t.trim() });
      setTitle(undefined);
      setDescription(undefined);
      web.reload();
    } catch (e) {
      ctx.log('Update site title / description', w.url, false, (e as Error).message);
      ctx.notify((e as Error).message, 'error');
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className={styles.view}>
      <ViewHeader title="Settings & audit log" hint="Edit basic site details, review what was changed from this tool, and see how it protects SharePoint." />

      <div className={styles.grid}>
        <Card title="Site details">
          <TextField label="Title" value={t} onChange={(_, v) => setTitle(v || '')} maxLength={255} />
          <TextField label="Description" value={d} onChange={(_, v) => setDescription(v || '')} multiline rows={3} maxLength={512} />
          <div className={styles.actions} style={{ marginTop: 10 }}>
            <PrimaryButton disabled={!dirty || saving} onClick={save}>
              Save changes
            </PrimaryButton>
            <DefaultButton disabled={!dirty || saving} onClick={() => { setTitle(undefined); setDescription(undefined); }}>
              Discard
            </DefaultButton>
          </div>
        </Card>

        <Card title="How this tool protects SharePoint">
          <ul style={{ margin: 0, paddingLeft: 18, fontSize: 13, lineHeight: 1.6 }}>
            <li>Runs as you, through the page&rsquo;s own sign-in. No app registration, no client ID or secret.</li>
            <li>Nothing runs in the background. Every request starts from a click or a page you open.</li>
            <li>At most 2 requests at a time, spaced ~150 ms apart.</li>
            <li>On a 429 / 503 the whole queue pauses and honours <em>Retry-After</em>; rate-limit headers slow it down early.</li>
            <li>Counts and top-N lists come from the search index in a single request, never from walking files.</li>
            <li>Results are cached for 5 minutes, so switching views repeats nothing.</li>
            <li>Bulk actions are capped and run one item at a time.</li>
          </ul>
          <p className={styles.muted} style={{ marginBottom: 0 }}>
            This session: {stats.requests} requests · {stats.cacheHits} answered from cache · {stats.throttled} throttled · {stats.retries} retries.
          </p>
        </Card>
      </div>

      <Card
        title="Actions taken from this tool (this session)"
        right={
          <DefaultButton
            iconProps={{ iconName: 'Download' }}
            disabled={ctx.actionLog.length === 0}
            onClick={() => downloadCsv('admin-actions.csv', ['Time', 'Action', 'Target', 'Result', 'Message'], ctx.actionLog.map((a) => [a.at, a.action, a.target, a.ok ? 'OK' : 'Failed', a.message]))}
          >
            Export
          </DefaultButton>
        }
      >
        {ctx.actionLog.length === 0 ? (
          <Empty text="No changes made yet. Everything you change here is listed for your records." />
        ) : (
          <DetailsList
            items={ctx.actionLog.slice().reverse()}
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
            columns={[
              { key: 'at', name: 'When', minWidth: 140, onRender: (a: IActionLogEntry) => formatDateTime(a.at) },
              { key: 'act', name: 'Action', minWidth: 160, onRender: (a: IActionLogEntry) => a.action },
              { key: 'tg', name: 'Target', minWidth: 200, isResizable: true, onRender: (a: IActionLogEntry) => a.target },
              { key: 'res', name: 'Result', minWidth: 160, onRender: (a: IActionLogEntry) => (a.ok ? <Pill kind="good">OK</Pill> : <Pill kind="critical">{(a.message || 'Failed').substring(0, 60)}</Pill>) }
            ]}
          />
        )}
        <p className={styles.muted} style={{ marginBottom: 0 }}>
          SharePoint&rsquo;s own audit log (Microsoft Purview) is the system of record; this list is a convenience for the current browser session.
        </p>
      </Card>
    </div>
  );
};
