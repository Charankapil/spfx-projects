import * as React from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  MessageBar,
  MessageBarType,
  SearchBox,
  SelectionMode,
  Spinner
} from '@fluentui/react';

import { ESignService, IRequestWithStatus } from '../services/ESignService';
import { sameEmail, slotForUser } from '../services/status';
import { errorText, formatDate, StatusBadge } from './shared';
import styles from './ESign.module.scss';

export interface IDashboardProps {
  service: ESignService;
  onOpen: (requestId: string) => void;
}

type Filter = 'action' | 'waiting' | 'completed' | 'closed' | 'all';

interface IRow {
  key: string;
  item: IRequestWithStatus;
  actionRequired: boolean;
  involved: boolean;
}

/** Reminders run at most once per page load, however often the dashboard re-renders. */
let remindersChecked = false;

export function Dashboard(props: IDashboardProps): JSX.Element {
  const { service } = props;
  const [items, setItems] = React.useState<IRequestWithStatus[]>();
  const [error, setError] = React.useState<string>();
  const [info, setInfo] = React.useState<string>();
  const [filter, setFilter] = React.useState<Filter>('action');
  const [query, setQuery] = React.useState('');
  const me = service.me;

  React.useEffect(() => {
    (async () => {
      const loaded = await service.loadAll();
      setItems(loaded);
      if (!loaded.some(i => slotForUser(i.status, me.email))) {
        setFilter('all');
      }
      if (!remindersChecked) {
        remindersChecked = true;
        const sent = await service.sendDueReminders(loaded).catch(() => 0);
        if (sent) {
          setInfo(`Sent scheduled reminders for ${sent} of your request(s).`);
        }
      }
    })().catch(e => setError(errorText(e)));
  }, [service, me.email]);

  const rows: IRow[] = React.useMemo(
    () =>
      (items || [])
        .map(item => ({
          key: item.request.requestId,
          item,
          actionRequired: !!slotForUser(item.status, me.email),
          involved:
            sameEmail(item.request.authorEmail, me.email) ||
            item.status.slots.some(s => sameEmail(s.current.email, me.email) || sameEmail(s.original.email, me.email)) ||
            item.status.ccs.some(c => sameEmail(c.email, me.email))
        }))
        .filter(r => r.involved),
    [items, me.email]
  );

  const counts = {
    action: rows.filter(r => r.actionRequired).length,
    waiting: rows.filter(r => !r.actionRequired && r.item.status.state === 'InProgress').length,
    completed: rows.filter(r => r.item.status.state === 'Completed').length,
    closed: rows.filter(r => ['Declined', 'Cancelled', 'Expired'].indexOf(r.item.status.state) >= 0).length,
    all: rows.length
  };

  const q = query.trim().toLowerCase();
  const visible = rows
    .filter(r => {
      const state = r.item.status.state;
      switch (filter) {
        case 'action':
          return r.actionRequired;
        case 'waiting':
          return !r.actionRequired && state === 'InProgress';
        case 'completed':
          return state === 'Completed';
        case 'closed':
          return state === 'Declined' || state === 'Cancelled' || state === 'Expired';
        default:
          return true;
      }
    })
    .filter(r => {
      if (!q) {
        return true;
      }
      const p = r.item.request.payload;
      return [p.title, p.documentName, p.senderName, p.senderEmail]
        .concat(p.recipients.map(x => `${x.name} ${x.email}`))
        .some(text => (text || '').toLowerCase().indexOf(q) >= 0);
    });

  const columns: IColumn[] = [
    {
      key: 'title',
      name: 'Request',
      minWidth: 200,
      isMultiline: true,
      onRender: (r: IRow) => (
        <div>
          <strong>{r.item.request.payload.title}</strong>
          <div className={styles.muted}>{r.item.request.payload.documentName}</div>
        </div>
      )
    },
    {
      key: 'status',
      name: 'Status',
      minWidth: 110,
      maxWidth: 130,
      onRender: (r: IRow) => <StatusBadge state={r.item.status.state} actionRequired={r.actionRequired} />
    },
    {
      key: 'progress',
      name: 'Signed',
      minWidth: 90,
      maxWidth: 110,
      onRender: (r: IRow) => (
        <div>
          <div className={styles.muted}>
            {r.item.status.signedCount} of {r.item.status.signerCount}
          </div>
          <div className={styles.progress}>
            <div className={styles.progressBar} style={{ width: `${(r.item.status.signedCount / Math.max(1, r.item.status.signerCount)) * 100}%` }} />
          </div>
        </div>
      )
    },
    {
      key: 'waiting',
      name: 'Waiting on',
      minWidth: 140,
      isMultiline: true,
      onRender: (r: IRow) => <span>{r.item.status.current.map(s => s.current.name).join(', ') || '-'}</span>
    },
    {
      key: 'sender',
      name: 'From',
      minWidth: 120,
      maxWidth: 180,
      onRender: (r: IRow) => <span>{sameEmail(r.item.request.authorEmail, me.email) ? 'You' : r.item.request.payload.senderName}</span>
    },
    {
      key: 'sent',
      name: 'Sent',
      minWidth: 90,
      maxWidth: 110,
      onRender: (r: IRow) => <span>{formatDate(r.item.request.payload.sentOn)}</span>
    },
    {
      key: 'expires',
      name: 'Expires',
      minWidth: 90,
      maxWidth: 110,
      onRender: (r: IRow) => <span>{r.item.status.state === 'InProgress' ? formatDate(r.item.request.payload.expiresOn) : '-'}</span>
    }
  ];

  const kpi = (key: Filter, label: string, value: number): JSX.Element => (
    <button key={key} className={`${styles.kpi} ${filter === key ? styles.kpiActive : ''}`} onClick={() => setFilter(key)} aria-pressed={filter === key}>
      <div className={styles.kpiValue}>{value}</div>
      <div className={styles.kpiLabel}>{label}</div>
    </button>
  );

  if (error) {
    return <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>;
  }
  if (!items) {
    return <Spinner label="Loading your requests..." styles={{ root: { padding: 32 } }} />;
  }

  return (
    <div>
      {info && (
        <MessageBar messageBarType={MessageBarType.info} onDismiss={() => setInfo(undefined)}>
          {info}
        </MessageBar>
      )}
      <div className={styles.kpis}>
        {kpi('action', 'Action required', counts.action)}
        {kpi('waiting', 'Waiting for others', counts.waiting)}
        {kpi('completed', 'Completed', counts.completed)}
        {kpi('closed', 'Declined, cancelled or expired', counts.closed)}
        {kpi('all', 'All my requests', counts.all)}
      </div>
      <div className={styles.card}>
        <div className={styles.toolbar}>
          <SearchBox placeholder="Search by title, document or person" value={query} onChange={(_, v) => setQuery(v || '')} styles={{ root: { width: 320, maxWidth: '100%' } }} />
        </div>
        {visible.length ? (
          <DetailsList
            items={visible}
            columns={columns}
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
            onItemInvoked={(r: IRow) => props.onOpen(r.key)}
            onRenderRow={(rowProps, render) =>
              rowProps && render ? (
                <div style={{ cursor: 'pointer' }} onClick={() => props.onOpen((rowProps.item as IRow).key)}>
                  {render(rowProps)}
                </div>
              ) : null
            }
          />
        ) : (
          <div className={styles.empty}>{filter === 'action' ? 'Nothing needs your signature right now.' : 'No requests here.'}</div>
        )}
      </div>
    </div>
  );
}
