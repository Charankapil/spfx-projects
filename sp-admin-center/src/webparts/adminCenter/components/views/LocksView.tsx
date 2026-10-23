import * as React from 'react';
import {
  ChoiceGroup,
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  Dropdown,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  SelectionMode,
  TextField,
  Toggle
} from '@fluentui/react';
import { LockScope } from '../../../../shared/readOnlyLocks';
import { IListInfo } from '../../models';
import { formatDateTime } from '../../services/format';
import { ILockRecord, ILockTarget, IPreview, LockService, MAX_MESSAGE } from '../../services/LockService';
import styles from '../AdminCenter.module.scss';
import { useAdmin } from '../shared/context';
import { useLoader } from '../shared/hooks';
import { Card, Empty, ErrorBar, Loading, Pill, ViewHeader } from '../shared/ui';

const DANGER = { root: { background: '#c4314b', borderColor: '#c4314b' }, rootHovered: { background: '#a4262c', borderColor: '#a4262c' } };
const scopeName = (s: LockScope): string => (s === 'web' ? 'Site' : s === 'list' ? 'Library / list' : 'Folder');

function serverRelative(url: string): string {
  const m = /^https:\/\/[^/]+(\/[^?#]*)?/i.exec(url.replace(/\/+$/, ''));
  return m && m[1] ? decodeURIComponent(m[1].replace(/\/+$/, '')) : '';
}

/**
 * Read-only locks for the site you are managing: lock the whole site, one library or
 * list, or one folder, with a banner message; unlock restores the exact permissions.
 */
export const LocksView: React.FC = () => {
  const ctx = useAdmin();
  const { target, api } = ctx;
  const service = React.useMemo(() => new LockService(ctx.client), [ctx.client]);
  const data = useLoader(async () => {
    const [locks, lists, isAdmin] = await Promise.all([service.loadLocks(target.webUrl), api.getLists(target.webUrl), service.isSiteAdmin(target.webUrl).catch(() => false)]);
    return { locks, lists, isAdmin };
  }, [target.webUrl]);
  const [busy, setBusy] = React.useState('');

  if (data.loading && !data.data) {
    return <Loading text="Loading locks…" />;
  }
  if (data.error || !data.data) {
    return <ErrorBar error={data.error || 'Nothing loaded.'} onRetry={data.reload} />;
  }
  const { locks, lists, isAdmin } = data.data;

  const unlock = async (l: ILockRecord): Promise<void> => {
    const ok = await ctx.confirm({
      title: 'Unlock',
      message: `Unlock "${l.title}"? Everyone gets back the permissions they had before it was locked, and the banner is removed.`,
      confirmText: 'Unlock'
    });
    if (!ok) {
      return;
    }
    setBusy('Unlocking…');
    try {
      await service.unlock(target.webUrl, l.id, setBusy);
      ctx.log('Unlock (read-only)', `${scopeName(l.scope)}: ${l.path}`, true);
      ctx.notify(`"${l.title}" is editable again.`, 'success');
      data.reload();
    } catch (e) {
      ctx.log('Unlock (read-only)', l.path, false, (e as Error).message);
      ctx.notify((e as Error).message, 'error');
    } finally {
      setBusy('');
    }
  };

  return (
    <div className={styles.view}>
      <ViewHeader title="Read-only locks" hint={`Make ${target.title}, one of its libraries or a folder read-only, with a message shown in a banner.`}>
        <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={data.reload} disabled={!!busy}>
          Refresh
        </DefaultButton>
      </ViewHeader>

      {busy && <Loading text={busy} />}

      <Card title={`Current locks on ${target.title}`}>
        {locks.length === 0 ? (
          <Empty text="Nothing on this site is locked." />
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            {locks.map((l) => (
              <LockCard key={l.id} lock={l} service={service} onUnlock={() => unlock(l)} onChanged={data.reload} disabled={!!busy} />
            ))}
          </div>
        )}
      </Card>

      <LockForm lists={lists} isAdmin={isAdmin} service={service} onLocked={data.reload} />
    </div>
  );
};

const LockCard: React.FC<{ lock: ILockRecord; service: LockService; onUnlock: () => void; onChanged: () => void; disabled: boolean }> = ({ lock, service, onUnlock, onChanged, disabled }) => {
  const ctx = useAdmin();
  const [editing, setEditing] = React.useState(false);
  const [message, setMessage] = React.useState(lock.message);
  const save = async (): Promise<void> => {
    try {
      await service.updateMessage(ctx.target.webUrl, lock.id, message);
      ctx.log('Change read-only banner message', lock.path, true);
      setEditing(false);
      onChanged();
    } catch (e) {
      ctx.notify((e as Error).message, 'error');
    }
  };
  return (
    <div className={styles.finding}>
      <span className={styles.statIcon} style={{ marginBottom: 0, background: '#fdf1dc', color: '#8a4b05' }}>🔒</span>
      <div style={{ minWidth: 0, flex: 1 }}>
        <div className={styles.findingTitle}>
          {lock.title} <Pill>{scopeName(lock.scope)}</Pill> {lock.status === 'locking' && <Pill kind="critical">Interrupted: unlock to restore</Pill>}
        </div>
        <div className={styles.findingDetail}>{lock.path}</div>
        {editing ? (
          <div style={{ marginTop: 8 }}>
            <TextField multiline rows={3} value={message} maxLength={MAX_MESSAGE} onChange={(_, v) => setMessage(v || '')} ariaLabel="Banner message" />
            <div className={styles.actions} style={{ marginTop: 6 }}>
              <PrimaryButton onClick={save}>Save message</PrimaryButton>
              <DefaultButton onClick={() => setEditing(false)}>Cancel</DefaultButton>
            </div>
          </div>
        ) : (
          <div style={{ marginTop: 6, fontSize: 13, padding: '6px 10px', background: '#fff8e1', borderRadius: 6, whiteSpace: 'pre-wrap' }}>{lock.message || <em>No message</em>}</div>
        )}
        <div className={styles.muted} style={{ marginTop: 6, fontSize: 12 }}>
          Locked {formatDateTime(new Date(lock.at * 1000))} by {lock.by || 'unknown'} · {lock.objects.length} object{lock.objects.length === 1 ? '' : 's'} · {lock.keepOwners ? 'owners kept edit rights' : 'owners set to Read'}
        </div>
      </div>
      <div className={styles.findingGo} style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
        <PrimaryButton iconProps={{ iconName: 'Unlock' }} onClick={onUnlock} disabled={disabled}>
          Unlock
        </PrimaryButton>
        <DefaultButton onClick={() => setEditing(!editing)} disabled={disabled}>
          Edit message
        </DefaultButton>
      </div>
    </div>
  );
};

const LockForm: React.FC<{ lists: IListInfo[]; isAdmin: boolean; service: LockService; onLocked: () => void }> = ({ lists, isAdmin, service, onLocked }) => {
  const ctx = useAdmin();
  const { target } = ctx;
  const [scope, setScope] = React.useState<LockScope>('list');
  const [listId, setListId] = React.useState<string | undefined>();
  const [folder, setFolder] = React.useState<string | undefined>();
  const [subfolders, setSubfolders] = React.useState<Array<{ name: string; path: string }>>([]);
  const [message, setMessage] = React.useState('This content has been made read-only. Contact the site owner if you need to make changes.');
  const [keepOwners, setKeepOwners] = React.useState(true);
  const [includeUnique, setIncludeUnique] = React.useState(true);
  const [preview, setPreview] = React.useState<IPreview | undefined>();
  const [busy, setBusy] = React.useState('');
  const [error, setError] = React.useState<string | undefined>();

  const visible = lists.filter((l) => !l.hidden && !l.isSystem);
  const libraries = visible.filter((l) => l.kind !== 'Other');
  const uniqueLists = visible.filter((l) => l.uniquePermissions).map((l) => ({ id: l.id, title: l.title }));
  const list = lists.filter((l) => l.id === listId)[0];

  const browse = async (path: string): Promise<void> => {
    setFolder(path);
    setPreview(undefined);
    try {
      setSubfolders(await service.listFolders(target.webUrl, path));
    } catch (e) {
      setSubfolders([]);
      setError((e as Error).message);
    }
  };

  const lockTarget = (): ILockTarget | undefined => {
    if (scope === 'web') {
      return { scope, title: target.title, path: serverRelative(target.webUrl) };
    }
    if (!list) {
      return undefined;
    }
    if (scope === 'list') {
      return { scope, title: list.title, path: list.rootFolder, listId: list.id };
    }
    if (!folder || folder === list.rootFolder) {
      return undefined;
    }
    return { scope, title: folder.substring(folder.lastIndexOf('/') + 1), path: folder, folderPath: folder };
  };
  const t = lockTarget();
  const options = { keepOwners: isAdmin ? keepOwners : true, includeUniqueLists: includeUnique };

  const runPreview = async (): Promise<void> => {
    if (!t) {
      return;
    }
    setBusy('Checking permissions…');
    setError(undefined);
    try {
      setPreview(await service.preview(target.webUrl, t, options, uniqueLists));
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy('');
    }
  };

  const runLock = async (): Promise<void> => {
    if (!t || !preview) {
      return;
    }
    const n = preview.objects.reduce((s, o) => s + o.changes.length, 0);
    const ok = await ctx.confirm({
      title: `Make "${t.title}" read-only`,
      message: `${n} permission${n === 1 ? '' : 's'} will be set to Read on ${preview.objects.length} object${preview.objects.length === 1 ? '' : 's'}, and the banner will show your message. The current permissions are saved first, so Unlock restores them exactly.`,
      confirmText: 'Lock',
      danger: true
    });
    if (!ok) {
      return;
    }
    setBusy('Locking…');
    setError(undefined);
    try {
      await service.lock(target.webUrl, t, message, options, uniqueLists, ctx.currentUser, setBusy);
      ctx.log('Lock (read-only)', `${scopeName(t.scope)}: ${t.path}`, true);
      ctx.notify(`"${t.title}" is now read-only.`, 'success');
      setPreview(undefined);
      onLocked();
    } catch (e) {
      ctx.log('Lock (read-only)', t.path, false, (e as Error).message);
      setError(`${(e as Error).message} If the lock was partly applied it is listed above; Unlock restores what was changed.`);
      onLocked();
    } finally {
      setBusy('');
    }
  };

  const changeRows = preview ? preview.objects.reduce<Array<{ key: string; object: string; who: string; before: string; after: string }>>((rows, o) => rows.concat(o.changes.map((c) => ({ key: `${o.object.title}-${c.principalId}`, object: o.object.title, who: c.title, before: c.before.join(', '), after: c.after.join(', ') }))), []) : [];

  return (
    <Card title="Lock something">
      <ChoiceGroup
        selectedKey={scope}
        onChange={(_, o) => {
          setScope((o ? o.key : 'list') as LockScope);
          setPreview(undefined);
        }}
        options={[
          { key: 'web', text: `The whole site (${target.title})` },
          { key: 'list', text: 'One library or list' },
          { key: 'folder', text: 'One folder in a library' }
        ]}
        label="What should be read-only?"
      />

      {scope !== 'web' && (
        <div className={styles.filters} style={{ marginTop: 10 }}>
          <Dropdown
            className={styles.field}
            label={scope === 'folder' ? 'Library' : 'Library or list'}
            selectedKey={listId}
            placeholder="Choose…"
            options={(scope === 'folder' ? libraries.filter((l) => l.kind === 'Library') : libraries).map((l) => ({ key: l.id, text: l.title }))}
            onChange={(_, o) => {
              const id = o ? String(o.key) : undefined;
              setListId(id);
              setPreview(undefined);
              const chosen = lists.filter((l) => l.id === id)[0];
              if (scope === 'folder' && chosen) {
                browse(chosen.rootFolder);
              }
            }}
          />
        </div>
      )}

      {scope === 'folder' && list && folder && (
        <div style={{ marginTop: 10 }}>
          <div className={styles.muted}>
            Folder: <strong>{folder.substring(list.rootFolder.length) || '(choose a folder below)'}</strong>
            {folder !== list.rootFolder && (
              <button type="button" className={styles.link} style={{ marginLeft: 10 }} onClick={() => browse(folder.substring(0, folder.lastIndexOf('/')))}>
                Up one level
              </button>
            )}
          </div>
          <ul className={styles.memberList}>
            {subfolders.length === 0 && <li className={styles.muted}>No subfolders here.</li>}
            {subfolders.slice(0, 200).map((f) => (
              <li key={f.path} className={styles.memberRow}>
                <span>📁 {f.name}</span>
                <button type="button" className={styles.link} onClick={() => browse(f.path)}>
                  Open
                </button>
              </li>
            ))}
          </ul>
        </div>
      )}

      <TextField
        label={`Banner message (shown at the top of the locked pages, up to ${MAX_MESSAGE} characters)`}
        multiline
        rows={3}
        maxLength={MAX_MESSAGE}
        value={message}
        onChange={(_, v) => setMessage(v || '')}
        styles={{ root: { marginTop: 12 } }}
      />

      <div className={styles.filters} style={{ marginTop: 10 }}>
        <Toggle
          label={isAdmin ? 'Site owners keep edit rights' : 'Site owners keep edit rights (only a site collection admin can change this)'}
          inlineLabel
          checked={isAdmin ? keepOwners : true}
          disabled={!isAdmin}
          onChange={(_, c) => {
            setKeepOwners(!!c);
            setPreview(undefined);
          }}
        />
        {scope === 'web' && (
          <Toggle
            label={`Also lock lists and libraries with their own permissions (${uniqueLists.length})`}
            inlineLabel
            checked={includeUnique}
            onChange={(_, c) => {
              setIncludeUnique(!!c);
              setPreview(undefined);
            }}
          />
        )}
      </div>

      <div className={styles.actions} style={{ marginTop: 12 }}>
        <DefaultButton iconProps={{ iconName: 'Preview' }} onClick={runPreview} disabled={!t || !!busy}>
          Preview changes
        </DefaultButton>
        <PrimaryButton
          iconProps={{ iconName: 'Lock' }}
          onClick={runLock}
          disabled={!preview || !!busy || !message.trim()}
          styles={!preview || !!busy || !message.trim() ? undefined : DANGER}
        >
          Lock
        </PrimaryButton>
        {!t && <span className={styles.muted}>{scope === 'folder' ? 'Choose a library, then open the folder to lock.' : scope === 'list' ? 'Choose a library or list.' : ''}</span>}
      </div>

      {busy && <Loading text={busy} />}
      {error && <ErrorBar error={error} />}

      {preview && (
        <div style={{ marginTop: 12 }}>
          {preview.warnings.map((w) => (
            <MessageBar key={w} messageBarType={MessageBarType.info} isMultiline>
              {w}
            </MessageBar>
          ))}
          {changeRows.length === 0 ? (
            <Empty text="Nobody has edit rights here (outside the people kept), so only the banner would be added." />
          ) : (
            <div className={styles.tableWrap} style={{ marginTop: 8 }}>
              <DetailsList
                items={changeRows}
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.justified}
                getKey={(r: { key: string }) => r.key}
                columns={[
                  { key: 'o', name: 'On', minWidth: 140, isResizable: true, onRender: (r: { object: string }) => r.object },
                  { key: 'w', name: 'Person or group', minWidth: 180, isResizable: true, onRender: (r: { who: string }) => <strong>{r.who}</strong> },
                  { key: 'b', name: 'Now', minWidth: 140, onRender: (r: { before: string }) => r.before },
                  { key: 'a', name: 'After lock', minWidth: 120, onRender: (r: { after: string }) => <Pill kind="warning">{r.after}</Pill> }
                ]}
              />
            </div>
          )}
          {preview.objects.some((o) => o.inherited) && (
            <p className={styles.muted}>
              {(() => {
                const inh = preview.objects.filter((o) => o.inherited);
                const one = inh.length === 1;
                return `${inh.map((o) => o.object.title).join(', ')} currently inherit${one ? 's' : ''} permissions. Locking gives ${one ? 'it its own copy' : 'them their own copies'} first, and Unlock makes ${one ? 'it' : 'them'} inherit again.`;
              })()}
            </p>
          )}
        </div>
      )}
    </Card>
  );
};
