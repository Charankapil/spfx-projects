import * as React from 'react';
import { Panel, PanelType, TextField, PrimaryButton, DefaultButton, Pivot, PivotItem, MessageBar, MessageBarType, Spinner, SpinnerSize, Dropdown, Icon, ComboBox } from '@fluentui/react';
import styles from '../WorkBoards.module.scss';
import { IBoard, IPerson } from '../../models/types';
import { IBoardRights } from '../../services/permissions';
import { IBoardMember, RoleName } from '../../services/BoardService';
import { useApp } from '../AppContext';
import { PeoplePicker, Swatches, useConfirm, RolesEditor } from '../common/Common';

export interface IBoardSettingsProps {
  board: IBoard;
  rights: IBoardRights;
  onChanged: (board: IBoard) => void;
  onDeleted: () => void;
  onClose: () => void;
}

export function BoardSettings(props: IBoardSettingsProps): JSX.Element {
  const app = useApp();
  const { services } = app;
  const { board, rights } = props;
  const [title, setTitle] = React.useState(board.title);
  const [description, setDescription] = React.useState(board.description);
  const [folder, setFolder] = React.useState(board.folder);
  const [color, setColor] = React.useState(board.color);
  const [owners, setOwners] = React.useState<IPerson[]>([]);
  const [roles, setRoles] = React.useState<{ head: IPerson[]; lead: IPerson[]; sponsor: IPerson[] }>({ head: [], lead: [], sponsor: [] });
  const roleIds = board.roles.head.concat(board.roles.lead, board.roles.sponsor);
  React.useEffect(() => {
    services.people.resolve(roleIds).then(m => setRoles({
      head: board.roles.head.map(id => m.get(id) as IPerson),
      lead: board.roles.lead.map(id => m.get(id) as IPerson),
      sponsor: board.roles.sponsor.map(id => m.get(id) as IPerson)
    })).catch(() => undefined);
  }, [board.id, roleIds.join(',')]);
  const [busy, setBusy] = React.useState('');
  const [error, setError] = React.useState('');
  const [members, setMembers] = React.useState<IBoardMember[] | null>(null);
  const [newPeople, setNewPeople] = React.useState<IPerson[]>([]);
  const [newRole, setNewRole] = React.useState<RoleName>('Member');
  const [confirmDialog, confirm] = useConfirm();
  const folders = app.boards.map(b => b.folder).filter((f, i, a) => f && a.indexOf(f) === i);

  // Board owner ids -> people (names) for the picker.
  React.useEffect(() => {
    Promise.all(board.ownerIds.map(id => services.sp.get<{ Id: number; Title: string; Email: string }>(`web/getuserbyid(${id})?$select=Id,Title,Email`).catch(() => null)))
      .then(us => setOwners(us.filter(u => u !== null).map(u => ({ id: (u as { Id: number }).Id, title: (u as { Title: string }).Title, email: (u as { Email: string }).Email }))))
      .catch(() => undefined);
  }, [board.id]);

  const loadMembers = (): void => {
    if (board.privacy !== 'Private') {
      return;
    }
    setMembers(null);
    services.boards.listMembers(board).then(setMembers).catch(e => { setError((e as Error).message); setMembers([]); });
  };
  React.useEffect(loadMembers, [board.id, board.privacy]);

  const run = async (label: string, fn: () => Promise<void>): Promise<void> => {
    setBusy(label);
    setError('');
    try {
      await fn();
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy('');
    }
  };

  const saveGeneral = (): Promise<void> => run('Saving…', async () => {
    const updated = await services.boards.updateProps(board, {
      title: title.trim() || board.title, description, folder: folder.trim(), color, ownerIds: owners.map(o => o.id),
      roles: { head: roles.head.map(p => p.id), lead: roles.lead.map(p => p.id), sponsor: roles.sponsor.map(p => p.id) }
    });
    if (rights.canManagePermissions) {
      // On a private board, people in project roles get read access so they can follow it.
      await services.boards.grantRoleViewers(updated);
    }
    props.onChanged(updated);
  });

  const togglePrivacy = async (): Promise<void> => {
    if (board.privacy === 'Main') {
      const ok = await confirm({
        title: 'Make this board private?',
        message: 'Only the people you add, the board owners and the site owners will be able to see it. Everyone else loses access straight away.',
        confirmText: 'Make private'
      });
      if (ok) {
        await run('Making the board private…', async () => props.onChanged(await services.boards.makePrivate(board, app.me.id, [])));
      }
    } else {
      const ok = await confirm({
        title: 'Make this board visible to the whole site?',
        message: 'Everyone with access to this site will get the same access to the board.',
        confirmText: 'Make it a main board'
      });
      if (ok) {
        await run('Updating permissions…', async () => props.onChanged(await services.boards.makeMain(board)));
      }
    }
  };

  const addMembers = (): Promise<void> => run('Adding people…', async () => {
    for (const p of newPeople) {
      await services.boards.grant(board, p.id, newRole);
    }
    setNewPeople([]);
    loadMembers();
  });

  const removeMember = async (m: IBoardMember): Promise<void> => {
    const ok = await confirm({ title: `Remove ${m.title}?`, message: 'They will no longer see this board.', confirmText: 'Remove', danger: true });
    if (ok) {
      await run('Removing…', async () => {
        await services.boards.revoke(board, m.principalId);
        loadMembers();
      });
    }
  };

  const archive = (): Promise<void> => run('Saving…', async () => props.onChanged(await services.boards.updateProps(board, { archived: !board.archived })));

  const remove = async (): Promise<void> => {
    const ok = await confirm({
      title: `Delete "${board.title}"?`,
      message: 'The board, its items, updates and files go to the site recycle bin. A site owner can restore them from there for 93 days.',
      confirmText: 'Delete board',
      danger: true
    });
    if (ok) {
      await run('Deleting…', async () => {
        await services.boards.deleteBoard(board);
        props.onDeleted();
      });
    }
  };

  return (
    <Panel isOpen type={PanelType.medium} onDismiss={props.onClose} headerText="Board settings" closeButtonAriaLabel="Close" isLightDismiss>
      {error && <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError('')}>{error}</MessageBar>}
      {busy && <Spinner size={SpinnerSize.small} label={busy} labelPosition="right" />}
      <Pivot>
        <PivotItem headerText="General">
          <div style={{ display: 'flex', flexDirection: 'column', gap: 12, paddingTop: 12 }}>
            <TextField label="Board name" value={title} onChange={(_, v) => setTitle(v || '')} disabled={!rights.canManage} required />
            <TextField label="Description" multiline rows={3} value={description} onChange={(_, v) => setDescription(v || '')} disabled={!rights.canManage} />
            <ComboBox label="Folder" allowFreeform autoComplete="on" text={folder} disabled={!rights.canManage}
              options={folders.map(f => ({ key: f, text: f }))}
              onChange={(_, option, __, value) => setFolder(option ? String(option.text) : value || '')}
              placeholder="No folder" />
            <div>
              <label className={styles.small}>Colour</label>
              <Swatches value={color} onChange={c => { if (rights.canManage) { setColor(c); } }} />
            </div>
            <PeoplePicker label="Board owners (can change columns, groups and settings)" selected={owners} onChange={setOwners} />
            <h3 className={styles.sectionTitle}>Project roles</h3>
            <p className={`${styles.small} ${styles.muted}`} style={{ margin: 0 }}>
              People in these roles see this board under My projects.{board.privacy === 'Private' ? ' On this private board they also get read access.' : ''}
            </p>
            <RolesEditor value={roles} onChange={setRoles} disabled={!rights.canManage} />
            <div className={styles.row}>
              <span className={`${styles.small} ${styles.muted}`}>Board key: <strong>{board.key}</strong> (item IDs look like {board.key}-12)</span>
            </div>
            {rights.canManage && <div><PrimaryButton text="Save" onClick={() => { saveGeneral().catch(() => undefined); }} disabled={!!busy} /></div>}
          </div>
        </PivotItem>
        <PivotItem headerText="Access">
          <div style={{ display: 'flex', flexDirection: 'column', gap: 12, paddingTop: 12 }}>
            <div className={styles.row}>
              <Icon iconName={board.privacy === 'Private' ? 'Lock' : 'People'} />
              <strong>{board.privacy === 'Private' ? 'Private board' : 'Main board'}</strong>
            </div>
            <p className={styles.muted} style={{ margin: 0 }}>
              {board.privacy === 'Private'
                ? 'Only the people listed below, and the site owners, can see this board.'
                : 'Everyone with access to this site can see this board. Site members can edit it, and visitors can only view it.'}
            </p>
            {rights.canManagePermissions ? (
              <div><DefaultButton text={board.privacy === 'Private' ? 'Make it a main board' : 'Make private'} onClick={() => { togglePrivacy().catch(() => undefined); }} disabled={!!busy} /></div>
            ) : (
              <p className={`${styles.small} ${styles.muted}`}>Only a site owner can change who can see this board.</p>
            )}
            {board.privacy === 'Private' && (
              <>
                <h3 className={styles.sectionTitle}>People with access</h3>
                {members === null && <Spinner size={SpinnerSize.small} />}
                {(members || []).map(m => (
                  <div key={m.principalId} className={styles.memberRow}>
                    <Icon iconName={m.isGroup ? 'Group' : 'Contact'} />
                    <span style={{ flex: 1 }}>{m.title}</span>
                    <span className={`${styles.small} ${styles.muted}`}>{m.roles.join(', ')}</span>
                    {rights.canManagePermissions && !m.isGroup && (
                      <button type="button" className={styles.iconBtn} onClick={() => { removeMember(m).catch(() => undefined); }} aria-label={`Remove ${m.title}`}>
                        <Icon iconName="Cancel" />
                      </button>
                    )}
                  </div>
                ))}
                {rights.canManagePermissions && (
                  <>
                    <PeoplePicker label="Add people" selected={newPeople} onChange={setNewPeople} />
                    <div className={styles.row}>
                      <Dropdown selectedKey={newRole} styles={{ root: { width: 200 } }} ariaLabel="Access level"
                        options={[
                          { key: 'Member', text: 'Member (can edit items)' },
                          { key: 'Viewer', text: 'Viewer (read only)' },
                          { key: 'Owner', text: 'Owner (can change the board)' }
                        ]}
                        onChange={(_, o) => o && setNewRole(o.key as RoleName)} />
                      <PrimaryButton text="Add" disabled={newPeople.length === 0 || !!busy} onClick={() => { addMembers().catch(() => undefined); }} />
                    </div>
                  </>
                )}
              </>
            )}
          </div>
        </PivotItem>
        {rights.canManage && (
          <PivotItem headerText="Archive or delete">
            <div style={{ display: 'flex', flexDirection: 'column', gap: 16, paddingTop: 12 }}>
              <div>
                <h3 className={styles.sectionTitle}>{board.archived ? 'Restore board' : 'Archive board'}</h3>
                <p className={styles.muted}>{board.archived ? 'Show this board in the sidebar and My Work again.' : 'Hide the board from the sidebar and My Work. Nothing is deleted, and you can restore it here.'}</p>
                <DefaultButton text={board.archived ? 'Restore' : 'Archive'} onClick={() => { archive().catch(() => undefined); }} disabled={!!busy} />
              </div>
              <div>
                <h3 className={styles.sectionTitle}>Delete board</h3>
                <p className={styles.muted}>Moves the board and everything on it to the site recycle bin.</p>
                <DefaultButton text="Delete board" onClick={() => { remove().catch(() => undefined); }} disabled={!!busy}
                  styles={{ root: { color: '#d83a52', borderColor: '#d83a52' } }} />
              </div>
            </div>
          </PivotItem>
        )}
      </Pivot>
      {confirmDialog}
    </Panel>
  );
}
