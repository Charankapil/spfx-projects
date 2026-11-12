import * as React from 'react';
import { Checkbox, Dropdown, IconButton, MessageBar, MessageBarType, Persona, PersonaSize, PrimaryButton, Shimmer } from '@fluentui/react';

import styles from './VantageApp.module.scss';
import { Card, type IPersonValue, PeoplePicker, useApp, useConfirm } from './common';
import { GROUP_ROLES, type GroupRole } from '../services/PermissionService';

type Members = Array<{ role: GroupRole; groupId: number; users: Array<{ Id: number; Title: string; LoginName: string; Email?: string }> }>;

export const AccessTab: React.FC<{ onSecured: () => void }> = ({ onSecured }) => {
  const { permissions, provisioner, workspace, perform, busy, notify, user, siteOwner } = useApp();
  const [members, setMembers] = React.useState<Members | undefined>(undefined);
  const [person, setPerson] = React.useState<IPersonValue | undefined>(undefined);
  const [role, setRole] = React.useState<GroupRole>('Members');
  const [siteAccess, setSiteAccess] = React.useState(true);
  const [confirm, confirmDialog] = useConfirm();

  const load = React.useCallback(async (): Promise<void> => {
    if (!workspace) return;
    setMembers(await permissions.getMembers(workspace));
  }, [workspace, permissions]);

  React.useEffect(() => {
    setMembers(undefined);
    load().catch(e => { setMembers([]); notify(MessageBarType.error, `Could not load members: ${(e as Error).message}`); });
  }, [load, notify]);

  if (!workspace) return null;

  const secure = (): Promise<boolean> => perform('Securing workspace…', async () => {
    await permissions.configureWorkspace(workspace, user.loginName);
    await provisioner.markWorkspaceSecured(workspace);
    notify(MessageBarType.success, `“${workspace.title}” now has its own permissions. You were added as an administrator.`);
    onSecured();
    await load();
  });

  const add = (): Promise<boolean> => perform('Adding member…', async () => {
    if (!person) return;
    await permissions.addMember(workspace, person.key, role, siteAccess);
    notify(MessageBarType.success, `${person.name} was added as ${GROUP_ROLES.filter(g => g.role === role)[0].label.toLowerCase().replace(/s$/, '')}.`);
    setPerson(undefined);
    await load();
  });

  return (
    <div className={styles.stack}>
      {!workspace.secured && (
        <MessageBar messageBarType={MessageBarType.severeWarning} isMultiline
          actions={siteOwner ? <PrimaryButton text="Secure workspace" onClick={() => { secure().catch(() => undefined); }} disabled={busy} /> : undefined}>
          This workspace still uses the site’s permissions, so everyone with access to the site can see its tasks.
          {siteOwner ? ' Secure it to give it its own Administrators, Contributors and Readers groups.' : ' Ask a site owner to secure it.'}
        </MessageBar>
      )}

      {workspace.secured && (
        <Card title="Add people">
          <div className={styles.inlineForm}>
            <PeoplePicker label="Person" value={person} onChange={setPerson} placeholder="Search by name or email" />
            <Dropdown label="Role" selectedKey={role} className={styles.rolePicker}
              options={GROUP_ROLES.map(g => ({ key: g.role, text: g.label, title: g.description }))} onChange={(_, o) => setRole(o?.key as GroupRole)} />
            <PrimaryButton text="Add" iconProps={{ iconName: 'AddFriend' }} onClick={() => { add().catch(() => undefined); }} disabled={busy || !person} />
          </div>
          <Checkbox className={styles.checkboxRow} label="Also give them read access to this site (needed to open this page)" checked={siteAccess}
            onChange={(_, v) => setSiteAccess(!!v)} />
          <p className={styles.muted}>Adding people to a group needs permission to manage that group (site owners, or the group owner).</p>
        </Card>
      )}

      <div className={styles.grid3}>
        {GROUP_ROLES.map(g => {
          const group = members && members.filter(m => m.role === g.role)[0];
          return (
            <Card key={g.role} title={g.label}>
              <p className={styles.muted}>{g.description}</p>
              {members === undefined ? <><Shimmer /><Shimmer width="70%" /></> : !group ? <p className={styles.muted}>Group not created yet.</p> : (
                <ul className={styles.memberList}>
                  {group.users.map(u => (
                    <li key={u.Id}>
                      <Persona text={u.Title} secondaryText={u.Email} size={PersonaSize.size32} />
                      <IconButton iconProps={{ iconName: 'Cancel' }} title="Remove" ariaLabel={`Remove ${u.Title}`} disabled={busy}
                        onClick={async () => {
                          const ok = await confirm({ title: 'Remove member?', confirmText: 'Remove', danger: true,
                            text: <>{u.Title} will lose {g.label.toLowerCase()} access to “{workspace.title}”.</> });
                          if (ok) await perform('Removing member…', async () => { await permissions.removeMember(group.groupId, u.Id); await load(); });
                        }} />
                    </li>
                  ))}
                  {!group.users.length && <li className={styles.muted}>Nobody yet.</li>}
                </ul>
              )}
            </Card>
          );
        })}
      </div>
      {confirmDialog}
    </div>
  );
};
