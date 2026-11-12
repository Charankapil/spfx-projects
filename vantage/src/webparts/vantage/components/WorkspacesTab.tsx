import * as React from 'react';
import { ActionButton, DetailsList, DetailsListLayoutMode, type IColumn, Link, MessageBarType, PrimaryButton, SelectionMode, TextField } from '@fluentui/react';

import styles from './VantageApp.module.scss';
import { Badge, Card, listUrl, useApp } from './common';
import type { IWorkspaceSummary, WorkspaceRole } from '../models/Models';

const ROLE_LABEL: Record<WorkspaceRole, string> = { owner: 'Site owner', admin: 'Administrator', member: 'Contributor', reader: 'Reader', none: 'No access' };

export interface IWorkspacesTabProps {
  workspaces: IWorkspaceSummary[];
  roles: Record<string, WorkspaceRole>;
  canCreate: boolean;
  onCreated: (ws: IWorkspaceSummary) => void;
  onSelect: (ws: IWorkspaceSummary) => void;
}

export const WorkspacesTab: React.FC<IWorkspacesTabProps> = ({ workspaces, roles, canCreate, onCreated, onSelect }) => {
  const { provisioner, permissions, perform, busy, notify, user, workspace, webServerRelativeUrl, siteOwner } = useApp();
  const [name, setName] = React.useState('');

  const create = (): Promise<boolean> => perform('Creating workspace lists…', async () => {
    const created = await provisioner.createWorkspace(name, workspaces, user.loginName);
    setName('');
    try {
      await permissions.configureWorkspace(created, user.loginName);
      await provisioner.markWorkspaceSecured(created);
      created.secured = true;
      notify(MessageBarType.success, `Workspace “${created.title}” was created and secured. You are its administrator; add people on the Access tab.`);
    } catch (e) {
      notify(MessageBarType.warning, `Workspace “${created.title}” was created, but its permissions could not be configured: ${(e as Error).message} ` +
        'Until a site owner secures it on the Access tab, everyone with access to this site can see its data.');
    }
    onCreated(created);
  });

  const columns: IColumn[] = [
    { key: 'title', name: 'Workspace', minWidth: 140, maxWidth: 240, isRowHeader: true, onRender: (w: IWorkspaceSummary) => <strong>{w.title}</strong> },
    { key: 'role', name: 'Your role', minWidth: 100, maxWidth: 130, onRender: (w: IWorkspaceSummary) => ROLE_LABEL[roles[w.id] || 'none'] },
    { key: 'security', name: 'Access', minWidth: 120, maxWidth: 150,
      onRender: (w: IWorkspaceSummary) => w.secured ? <Badge value="Secured" tone={styles.ok} /> : <Badge value="Site permissions" tone={styles.warn} /> },
    { key: 'lists', name: 'Lists', minWidth: 160,
      onRender: (w: IWorkspaceSummary) => roles[w.id] === 'none' ? '—' : <>
        <Link href={listUrl(webServerRelativeUrl, w.workItemsList)} target="_blank" data-interception="off">Active</Link>{' · '}
        <Link href={listUrl(webServerRelativeUrl, w.archiveList)} target="_blank" data-interception="off">Archive</Link>
      </> },
    { key: 'select', name: '', minWidth: 100,
      onRender: (w: IWorkspaceSummary) => workspace && workspace.id === w.id
        ? <Badge value="Current" tone={styles.info} />
        : <ActionButton iconProps={{ iconName: 'Switch' }} disabled={busy || roles[w.id] === 'none'} onClick={() => onSelect(w)} ariaLabel={`Open ${w.title}`}>Open</ActionButton> }
  ];

  return (
    <div className={styles.stack}>
      {canCreate && (
        <Card title="Create a workspace">
          <p className={styles.muted}>
            A workspace gets its own active list and archive list, plus Administrators, Contributors and Readers groups.
            It is the security boundary: people only see workspaces they belong to.
          </p>
          <form className={styles.inlineForm} onSubmit={e => { e.preventDefault(); if (name.trim() && !busy) create().catch(() => undefined); }}>
            <TextField label="Workspace name" value={name} maxLength={80} placeholder="e.g. IT Operations" onChange={(_, v) => setName(v || '')} />
            <PrimaryButton type="submit" text="Create workspace" iconProps={{ iconName: 'Add' }} disabled={busy || !name.trim()} />
          </form>
          {!siteOwner && <p className={styles.muted}>Securing a new workspace needs permission to manage permissions on this site.</p>}
        </Card>
      )}
      <Card title="Workspaces on this site">
        {workspaces.length
          ? <DetailsList items={workspaces} columns={columns} selectionMode={SelectionMode.none} layoutMode={DetailsListLayoutMode.justified}
              compact getKey={(w: IWorkspaceSummary) => w.id} />
          : <p className={styles.muted}>No workspaces yet.</p>}
      </Card>
    </div>
  );
};
