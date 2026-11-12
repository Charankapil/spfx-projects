import { ROLE, SharePointService } from './SharePointService';
import { LISTS } from './ProvisioningService';
import { IWorkspaceSummary, WorkspaceRole } from '../models/Models';

export type GroupRole = 'Admins' | 'Members' | 'Readers';

export const GROUP_ROLES: Array<{ role: GroupRole; label: string; description: string }> = [
  { role: 'Admins', label: 'Administrators', description: 'Full control of workspace data, archive policy and membership' },
  { role: 'Members', label: 'Contributors', description: 'Create and edit boards, tasks and comments' },
  { role: 'Readers', label: 'Readers', description: 'Read-only access to workspace data' }
];

/** Control lists that only site owners and workspace administrators can read and write. */
const ARCHIVE_CONTROL_LISTS = [LISTS.policies, LISTS.jobs, LISTS.log, LISTS.locks];

/**
 * The workspace is the security boundary: its two data lists get unique permissions for the
 * workspace's Admins / Members / Readers SharePoint groups and the site Owners group.
 *
 * Design note: archive control lists are secured once at *list* level (site Owners + every
 * workspace Admins group). Only the single policy row of each workspace gets item-level unique
 * permissions. Job and log rows deliberately inherit the list permissions; giving every audit row
 * its own permission scope (as the 1.0.0 source did) would exceed SharePoint's recommended limit
 * of 5,000 unique scopes per list after a few thousand archived tasks and slow every query.
 * Job and log rows contain only IDs, counts and outcomes - no task content.
 */
export class PermissionService {
  constructor(private readonly sp: SharePointService) {}

  public static groupName(ws: IWorkspaceSummary, role: GroupRole): string {
    return `VNT_${ws.id.replace(/[^A-Za-z0-9]/g, '')}_${role}`;
  }

  /** Resolve the current user's role in every workspace with a single request. */
  public async resolveRoles(
    workspaces: IWorkspaceSummary[], siteOwner: boolean, siteCanEdit: boolean
  ): Promise<Record<string, WorkspaceRole>> {
    const groups = new Set(siteOwner ? [] : await this.sp.getCurrentUserGroupTitles());
    const roles: Record<string, WorkspaceRole> = {};
    for (const ws of workspaces) {
      if (siteOwner) roles[ws.id] = 'owner';
      else if (groups.has(PermissionService.groupName(ws, 'Admins'))) roles[ws.id] = 'admin';
      else if (groups.has(PermissionService.groupName(ws, 'Members'))) roles[ws.id] = 'member';
      else if (groups.has(PermissionService.groupName(ws, 'Readers'))) roles[ws.id] = 'reader';
      // Not yet secured: the lists still inherit the site's permissions.
      else if (!ws.secured) roles[ws.id] = siteCanEdit ? 'member' : 'reader';
      else roles[ws.id] = 'none';
    }
    return roles;
  }

  /** Create the workspace groups and apply unique permissions. Idempotent. Requires Manage Permissions. */
  public async configureWorkspace(ws: IWorkspaceSummary, ownerLogin: string, onStep?: (m: string) => void): Promise<void> {
    const step = (m: string): void => { if (onStep) onStep(m); };
    step('Creating workspace groups…');
    const ids: Record<GroupRole, number> = { Admins: 0, Members: 0, Readers: 0 };
    for (const g of GROUP_ROLES) {
      ids[g.role] = await this.sp.ensureGroup(PermissionService.groupName(ws, g.role), `Vantage · ${ws.title} · ${g.label}`);
    }
    if (ownerLogin) await this.sp.addUserToGroup(ids.Admins, ownerLogin);
    const owners = await this.sp.getAssociatedOwnerGroupId();

    step('Securing workspace lists…');
    const dataGrants = [
      { principalId: owners, roleId: ROLE.fullControl },
      { principalId: ids.Admins, roleId: ROLE.fullControl },
      { principalId: ids.Members, roleId: ROLE.contribute },
      { principalId: ids.Readers, roleId: ROLE.read }
    ];
    await this.sp.setUniquePermissions(this.sp.listPath(ws.workItemsList), dataGrants);
    await this.sp.setUniquePermissions(this.sp.listPath(ws.archiveList), dataGrants);

    step('Securing archive control lists…');
    for (const list of ARCHIVE_CONTROL_LISTS) {
      await this.sp.setUniquePermissions(this.sp.listPath(list), [
        { principalId: owners, roleId: ROLE.fullControl },
        { principalId: ids.Admins, roleId: ROLE.contribute }
      ]);
    }
    const policy = await this.sp.getCollection<{ Id: number }>(
      `${this.sp.listPath(LISTS.policies)}/items?$select=Id&${SharePointService.filter(`WorkspaceId eq '${SharePointService.odataString(ws.id)}'`)}&$top=1`
    );
    if (policy.value.length) {
      await this.sp.setUniquePermissions(`${this.sp.listPath(LISTS.policies)}/items(${policy.value[0].Id})`, [
        { principalId: owners, roleId: ROLE.fullControl },
        { principalId: ids.Admins, roleId: ROLE.contribute }
      ]);
    }
  }

  public async getMembers(ws: IWorkspaceSummary): Promise<Array<{ role: GroupRole; groupId: number; users: Array<{ Id: number; Title: string; LoginName: string; Email?: string }> }>> {
    const out = [];
    for (const g of GROUP_ROLES) {
      try {
        const group = await this.sp.get<{ Id: number }>(`web/sitegroups/getbyname('${encodeURIComponent(PermissionService.groupName(ws, g.role))}')?$select=Id`);
        out.push({ role: g.role, groupId: group.Id, users: await this.sp.getGroupUsers(group.Id) });
      } catch (e) {
        if (!SharePointService.isNotFound(e)) throw e;
      }
    }
    return out;
  }

  /**
   * Adds a user to a workspace group. Workspace groups only grant access to the workspace lists;
   * people also need read access to the site to open the page and the workspace/board registries,
   * so `ensureSiteAccess` also adds them to the site's Visitors group.
   */
  public async addMember(ws: IWorkspaceSummary, loginName: string, role: GroupRole, ensureSiteAccess: boolean): Promise<void> {
    const group = await this.sp.get<{ Id: number }>(`web/sitegroups/getbyname('${encodeURIComponent(PermissionService.groupName(ws, role))}')?$select=Id`);
    await this.sp.addUserToGroup(group.Id, loginName);
    if (ensureSiteAccess) {
      const visitors = await this.sp.get<{ Id: number }>('web/associatedvisitorgroup?$select=Id');
      await this.sp.addUserToGroup(visitors.Id, loginName);
    }
  }

  public removeMember(groupId: number, userId: number): Promise<void> {
    return this.sp.removeUserFromGroup(groupId, userId);
  }
}
