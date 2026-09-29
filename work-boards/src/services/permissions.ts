/** SharePoint base permission bits (SP.PermissionKind). */
export enum PermissionKind {
  ViewListItems = 1,
  AddListItems = 2,
  EditListItems = 3,
  DeleteListItems = 4,
  ManageLists = 12,
  ManagePermissions = 25,
  ManageWeb = 30
}

export interface IBasePermissions {
  High: string | number;
  Low: string | number;
}

export function hasPermission(perms: IBasePermissions | null | undefined, kind: PermissionKind): boolean {
  if (!perms) {
    return false;
  }
  const bit = kind - 1;
  const low = Number(perms.Low) >>> 0;
  const high = Number(perms.High) >>> 0;
  if (bit < 32) {
    return ((low >>> bit) & 1) === 1;
  }
  return ((high >>> (bit - 32)) & 1) === 1;
}

/** What the current user can do on a board, derived from list permissions. */
export interface IBoardRights {
  canView: boolean;
  canAdd: boolean;
  canEdit: boolean;
  canDelete: boolean;
  /** Change columns, groups and board settings. */
  canManage: boolean;
  canManagePermissions: boolean;
}

export function boardRights(perms: IBasePermissions | null, isBoardOwner: boolean): IBoardRights {
  const canManageList = hasPermission(perms, PermissionKind.ManageLists);
  return {
    canView: hasPermission(perms, PermissionKind.ViewListItems),
    canAdd: hasPermission(perms, PermissionKind.AddListItems),
    canEdit: hasPermission(perms, PermissionKind.EditListItems),
    canDelete: hasPermission(perms, PermissionKind.DeleteListItems),
    canManage: canManageList || (isBoardOwner && hasPermission(perms, PermissionKind.EditListItems)),
    canManagePermissions: hasPermission(perms, PermissionKind.ManagePermissions)
  };
}

export const READ_ONLY: IBoardRights = {
  canView: true,
  canAdd: false,
  canEdit: false,
  canDelete: false,
  canManage: false,
  canManagePermissions: false
};
