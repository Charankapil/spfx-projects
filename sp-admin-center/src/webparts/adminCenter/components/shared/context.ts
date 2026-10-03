import * as React from 'react';
import { SPClient } from '../../core/SPClient';
import { IActionLogEntry, ITarget, ViewKey } from '../../models';
import { AdminApi } from '../../services/AdminApi';
import { SearchApi } from '../../services/SearchApi';

export interface IConfirmOptions {
  title: string;
  message: string;
  confirmText: string;
  danger?: boolean;
}

export interface IAdminContext {
  client: SPClient;
  api: AdminApi;
  search: SearchApi;
  /** Web the page lives on (used for tenant-wide search). */
  homeWebUrl: string;
  target: ITarget;
  setTarget: (t: ITarget) => void;
  openView: (v: ViewKey) => void;
  confirm: (o: IConfirmOptions) => Promise<boolean>;
  log: (action: string, target: string, ok: boolean, message?: string) => void;
  actionLog: IActionLogEntry[];
  /** Bumped by "Refresh"; views reload when it changes. */
  reloadToken: number;
  reloadAll: () => void;
  notify: (message: string, kind?: 'success' | 'error' | 'info') => void;
  /** Findings count for the Health badge in the nav. */
  setHealthBadge: (n: number) => void;
}

export const AdminContext = React.createContext<IAdminContext | undefined>(undefined);

export function useAdmin(): IAdminContext {
  const ctx = React.useContext(AdminContext);
  if (!ctx) {
    throw new Error('AdminContext missing');
  }
  return ctx;
}
