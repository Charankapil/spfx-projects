import { useEffect, useMemo, useRef } from 'react';
import { SPPermission } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { ListService } from '../lists';
import { SpClient } from '../SpClient';

export interface IToolkit {
  sp: SpClient;
  lists: ListService;
  rootLists: ListService;
  /** Current user's id in this site's user list. */
  userId: number;
  userEmail: string;
  userName: string;
  /** Can create lists and change site settings (site owner). */
  canManage: boolean;
  /** Stable id of the web, for scoping cached values. */
  scope: string;
}

interface ILegacyContext {
  userId?: number;
  isSiteAdmin?: boolean;
}

export function useToolkit(context: WebPartContext): IToolkit {
  return useMemo(() => {
    const sp = SpClient.fromContext(context);
    const page = context.pageContext;
    const legacy = (page.legacyPageContext || {}) as ILegacyContext;
    return {
      sp,
      lists: ListService.forWeb(sp),
      rootLists: ListService.forSiteRoot(sp),
      userId: legacy.userId || 0,
      userEmail: page.user.email || page.user.loginName || '',
      userName: page.user.displayName || '',
      canManage: !!legacy.isSiteAdmin || page.web.permissions.hasPermission(SPPermission.manageWeb),
      scope: page.web.id.toString()
    };
  }, [context]);
}

/** True until the component unmounts; guards state updates after awaits. */
export function useMounted(): { current: boolean } {
  const mounted = useRef(true);
  useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
    };
  }, []);
  return mounted;
}
