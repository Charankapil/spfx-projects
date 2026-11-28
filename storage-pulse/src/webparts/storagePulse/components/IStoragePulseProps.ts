import { WebPartContext } from '@microsoft/sp-webpart-base';

import { ScanScope } from '../models/IScanResult';
import { ScanPermission } from '../services/access';
import { ISpeedProfile, ScanSpeed } from '../services/RequestGovernor';

export type { ScanPermission };
export type ScanMode = 'quick' | 'detailed';

export interface IStoragePulseTheme {
  /** True on dark section backgrounds, dark SharePoint themes and dark / high-contrast Teams. */
  isDark: boolean;
  accent?: string;
  text?: string;
}

/** 'site': scans the site it is on. 'central': scans any site collection on the tenant, picked from an overview. */
export type PulseMode = 'site' | 'central';

export interface IStoragePulseProps {
  title: string;
  mode?: PulseMode;
  scope: ScanScope;
  thresholdMonths: number;
  scanPermission: ScanPermission;
  /** Emails or sign-in names allowed to scan when scanPermission is 'people'. */
  scanAllowedPeople: string[];
  scanMode: ScanMode;
  scanSpeed: ScanSpeed;
  includeHidden: boolean;
  excludeSystemLibraries: boolean;
  excludedLibraries: string[];
  staleAfterDays: number;
  theme: IStoragePulseTheme;
  /** Test hooks (not exposed in the property pane): shorten the scan's waiting and checkpoint intervals. */
  tuning?: {
    throttlePatienceMs?: number;
    checkpointEveryMs?: number;
    autoRetryDelayMs?: number;
    profileOverride?: Partial<ISpeedProfile>;
  };
  context: WebPartContext;
  /** Central mode: the site collection this view is about, when it is not the page's own. */
  targetSite?: { url: string; title: string };
  /** Central mode: back to the overview. */
  onBack?: () => void;
  /** Called after results have been saved, so the overview can refresh. */
  onSaved?: () => void;
}
