import { WebPartContext } from '@microsoft/sp-webpart-base';

import { ScanScope } from '../models/IScanResult';
import { ISpeedProfile, ScanSpeed } from '../services/RequestGovernor';

export type ScanPermission = 'owners' | 'everyone';
export type ScanMode = 'quick' | 'detailed';

export interface IStoragePulseTheme {
  /** True on dark section backgrounds, dark SharePoint themes and dark / high-contrast Teams. */
  isDark: boolean;
  accent?: string;
  text?: string;
}

export interface IStoragePulseProps {
  title: string;
  scope: ScanScope;
  thresholdMonths: number;
  scanPermission: ScanPermission;
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
}
