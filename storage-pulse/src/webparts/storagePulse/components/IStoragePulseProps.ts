import { WebPartContext } from '@microsoft/sp-webpart-base';

import { ScanScope } from '../models/IScanResult';

export type ScanPermission = 'owners' | 'everyone';

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
  includeHidden: boolean;
  excludeSystemLibraries: boolean;
  excludedLibraries: string[];
  staleAfterDays: number;
  theme: IStoragePulseTheme;
  context: WebPartContext;
}
