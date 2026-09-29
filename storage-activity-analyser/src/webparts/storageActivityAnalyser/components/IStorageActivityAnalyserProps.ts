import { WebPartContext } from '@microsoft/sp-webpart-base';

import { ScanScope } from '../models/IScanResult';

export interface IStorageActivityAnalyserProps {
  title: string;
  scope: ScanScope;
  thresholdMonths: number;
  context: WebPartContext;
}
