export type ScanPhase =
  | 'idle'
  | 'starting'
  | 'discovering-structure'
  | 'aggregating-file-types'
  | 'estimating-storage'
  | 'completed'
  | 'cancelled'
  | 'error';

export interface IScanProgress {
  phase: ScanPhase;
  currentItem: string;
  websDiscovered: number;
  librariesDiscovered: number;
  librariesScanned: number;
  typesToEstimate?: number;
  typesEstimated?: number;
  message?: string;
}
