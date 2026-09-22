export type ScanPhase =
  | 'idle'
  | 'starting'
  | 'discovering-structure'
  | 'aggregating-file-types'
  | 'completed'
  | 'cancelled'
  | 'error';

export interface IScanProgress {
  phase: ScanPhase;
  currentItem: string;
  websDiscovered: number;
  librariesDiscovered: number;
  librariesScanned: number;
  message?: string;
}
