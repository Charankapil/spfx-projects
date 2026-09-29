export type RunPhase = 'idle' | 'scanning' | 'restoring';

export interface IRunProgress {
  phase: RunPhase;
  currentItem: string;
  websChecked: number;
  listsChecked: number;
  itemsChecked: number;
  uniqueFound: number;
  /** The list whose items are being read, and roughly how far through it the scan is. */
  listTitle?: string;
  listItemsTotal?: number;
  listItemsDone?: number;
  restoreTotal?: number;
  restoreDone?: number;
  restored?: number;
  failed?: number;
}

export const IDLE_PROGRESS: IRunProgress = {
  phase: 'idle',
  currentItem: '',
  websChecked: 0,
  listsChecked: 0,
  itemsChecked: 0,
  uniqueFound: 0
};
