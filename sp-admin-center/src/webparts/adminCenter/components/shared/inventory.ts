import { ISiteRow } from '../../models';

/**
 * The tenant site inventory, loaded once and shared by the Sites tabs ("All sites" and
 * "Owners & members"), so switching between them never walks the search index again.
 * Kept for the browser session; the Reload buttons refresh it.
 */
export interface IInventory {
  sites: ISiteRow[];
  truncated: boolean;
  oneDriveHidden: number;
  at: Date;
}

let current: IInventory | undefined;

export function getInventory(): IInventory | undefined {
  return current;
}

export function setInventory(inv: IInventory): void {
  current = inv;
}
