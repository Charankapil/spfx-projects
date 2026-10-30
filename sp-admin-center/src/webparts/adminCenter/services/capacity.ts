import { ITenantOverview } from './TenantGrowth';

const TB = 1099511627776;

export function formatTB(bytes: number): string {
  const tb = bytes / TB;
  return `${tb >= 100 ? tb.toFixed(0) : tb >= 10 ? tb.toFixed(1) : tb.toFixed(2)} TB`;
}

export interface ICapacity {
  usedBytes: number;
  capacityBytes?: number;
  leftBytes?: number;
  fraction?: number;
}

export function capacityOf(o: ITenantOverview, capacityTB: number | undefined, countDeleted: boolean): ICapacity {
  // Microsoft 365 Archive storage is billed separately from SharePoint storage, so archived
  // sites are reported on their own and not counted as used tenant storage.
  const used = o.activeBytes + (countDeleted ? o.deletedBytes : 0);
  if (!capacityTB) {
    return { usedBytes: used };
  }
  const cap = capacityTB * TB;
  return { usedBytes: used, capacityBytes: cap, leftBytes: cap - used, fraction: used / cap };
}

