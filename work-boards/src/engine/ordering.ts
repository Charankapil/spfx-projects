/**
 * Items are ordered by a numeric sort key. Moving an item writes only that item:
 * its new key is the midpoint of its new neighbours. When the gap gets too small
 * the caller renumbers the group (see needsRebalance / rebalance).
 */

export const ORDER_STEP = 1024;
const MIN_GAP = 1e-6;

export function keyBetween(before: number | null, after: number | null): number {
  if (before === null && after === null) {
    return ORDER_STEP;
  }
  if (before === null) {
    return (after as number) - ORDER_STEP;
  }
  if (after === null) {
    return before + ORDER_STEP;
  }
  return (before + after) / 2;
}

export function needsRebalance(before: number | null, after: number | null): boolean {
  return before !== null && after !== null && Math.abs(after - before) < MIN_GAP;
}

/** New keys for a list already in the wanted order. Returns only the entries whose key changes. */
export function rebalance<T extends { id: number; sortOrder: number }>(items: T[]): { id: number; sortOrder: number }[] {
  const changes: { id: number; sortOrder: number }[] = [];
  items.forEach((item, index) => {
    const wanted = (index + 1) * ORDER_STEP;
    if (item.sortOrder !== wanted) {
      changes.push({ id: item.id, sortOrder: wanted });
    }
  });
  return changes;
}

/**
 * Key for moving an item to position `index` in `siblings` (siblings exclude the moved item,
 * and are already sorted).
 */
export function keyForIndex(siblings: { sortOrder: number }[], index: number): { key: number; rebalance: boolean } {
  const before = index > 0 ? siblings[index - 1].sortOrder : null;
  const after = index < siblings.length ? siblings[index].sortOrder : null;
  return { key: keyBetween(before, after), rebalance: needsRebalance(before, after) };
}
