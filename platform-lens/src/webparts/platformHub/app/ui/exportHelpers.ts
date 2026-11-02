import { PageId } from './AppContext';
import { IPageDef, PAGES } from './sections';

export function pageDef(id: PageId): IPageDef | undefined {
  return PAGES.filter((p) => p.id === id)[0];
}

const PRICE_KEY = 'platformLens.licencePrices.v1';

/** Monthly unit prices the admin typed in (stored only in this browser). */
export function loadPrices(): Record<string, number> {
  try {
    const raw = window.localStorage.getItem(PRICE_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    const out: Record<string, number> = {};
    Object.keys(parsed || {}).forEach((k) => {
      const n = Number(parsed[k]);
      if (isFinite(n) && n >= 0) out[k] = n;
    });
    return out;
  } catch {
    return {};
  }
}

export function savePrices(prices: Record<string, number>): void {
  try {
    window.localStorage.setItem(PRICE_KEY, JSON.stringify(prices));
  } catch {
    /* storage can be blocked; prices then last for this session only */
  }
}
