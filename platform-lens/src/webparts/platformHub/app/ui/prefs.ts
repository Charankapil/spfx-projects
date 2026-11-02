/** Per-browser preferences (localStorage; every access tolerates blocked storage). */

const PBI_DAYS_KEY = 'platformLens.pbiDays.v1';
export const PBI_DAY_OPTIONS = [1, 7, 14, 30];

/** Activity period (days) used by the Power BI pages and their exports. */
export function loadPbiDays(): number {
  try {
    const n = Number(window.localStorage.getItem(PBI_DAYS_KEY));
    return PBI_DAY_OPTIONS.indexOf(n) >= 0 ? n : 7;
  } catch {
    return 7;
  }
}

export function savePbiDays(days: number): void {
  try {
    window.localStorage.setItem(PBI_DAYS_KEY, String(days));
  } catch {
    /* storage can be blocked */
  }
}
