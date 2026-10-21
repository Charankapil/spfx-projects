/**
 * Reads a typed date in the browser's own day / month / year order, so
 * "04/11/2026" is 4 November in en-GB and April 11 in en-US. (The date
 * picker's default parser always assumes the US order.)
 */
export function parseLocalDate(text: string, locale?: string): Date | null {
  const parts = text.trim().split(/[^0-9]+/).filter((p) => p !== '').map(Number);
  if (parts.length < 2 || parts.length > 3) {
    return null;
  }
  const sample = new Date(2001, 10, 22).toLocaleDateString(locale);
  const order = ['22', '11', '2001']
    .map((token, i) => ({ part: ['d', 'm', 'y'][i], at: sample.indexOf(token) }))
    .sort((a, b) => a.at - b.at)
    .map((x) => x.part);
  // ISO-style input (a four-digit first part) is always year-month-day.
  const fields = parts[0] > 999 ? ['y', 'm', 'd'] : parts.length === 2 ? order.filter((p) => p !== 'y') : order;
  let day = 0;
  let month = 0;
  let year = new Date().getFullYear();
  fields.forEach((f, i) => {
    if (f === 'd') {
      day = parts[i];
    } else if (f === 'm') {
      month = parts[i];
    } else {
      year = parts[i] < 100 ? 2000 + parts[i] : parts[i];
    }
  });
  const date = new Date(year, month - 1, day);
  return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day ? date : null;
}
