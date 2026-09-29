/**
 * Dates are kept as "YYYY-MM-DD" strings in the app. SharePoint DateTime fields are
 * written at 12:00 UTC so the calendar date reads the same in every time zone from
 * UTC-11 to UTC+11, both in the app and in the SharePoint list UI.
 */

export function toSpDate(day: string | null): string | null {
  if (!day) {
    return null;
  }
  return day + 'T12:00:00Z';
}

export function fromSpDate(value: string | null | undefined): string | null {
  if (!value) {
    return null;
  }
  return value.substring(0, 10);
}

export function todayIso(now: Date = new Date()): string {
  return dateToIso(now);
}

export function dateToIso(d: Date): string {
  const m = d.getMonth() + 1;
  const day = d.getDate();
  return d.getFullYear() + '-' + (m < 10 ? '0' : '') + m + '-' + (day < 10 ? '0' : '') + day;
}

export function isoToDate(iso: string): Date {
  const parts = iso.split('-');
  return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
}

export function addDays(iso: string, days: number): string {
  const d = isoToDate(iso);
  d.setDate(d.getDate() + days);
  return dateToIso(d);
}

export function diffDays(fromIso: string, toIso: string): number {
  const ms = isoToDate(toIso).getTime() - isoToDate(fromIso).getTime();
  return Math.round(ms / 86400000);
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** "3 Oct", or "3 Oct 2027" when the year differs from `now`. */
export function formatDay(iso: string | null, now: Date = new Date()): string {
  if (!iso) {
    return '';
  }
  const d = isoToDate(iso);
  const base = d.getDate() + ' ' + MONTHS[d.getMonth()];
  return d.getFullYear() === now.getFullYear() ? base : base + ' ' + d.getFullYear();
}

export function formatRange(start: string | null, end: string | null, now: Date = new Date()): string {
  if (start && end) {
    return start === end ? formatDay(start, now) : formatDay(start, now) + ' – ' + formatDay(end, now);
  }
  return formatDay(start || end, now);
}

/** "Today", "Tomorrow", "Yesterday", or a formatted day. */
export function relativeDay(iso: string | null, now: Date = new Date()): string {
  if (!iso) {
    return '';
  }
  const diff = diffDays(todayIso(now), iso);
  if (diff === 0) {
    return 'Today';
  }
  if (diff === 1) {
    return 'Tomorrow';
  }
  if (diff === -1) {
    return 'Yesterday';
  }
  return formatDay(iso, now);
}

/** "5 min ago", "3 h ago", "2 d ago", or a date, for timestamps. */
export function timeAgo(isoDateTime: string, now: Date = new Date()): string {
  const then = new Date(isoDateTime).getTime();
  const secs = Math.max(0, Math.round((now.getTime() - then) / 1000));
  if (secs < 60) {
    return 'just now';
  }
  const mins = Math.round(secs / 60);
  if (mins < 60) {
    return mins + ' min ago';
  }
  const hours = Math.round(mins / 60);
  if (hours < 24) {
    return hours + ' h ago';
  }
  const days = Math.round(hours / 24);
  if (days < 7) {
    return days + ' d ago';
  }
  return formatDay(dateToIso(new Date(isoDateTime)), now);
}

export const MONTH_NAMES = MONTHS;
