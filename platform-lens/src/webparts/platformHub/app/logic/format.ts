const DAY = 86400000;

export function fmtNumber(n: number | undefined): string {
  if (n === undefined || n === null || isNaN(n)) return '—';
  return Math.round(n).toLocaleString('en-US');
}

export function fmtCompact(n: number): string {
  const abs = Math.abs(n);
  if (abs >= 1e9) return (n / 1e9).toFixed(abs >= 1e10 ? 0 : 1) + 'B';
  if (abs >= 1e6) return (n / 1e6).toFixed(abs >= 1e7 ? 0 : 1) + 'M';
  if (abs >= 1e4) return (n / 1e3).toFixed(0) + 'k';
  if (abs >= 1e3) return (n / 1e3).toFixed(1) + 'k';
  return String(Math.round(n));
}

export function fmtBytes(bytes: number): string {
  if (!bytes) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB', 'TB', 'PB'];
  let i = 0;
  let v = bytes;
  while (v >= 1024 && i < units.length - 1) {
    v /= 1024;
    i++;
  }
  return `${v >= 100 || i === 0 ? Math.round(v) : v.toFixed(1)} ${units[i]}`;
}

export function fmtMb(mb: number): string {
  return fmtBytes(mb * 1024 * 1024);
}

export function fmtPercent(part: number, whole: number): string {
  if (!whole) return '0%';
  const p = (part / whole) * 100;
  return `${p >= 10 || p === 0 ? Math.round(p) : p.toFixed(1)}%`;
}

export function fmtDate(iso: string | undefined): string {
  if (!iso) return '—';
  const t = Date.parse(iso);
  if (isNaN(t)) return '—';
  const d = new Date(t);
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
}

export function fmtDateTime(iso: string | undefined): string {
  if (!iso) return '—';
  const t = Date.parse(iso);
  if (isNaN(t)) return '—';
  const d = new Date(t);
  const hh = String(d.getHours()).padStart(2, '0');
  const mm = String(d.getMinutes()).padStart(2, '0');
  return `${fmtDate(iso)} ${hh}:${mm}`;
}

/** Whole days since `iso`; undefined when the date is missing. */
export function daysSince(iso: string | undefined, now: number = Date.now()): number | undefined {
  if (!iso) return undefined;
  const t = Date.parse(iso);
  if (isNaN(t)) return undefined;
  return Math.floor((now - t) / DAY);
}

export function fmtAge(iso: string | undefined): string {
  const d = daysSince(iso);
  if (d === undefined) return 'never';
  if (d <= 0) return 'today';
  if (d === 1) return '1 day ago';
  if (d < 60) return `${d} days ago`;
  if (d < 730) return `${Math.round(d / 30)} months ago`;
  return `${Math.round(d / 365)} years ago`;
}

/** Only https links from API data are rendered as clickable links. */
export function safeUrl(url: string | undefined): string | undefined {
  return url && /^https:\/\//i.test(url) ? url : undefined;
}

export function plural(n: number, one: string, many?: string): string {
  return `${fmtNumber(n)} ${n === 1 ? one : many || one + 's'}`;
}
