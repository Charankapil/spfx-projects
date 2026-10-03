export function formatBytes(bytes: number): string {
  if (!isFinite(bytes) || bytes <= 0) {
    return '0 B';
  }
  const units = ['B', 'KB', 'MB', 'GB', 'TB', 'PB'];
  const i = Math.min(units.length - 1, Math.floor(Math.log(bytes) / Math.log(1024)));
  const v = bytes / Math.pow(1024, i);
  return `${v >= 100 || i === 0 ? v.toFixed(0) : v.toFixed(1)} ${units[i]}`;
}

export function formatCompact(n: number): string {
  const a = Math.abs(n);
  if (a < 10000) {
    return n.toLocaleString();
  }
  if (a < 1e6) {
    return `${(n / 1e3).toFixed(a < 1e5 ? 1 : 0)}K`;
  }
  if (a < 1e9) {
    return `${(n / 1e6).toFixed(a < 1e8 ? 1 : 0)}M`;
  }
  return `${(n / 1e9).toFixed(1)}B`;
}

export function formatDate(d?: Date): string {
  return d ? d.toLocaleDateString(undefined, { day: 'numeric', month: 'short', year: 'numeric' }) : '-';
}

export function formatDateTime(d?: Date): string {
  return d ? d.toLocaleString(undefined, { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '-';
}

export function daysAgo(d?: Date): number | undefined {
  return d ? Math.floor((Date.now() - d.getTime()) / 86400000) : undefined;
}

export function relativeTime(d?: Date): string {
  const n = daysAgo(d);
  if (n === undefined) {
    return '-';
  }
  if (n < 1) {
    return 'today';
  }
  if (n < 31) {
    return `${n} day${n === 1 ? '' : 's'} ago`;
  }
  if (n < 365) {
    return `${Math.floor(n / 30)} mo ago`;
  }
  return `${(n / 365).toFixed(1)} yr ago`;
}
