/** 1,284 / 12.9K / 4.2M - full digits below 10,000, compact above. */
export function formatCompact(value: number): string {
  const abs = Math.abs(value);
  if (abs < 10000) {
    return value.toLocaleString();
  }
  if (abs < 1000000) {
    return `${(value / 1000).toFixed(abs < 100000 ? 1 : 0)}K`;
  }
  if (abs < 1000000000) {
    return `${(value / 1000000).toFixed(abs < 100000000 ? 1 : 0)}M`;
  }
  return `${(value / 1000000000).toFixed(1)}B`;
}

export function formatDate(date: Date): string {
  return date.toLocaleString(undefined, {
    day: 'numeric',
    month: 'short',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });
}

export function formatShortDate(date: Date): string {
  return date.toLocaleDateString(undefined, { day: 'numeric', month: 'short' });
}

export function formatDuration(ms: number): string {
  const seconds = Math.max(0, Math.round(ms / 1000));
  if (seconds < 60) {
    return `${seconds}s`;
  }
  const minutes = Math.round(seconds / 60);
  if (minutes < 60) {
    return `${minutes} min`;
  }
  return `${Math.floor(minutes / 60)} h ${minutes % 60} min`;
}

export function formatPercent(part: number, whole: number): string {
  if (!whole) {
    return '0%';
  }
  const pct = (part / whole) * 100;
  if (pct > 0 && pct < 1) {
    return '<1%';
  }
  return `${Math.round(pct)}%`;
}
