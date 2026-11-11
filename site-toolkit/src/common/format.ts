export function formatBytes(bytes: number): string {
  if (!bytes || bytes <= 0 || isNaN(bytes)) {
    return '0 B';
  }
  const units = ['B', 'KB', 'MB', 'GB', 'TB', 'PB'];
  const exponent = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / Math.pow(1024, exponent);
  return `${value.toFixed(value >= 100 || exponent === 0 ? 0 : 1)} ${units[exponent]}`;
}

export function formatNumber(value: number): string {
  return isNaN(value) ? '0' : value.toLocaleString();
}

export function parseDate(value: string | undefined | null): Date | undefined {
  if (!value) {
    return undefined;
  }
  const date = new Date(value);
  return isNaN(date.getTime()) ? undefined : date;
}

export function formatDate(value: string | Date | undefined | null): string {
  const date = value instanceof Date ? value : parseDate(value);
  return date ? date.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' }) : '';
}

export function formatDateTime(value: string | Date | undefined | null): string {
  const date = value instanceof Date ? value : parseDate(value);
  return date
    ? date.toLocaleString(undefined, { year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' })
    : '';
}

/** "3 months ago", "2 years ago" - coarse on purpose. */
export function formatAge(value: string | Date | undefined | null, now: Date = new Date()): string {
  const date = value instanceof Date ? value : parseDate(value);
  if (!date) {
    return '';
  }
  const days = Math.floor((now.getTime() - date.getTime()) / 86400000);
  if (days < 1) {
    return 'today';
  }
  if (days < 31) {
    return `${days} day${days === 1 ? '' : 's'} ago`;
  }
  if (days < 365) {
    const months = Math.floor(days / 30.44);
    return `${months} month${months === 1 ? '' : 's'} ago`;
  }
  const years = Math.floor(days / 365.25);
  const rest = Math.floor((days - years * 365.25) / 30.44);
  return rest > 0 ? `${years} yr ${rest} mo ago` : `${years} year${years === 1 ? '' : 's'} ago`;
}

export function extensionOf(fileName: string): string {
  const dot = fileName.lastIndexOf('.');
  return dot > 0 ? fileName.substring(dot + 1).toLowerCase() : '';
}

/** Fluent icon for a file extension; no network request involved. */
export function fileIcon(extension: string): string {
  switch (extension.toLowerCase()) {
    case 'doc':
    case 'docx':
    case 'dotx':
      return 'WordDocument';
    case 'xls':
    case 'xlsx':
    case 'xltx':
    case 'csv':
      return 'ExcelDocument';
    case 'ppt':
    case 'pptx':
    case 'potx':
      return 'PowerPointDocument';
    case 'pdf':
      return 'PDF';
    case 'one':
      return 'OneNoteLogo';
    case 'png':
    case 'jpg':
    case 'jpeg':
    case 'gif':
    case 'svg':
      return 'FileImage';
    case 'mp4':
    case 'mov':
    case 'avi':
      return 'Video';
    case 'zip':
    case '7z':
    case 'rar':
      return 'ZipFolder';
    case 'aspx':
      return 'FileASPX';
    case 'txt':
    case 'md':
      return 'TextDocument';
    default:
      return 'Page';
  }
}

/**
 * Links typed into a list by site owners are only rendered when they are
 * ordinary web or site-relative addresses, never script URLs.
 */
export function safeHref(value: string | undefined | null): string | undefined {
  const url = (value || '').trim();
  if (/^https?:\/\//i.test(url) || (url.charAt(0) === '/' && url.charAt(1) !== '/')) {
    return url;
  }
  return undefined;
}

/** Date as yyyy-mm-dd in the browser's time zone, for date inputs. */
export function toDateInput(value: string | Date | undefined | null): string {
  const date = value instanceof Date ? value : parseDate(value);
  if (!date) {
    return '';
  }
  const pad = (n: number): string => (n < 10 ? `0${n}` : String(n));
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

/** Local midnight of a yyyy-mm-dd value as an ISO string; undefined when the text is not a date. */
export function fromDateInput(value: string): string | undefined {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) {
    return undefined;
  }
  const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  // JavaScript rolls "2024-13-45" over into a real date; only accept what was actually typed.
  if (isNaN(date.getTime()) || date.getMonth() !== Number(match[2]) - 1 || date.getDate() !== Number(match[3])) {
    return undefined;
  }
  return date.toISOString();
}
