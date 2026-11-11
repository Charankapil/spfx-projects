/** Marks the principals an owner most needs to notice. */
export function flagOf(login: string): 'external' | 'everyone' | undefined {
  const lower = (login || '').toLowerCase();
  if (lower.indexOf('#ext#') >= 0 || lower.indexOf('urn:spo:guest') >= 0) {
    return 'external';
  }
  if (lower.indexOf('spo-grid-all-users') >= 0 || lower === 'c:0(.s|true' || lower.indexOf('rolemanager|spo-grid') >= 0) {
    return 'everyone';
  }
  return undefined;
}
