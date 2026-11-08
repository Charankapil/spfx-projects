import { approvalAssignee, buildKql, formatDate, isPending, safeUrl, sameEmail, statusClass } from '../utils/helpers';

describe('sameEmail', () => {
  it('compares case-insensitively and trims', () => expect(sameEmail(' Ann@Contoso.com', 'ann@contoso.com ')).toBe(true));
  it('never matches empty values', () => {
    expect(sameEmail('', '')).toBe(false);
    expect(sameEmail(undefined, undefined)).toBe(false);
  });
  it('rejects different addresses', () => expect(sameEmail('a@x.com', 'b@x.com')).toBe(false));
});

describe('approvalAssignee / isPending', () => {
  it('prefers ApproverEmail over Owner', () => expect(approvalAssignee({ Id: 1, Title: 't', ApproverEmail: 'a@x.com', Owner: 'o@x.com' })).toBe('a@x.com'));
  it('falls back to Owner', () => expect(approvalAssignee({ Id: 1, Title: 't', Owner: 'o@x.com' })).toBe('o@x.com'));
  it('detects pending status case-insensitively', () => {
    expect(isPending({ Id: 1, Title: 't', Status: 'pending' })).toBe(true);
    expect(isPending({ Id: 1, Title: 't', Status: 'Approved' })).toBe(false);
  });
});

describe('statusClass', () => {
  it('makes multi-word statuses a single class', () => expect(statusClass('In Progress')).toBe('in-progress'));
  it('handles undefined', () => expect(statusClass(undefined)).toBe(''));
});

describe('safeUrl', () => {
  it('allows http, https and mailto', () => {
    expect(safeUrl('https://contoso.sharepoint.com/a')).toBe('https://contoso.sharepoint.com/a');
    expect(safeUrl('mailto:a@x.com')).toBe('mailto:a@x.com');
  });
  it('blocks script and relative urls', () => {
    // eslint-disable-next-line no-script-url
    expect(safeUrl('javascript:alert(1)')).toBeUndefined();
    expect(safeUrl(' JAVASCRIPT:alert(1)')).toBeUndefined();
    expect(safeUrl('')).toBeUndefined();
  });
});

describe('buildKql', () => {
  it('returns only the restriction for blank input', () => expect(buildKql('  ', 'IsDocument:1')).toBe('IsDocument:1'));
  it('wraps user text so OR cannot escape the restriction', () => expect(buildKql('a OR b', 'IsDocument:1')).toBe('(a OR b) AND IsDocument:1'));
});

describe('formatDate', () => {
  it('returns empty for invalid or missing values', () => {
    expect(formatDate(undefined)).toBe('');
    expect(formatDate('not a date')).toBe('');
  });
  it('formats valid ISO dates', () => expect(formatDate('2026-01-15T10:00:00Z')).not.toBe(''));
});
