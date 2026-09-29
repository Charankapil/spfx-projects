const ALPHABET = 'abcdefghijklmnopqrstuvwxyz0123456789';

/** Short random id for columns, groups and labels. Starts with a letter so it is a valid field name part. */
export function shortId(length: number = 6): string {
  let out = ALPHABET.charAt(Math.floor(Math.random() * 26));
  for (let i = 1; i < length; i++) {
    out += ALPHABET.charAt(Math.floor(Math.random() * ALPHABET.length));
  }
  return out;
}

/** Board key from a title: "Q4 Marketing Campaign" -> "QMC". Upper-case letters and digits, 2-8 chars. */
export function suggestBoardKey(title: string, taken: string[]): string {
  const words = title
    .toUpperCase()
    .replace(/[^A-Z0-9 ]/g, ' ')
    .split(' ')
    .filter(w => w.length > 0);
  let base = words.length >= 2 ? words.map(w => w.charAt(0)).join('') : (words[0] || 'BRD').substring(0, 4);
  if (base.length < 2) {
    base = (base + 'BRD').substring(0, 3);
  }
  base = base.substring(0, 6);
  const upperTaken = taken.map(t => t.toUpperCase());
  if (upperTaken.indexOf(base) < 0) {
    return base;
  }
  for (let n = 2; n < 1000; n++) {
    const candidate = base + n;
    if (upperTaken.indexOf(candidate) < 0) {
      return candidate;
    }
  }
  return base + shortId(3).toUpperCase();
}

export function isValidBoardKey(key: string): boolean {
  return /^[A-Z][A-Z0-9]{1,7}$/.test(key);
}
