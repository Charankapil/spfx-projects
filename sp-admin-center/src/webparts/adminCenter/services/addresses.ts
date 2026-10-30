/** Splits pasted text (lines, commas, semicolons) into unique e-mail addresses or claims logins. */
export function parseAddresses(input: string): string[] {
  const seen: { [k: string]: boolean } = {};
  return input
    .split(/[\n\r,;]+/)
    .map((s) => s.trim())
    .filter((s) => {
      const key = s.toLowerCase();
      if (!s || seen[key] || (s.indexOf('@') < 0 && s.indexOf('|') < 0)) {
        return false;
      }
      seen[key] = true;
      return true;
    });
}
