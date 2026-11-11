/** Distinct, non-empty answers in the order they were typed. */
export function parseOptions(text: string): string[] {
  const seen: Record<string, boolean> = {};
  const out: string[] = [];
  for (const line of (text || '').split('\n')) {
    const option = line.trim().substring(0, 255);
    if (option && !seen[option.toLowerCase()]) {
      seen[option.toLowerCase()] = true;
      out.push(option);
    }
  }
  return out.slice(0, 10);
}

export function tally(options: string[], votes: string[]): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const option of options) {
    counts[option] = 0;
  }
  for (const vote of votes) {
    if (Object.prototype.hasOwnProperty.call(counts, vote)) {
      counts[vote]++;
    }
  }
  return counts;
}
