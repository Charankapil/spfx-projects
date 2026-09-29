/**
 * Update bodies are plain text. A mention is stored as @[Display Name](123) where 123 is the
 * site user id; the ids also go into the WB_Mentions people field so the Inbox can query them.
 */

export type Segment = { kind: 'text'; text: string } | { kind: 'mention'; id: number; name: string };

const MENTION = /@\[([^\]\n]{1,120})\]\((\d{1,10})\)/g;

export function mentionToken(name: string, id: number): string {
  return `@[${name.replace(/[[\]\n]/g, '')}](${id})`;
}

export function parseSegments(body: string): Segment[] {
  const out: Segment[] = [];
  let last = 0;
  MENTION.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = MENTION.exec(body)) !== null) {
    if (m.index > last) {
      out.push({ kind: 'text', text: body.substring(last, m.index) });
    }
    out.push({ kind: 'mention', name: m[1], id: parseInt(m[2], 10) });
    last = m.index + m[0].length;
  }
  if (last < body.length) {
    out.push({ kind: 'text', text: body.substring(last) });
  }
  return out;
}

export function mentionIds(body: string): number[] {
  const ids: number[] = [];
  parseSegments(body).forEach(s => {
    if (s.kind === 'mention' && ids.indexOf(s.id) < 0) {
      ids.push(s.id);
    }
  });
  return ids;
}

/** Body as readable text: "@Anna Smith, can you check?" */
export function plainText(body: string): string {
  return parseSegments(body).map(s => (s.kind === 'text' ? s.text : '@' + s.name)).join('');
}
