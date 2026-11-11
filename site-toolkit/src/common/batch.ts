/**
 * Builds and reads SharePoint `$batch` requests by hand. SPFx 1.18 does not
 * expose its own batch helper publicly, and the format is small: every
 * update travels in its own change set, so one failing update never takes
 * the others down with it.
 */

export interface IBatchUpdate {
  /** Absolute URL of the entity to update. */
  url: string;
  body: Record<string, unknown>;
}

export interface IBatchResult {
  ok: boolean;
  status: number;
  error?: string;
  /** The response body of this operation (reads only). */
  body?: string;
  /** Seconds SharePoint asked to wait, when this operation was throttled. */
  retryAfter?: number;
}

const CRLF = '\r\n';

export function newBatchId(): string {
  let id = '';
  for (let i = 0; i < 32; i++) {
    id += Math.floor(Math.random() * 16).toString(16);
  }
  return id;
}

export function buildBatchBody(updates: IBatchUpdate[], batchId: string): string {
  const lines: string[] = [];
  updates.forEach((update, index) => {
    const changeSet = `changeset_${batchId}_${index}`;
    lines.push(`--batch_${batchId}`);
    lines.push(`Content-Type: multipart/mixed; boundary="${changeSet}"`);
    lines.push('');
    lines.push(`--${changeSet}`);
    lines.push('Content-Type: application/http');
    lines.push('Content-Transfer-Encoding: binary');
    lines.push('');
    lines.push(`PATCH ${update.url} HTTP/1.1`);
    lines.push('Accept: application/json');
    lines.push('Content-Type: application/json;charset=utf-8');
    lines.push('IF-MATCH: *');
    lines.push('');
    lines.push(JSON.stringify(update.body));
    lines.push('');
    lines.push(`--${changeSet}--`);
    lines.push('');
  });
  lines.push(`--batch_${batchId}--`);
  lines.push('');
  return lines.join(CRLF);
}

/** A batch of reads. Reads need no change sets: each is its own part. */
export function buildBatchGetBody(urls: string[], batchId: string): string {
  const lines: string[] = [];
  for (const url of urls) {
    lines.push(`--batch_${batchId}`);
    lines.push('Content-Type: application/http');
    lines.push('Content-Transfer-Encoding: binary');
    lines.push('');
    lines.push(`GET ${url} HTTP/1.1`);
    lines.push('Accept: application/json;odata=nometadata');
    lines.push('');
  }
  lines.push(`--batch_${batchId}--`);
  lines.push('');
  return lines.join(CRLF);
}

/** The body of one inner response: everything after its headers, up to the next boundary. */
function bodyFrom(block: string): string {
  const headersEnd = block.search(/\r?\n\r?\n/);
  if (headersEnd < 0) {
    return '';
  }
  const rest = block.substring(headersEnd).replace(/^\r?\n\r?\n/, '');
  const boundary = rest.search(/(^|\r?\n)--batchresponse/);
  return (boundary >= 0 ? rest.substring(0, boundary) : rest).trim();
}

function errorFrom(block: string): string | undefined {
  const start = block.indexOf('{');
  const end = block.lastIndexOf('}');
  if (start < 0 || end <= start) {
    return undefined;
  }
  try {
    const parsed = JSON.parse(block.substring(start, end + 1));
    const error = parsed.error || parsed['odata.error'];
    const message = error && error.message;
    if (typeof message === 'string') {
      return message;
    }
    if (message && typeof message.value === 'string') {
      return message.value;
    }
  } catch {
    // Not JSON: fall through.
  }
  return undefined;
}

/**
 * One result per operation, in the order the operations were sent. The
 * response is a multipart document; each part carries an inner HTTP
 * response whose status line is what matters.
 */
export function parseBatchResponse(text: string): IBatchResult[] {
  const results: IBatchResult[] = [];
  const statusLine = /^HTTP\/1\.1 (\d{3})[^\r\n]*/gm;
  const starts: Array<{ index: number; status: number }> = [];
  let match = statusLine.exec(text);
  while (match) {
    starts.push({ index: match.index, status: parseInt(match[1], 10) });
    match = statusLine.exec(text);
  }
  starts.forEach((start, i) => {
    const block = text.substring(start.index, i + 1 < starts.length ? starts[i + 1].index : text.length);
    const ok = start.status >= 200 && start.status < 300;
    const result: IBatchResult = { ok, status: start.status };
    const body = bodyFrom(block);
    if (ok && body) {
      result.body = body;
    }
    if (!ok) {
      result.error = errorFrom(block) || `SharePoint answered ${start.status}.`;
      const retry = /^Retry-After:\s*(\d+)/im.exec(block);
      if (retry) {
        result.retryAfter = parseInt(retry[1], 10);
      }
    }
    results.push(result);
  });
  return results;
}
