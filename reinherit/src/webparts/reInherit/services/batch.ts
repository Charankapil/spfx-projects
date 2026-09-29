/**
 * Builds and parses SharePoint REST $batch (OData multipart/mixed) payloads.
 * Kept free of SPFx imports so the format can be exercised on its own.
 *
 * Every write goes in its own changeset: SharePoint does not roll a changeset
 * back when one of its requests fails, but a changeset per request keeps the
 * responses one-to-one with the requests and makes each failure attributable
 * to exactly one object.
 */

export interface IBatchRequest {
  method: 'GET' | 'POST';
  /** Absolute URL of the REST call. */
  url: string;
  body?: string;
}

export interface IBatchResponse {
  status: number;
  statusText: string;
  body: string;
}

export function buildBatchBody(batchId: string, requests: IBatchRequest[]): string {
  const lines: string[] = [];
  requests.forEach((request, index) => {
    if (request.method === 'GET') {
      lines.push(
        `--batch_${batchId}`,
        'Content-Type: application/http',
        'Content-Transfer-Encoding: binary',
        '',
        `GET ${request.url} HTTP/1.1`,
        'Accept: application/json;odata=nometadata',
        ''
      );
      return;
    }
    const changeset = `${batchId}_${index}`;
    lines.push(
      `--batch_${batchId}`,
      `Content-Type: multipart/mixed; boundary="changeset_${changeset}"`,
      'Content-Transfer-Encoding: binary',
      '',
      `--changeset_${changeset}`,
      'Content-Type: application/http',
      'Content-Transfer-Encoding: binary',
      '',
      `POST ${request.url} HTTP/1.1`,
      'Accept: application/json;odata=nometadata',
      'Content-Type: application/json;odata=verbose',
      '',
      request.body || '',
      `--changeset_${changeset}--`
    );
  });
  lines.push(`--batch_${batchId}--`, '');
  return lines.join('\r\n');
}

const BOUNDARY_LINE = /^--(?:batchresponse|changesetresponse)_[^\r\n]*$/;
const STATUS_LINE = /^HTTP\/1\.1 (\d{3}) ?(.*)$/;

/**
 * Returns one entry per inner response, in request order. Works whether
 * SharePoint nests changeset responses inside the batch response or (as it
 * usually does) returns them flat: any part without an HTTP status line, such
 * as a nested multipart header, is skipped.
 */
export function parseBatchResponse(text: string): IBatchResponse[] {
  const lines = text.split(/\r?\n/);
  const parts: string[][] = [];
  let current: string[] | undefined;
  for (const line of lines) {
    if (BOUNDARY_LINE.test(line)) {
      current = [];
      parts.push(current);
    } else if (current) {
      current.push(line);
    }
  }

  const responses: IBatchResponse[] = [];
  for (const part of parts) {
    let i = 0;
    let match: RegExpExecArray | null = null;
    for (; i < part.length; i++) {
      match = STATUS_LINE.exec(part[i]);
      if (match) {
        break;
      }
    }
    if (!match) {
      continue;
    }
    // Skip the response's own headers; the body follows the first blank line.
    i++;
    while (i < part.length && part[i].trim() !== '') {
      i++;
    }
    const body = part.slice(i + 1).join('\n').trim();
    responses.push({ status: parseInt(match[1], 10), statusText: match[2].trim(), body });
  }
  return responses;
}

/** SharePoint's error message from a JSON error body, or a trimmed raw body. */
export function errorMessageFromBody(body: string, fallback: string): string {
  if (!body) {
    return fallback;
  }
  try {
    const parsed = JSON.parse(body);
    const message = parsed?.error?.message ?? parsed?.['odata.error']?.message;
    if (typeof message === 'string') {
      return message;
    }
    if (message && typeof message.value === 'string') {
      return message.value;
    }
  } catch {
    // Not JSON - fall through to the raw text.
  }
  return body.substring(0, 300);
}
