/**
 * Pure helpers for the Broken Link Checker: pulling links out of a modern
 * page's stored content and deciding which ones can be checked.
 */

export interface IFoundLink {
  /** Absolute URL without its #fragment. */
  url: string;
  internal: boolean;
}

/** Every link address in a page's canvas HTML, as written (may be relative). */
export function extractHrefs(canvasHtml: string): string[] {
  if (!canvasHtml) {
    return [];
  }
  const out: string[] = [];
  let doc: Document;
  try {
    doc = new DOMParser().parseFromString(canvasHtml, 'text/html');
  } catch {
    return [];
  }

  const anchors = doc.querySelectorAll('a[href]');
  for (let i = 0; i < anchors.length; i++) {
    const href = anchors[i].getAttribute('href');
    if (href) {
      out.push(href);
    }
  }

  // Web parts such as Quick Links keep their links in a JSON attribute rather than in <a> tags.
  const parts = doc.querySelectorAll('[data-sp-webpartdata]');
  for (let i = 0; i < parts.length; i++) {
    try {
      const data = JSON.parse(parts[i].getAttribute('data-sp-webpartdata') || '{}');
      const links = data && data.serverProcessedContent && data.serverProcessedContent.links;
      if (links && typeof links === 'object') {
        for (const key of Object.keys(links)) {
          if (typeof links[key] === 'string' && links[key]) {
            out.push(links[key]);
          }
        }
      }
    } catch {
      // Not JSON: nothing to read from this web part.
    }
  }
  return out;
}

/**
 * Turns a link as written into an absolute URL and says whether it points
 * at this tenant. Returns undefined for things that are not web addresses
 * (mail, phone, script, in-page anchors).
 */
export function normalizeLink(href: string, pageUrl: string, origin: string): IFoundLink | undefined {
  const trimmed = (href || '').trim();
  if (!trimmed || trimmed.charAt(0) === '#' || /^(mailto|tel|javascript|data|sms|file|ms-[a-z]+):/i.test(trimmed)) {
    return undefined;
  }
  let absolute: URL;
  try {
    absolute = new URL(trimmed, pageUrl);
  } catch {
    return undefined;
  }
  if (absolute.protocol !== 'http:' && absolute.protocol !== 'https:') {
    return undefined;
  }
  absolute.hash = '';
  return { url: absolute.toString(), internal: absolute.origin.toLowerCase() === origin.toLowerCase() };
}

export type LinkStatus = 'ok' | 'broken' | 'restricted' | 'unknown';

/** What an HTTP status means for a link. */
export function classify(status: number): LinkStatus {
  if (status === 404 || status === 410) {
    return 'broken';
  }
  if (status === 401 || status === 403) {
    return 'restricted';
  }
  if (status >= 200 && status < 400) {
    return 'ok';
  }
  return 'unknown';
}
