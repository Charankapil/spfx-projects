import { cacheGet, cacheKey, cacheSet } from '../../common/cache';
import { enqueue } from '../../common/RequestQueue';
import { SpClient } from '../../common/SpClient';

/**
 * Profile photos would normally be <img> tags, which the browser fetches
 * all at once and outside the page-wide queue. Instead each photo is fetched
 * through the queue, one after another, and kept as a small data URL, so a
 * person's photo costs one request per browser session at most.
 */

/** Photos stored in the browser only when they are this small (the "S" size is ~2 KB). */
const MAX_CACHED_BYTES = 12000;
const PHOTO_CACHE_MINUTES = 12 * 60;
/** A page never loads more photos than this; the rest show initials. */
export const MAX_PHOTOS_PER_PAGE = 60;

const inMemory: Record<string, Promise<string | undefined> | undefined> = {};
let loadedThisPage = 0;

function key(sp: SpClient, email: string): string {
  return cacheKey(sp.siteUrl, `photo:${email.toLowerCase()}`);
}

function toDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

/** Resolves to an image URL, or undefined when the person has no photo or the limit is reached. */
export function getPhoto(sp: SpClient, email: string): Promise<string | undefined> {
  if (!email) {
    return Promise.resolve(undefined);
  }
  const lower = email.toLowerCase();
  const running = inMemory[lower];
  if (running) {
    return running;
  }
  const stored = cacheGet<string>('local', key(sp, email));
  if (stored !== undefined) {
    const known = Promise.resolve(stored || undefined);
    inMemory[lower] = known;
    return known;
  }
  if (loadedThisPage >= MAX_PHOTOS_PER_PAGE) {
    return Promise.resolve(undefined);
  }
  loadedThisPage++;

  const url = `${sp.webUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${encodeURIComponent(email)}`;
  const request = enqueue(() => fetch(url, { credentials: 'same-origin' }))
    .then(async (response) => {
      const type = response.headers.get('Content-Type') || '';
      if (!response.ok || type.indexOf('image/') !== 0) {
        // Remember "no photo" too, so the next page does not ask again.
        cacheSet('local', key(sp, email), '', PHOTO_CACHE_MINUTES);
        return undefined;
      }
      const blob = await response.blob();
      const dataUrl = await toDataUrl(blob);
      if (blob.size <= MAX_CACHED_BYTES) {
        cacheSet('local', key(sp, email), dataUrl, PHOTO_CACHE_MINUTES);
      }
      return dataUrl;
    })
    .catch(() => {
      // A failed photo is not worth an error message; initials are shown instead.
      inMemory[lower] = undefined;
      return undefined;
    });
  inMemory[lower] = request;
  return request;
}
