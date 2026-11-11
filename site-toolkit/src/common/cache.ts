/**
 * Small time-limited cache over browser storage, so repeat page loads reuse
 * an answer instead of asking SharePoint again. Storage can be unavailable
 * (private mode, blocked site data); every access is guarded and the caller
 * simply falls back to asking SharePoint.
 */

interface IEntry<T> {
  expires: number;
  value: T;
}

export type CacheArea = 'session' | 'local';

function storage(area: CacheArea): Storage | undefined {
  try {
    return area === 'local' ? window.localStorage : window.sessionStorage;
  } catch {
    return undefined;
  }
}

export function cacheKey(scope: string, name: string): string {
  return `siteToolkit:${scope.toLowerCase()}:${name}`;
}

export function cacheGet<T>(area: CacheArea, key: string): T | undefined {
  const store = storage(area);
  if (!store) {
    return undefined;
  }
  try {
    const raw = store.getItem(key);
    if (!raw) {
      return undefined;
    }
    const entry = JSON.parse(raw) as IEntry<T>;
    if (!entry || typeof entry.expires !== 'number' || entry.expires < Date.now()) {
      store.removeItem(key);
      return undefined;
    }
    return entry.value;
  } catch {
    return undefined;
  }
}

export function cacheSet<T>(area: CacheArea, key: string, value: T, ttlMinutes: number): void {
  const store = storage(area);
  if (!store) {
    return;
  }
  try {
    const entry: IEntry<T> = { expires: Date.now() + ttlMinutes * 60 * 1000, value };
    store.setItem(key, JSON.stringify(entry));
  } catch {
    // Quota exceeded or storage blocked: carry on without the cache.
  }
}

export function cacheRemove(area: CacheArea, key: string): void {
  const store = storage(area);
  if (!store) {
    return;
  }
  try {
    store.removeItem(key);
  } catch {
    // Nothing to do.
  }
}

/** Returns the cached value, or loads it once and caches the result. */
export async function cached<T>(
  area: CacheArea,
  key: string,
  ttlMinutes: number,
  load: () => Promise<T>
): Promise<T> {
  const hit = cacheGet<T>(area, key);
  if (hit !== undefined) {
    return hit;
  }
  const value = await load();
  cacheSet(area, key, value, ttlMinutes);
  return value;
}
