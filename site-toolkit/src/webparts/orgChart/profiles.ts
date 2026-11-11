import { cacheGet, cacheKey, cacheSet } from '../../common/cache';
import { CancelToken } from '../../common/RequestQueue';
import { SpClient } from '../../common/SpClient';

/**
 * Reads people from SharePoint's own user profile service
 * (SP.UserProfiles.PeopleManager), which every signed-in user may read.
 * No Microsoft Graph permission or admin consent is involved.
 */

export interface IProfile {
  account: string;
  name: string;
  email: string;
  jobTitle: string;
  department: string;
  office: string;
  phone: string;
  /** Managers from the top of the organisation down to the direct manager. */
  managers: string[];
  reports: string[];
  peers: string[];
}

type Collection = string[] | { results?: string[] } | null | undefined;

interface IKeyValue {
  Key: string;
  Value: string | null;
}

interface IRawProfile {
  AccountName?: string | null;
  DisplayName?: string | null;
  Email?: string | null;
  Title?: string | null;
  DirectReports?: Collection;
  ExtendedManagers?: Collection;
  Peers?: Collection;
  UserProfileProperties?: IKeyValue[] | { results?: IKeyValue[] } | null;
}

/** People are cached for this long, so moving around the chart rarely asks twice. */
export const PROFILE_CACHE_MINUTES = 30;
/** SharePoint accepts at most 100 operations per batch; stay well under. */
export const BATCH_SIZE = 50;
export const MAX_REPORTS = 100;
export const MAX_PEERS = 50;

function list(collection: Collection): string[] {
  if (!collection) {
    return [];
  }
  const values = Array.isArray(collection) ? collection : collection.results || [];
  return values.filter((v) => typeof v === 'string' && v.length > 0);
}

/**
 * Turns a profile service answer (verbose or nometadata) into a profile.
 * Returns undefined for an account that has no profile, which the service
 * answers with an object full of nulls rather than an error.
 */
export function parseProfile(json: unknown): IProfile | undefined {
  const wrapper = json as { d?: IRawProfile } | undefined;
  const raw: IRawProfile | undefined = wrapper && wrapper.d ? wrapper.d : (json as IRawProfile);
  if (!raw || !raw.AccountName) {
    return undefined;
  }
  const props: Record<string, string> = {};
  const rawProps = raw.UserProfileProperties;
  const entries = !rawProps ? [] : Array.isArray(rawProps) ? rawProps : rawProps.results || [];
  for (const entry of entries) {
    if (entry && entry.Key) {
      props[entry.Key] = entry.Value || '';
    }
  }
  return {
    account: raw.AccountName,
    name: raw.DisplayName || props.PreferredName || raw.AccountName,
    email: raw.Email || props.WorkEmail || '',
    jobTitle: raw.Title || props['SPS-JobTitle'] || '',
    department: props.Department || props['SPS-Department'] || '',
    office: props.Office || props['SPS-Location'] || '',
    phone: props.WorkPhone || '',
    managers: list(raw.ExtendedManagers),
    reports: list(raw.DirectReports),
    peers: list(raw.Peers)
  };
}

export function profileUrl(sp: SpClient, account: string): string {
  return (
    `${sp.webUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)` +
    `?@v='${encodeURIComponent(account.replace(/'/g, "''"))}'`
  );
}

function key(sp: SpClient, account: string): string {
  return cacheKey(sp.siteUrl, `profile:${account.toLowerCase()}`);
}

/** The managers to show above a person: the nearest `levels`, top first. */
export function nearestManagers(profile: IProfile, levels: number): string[] {
  return levels <= 0 ? [] : profile.managers.slice(-levels);
}

export class ProfileStore {
  constructor(private sp: SpClient) {}

  /** The signed-in user's own profile: one request, then cached. */
  public async me(token?: CancelToken): Promise<IProfile | undefined> {
    const meKey = cacheKey(this.sp.siteUrl, 'profile:__me__');
    const account = cacheGet<string>('session', meKey);
    const hit = account ? this.cached(account) : undefined;
    if (hit) {
      return hit;
    }
    const json = await this.sp.getJson<unknown>(`${this.sp.webUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`, { token });
    const profile = parseProfile(json);
    if (profile) {
      this.remember(profile);
      cacheSet('session', meKey, profile.account, PROFILE_CACHE_MINUTES);
    }
    return profile;
  }

  /** A profile, `false` for an account already known to have none, or undefined when not yet asked. */
  public cached(account: string): IProfile | false | undefined {
    return cacheGet<IProfile | false>('session', key(this.sp, account));
  }

  private rememberMissing(account: string): void {
    cacheSet('session', key(this.sp, account), false, PROFILE_CACHE_MINUTES);
  }

  private remember(profile: IProfile): void {
    cacheSet('session', key(this.sp, profile.account), profile, PROFILE_CACHE_MINUTES);
  }

  /**
   * Profiles for many accounts. Cached ones cost nothing; the rest are read
   * in batches of 50 per request, one batch after another.
   * Accounts without a profile are simply left out of the result.
   */
  public async many(accounts: string[], token?: CancelToken): Promise<Record<string, IProfile>> {
    const out: Record<string, IProfile> = {};
    const missing: string[] = [];
    const seen: Record<string, boolean> = {};
    for (const account of accounts) {
      const lower = account.toLowerCase();
      if (seen[lower]) {
        continue;
      }
      seen[lower] = true;
      const hit = this.cached(account);
      if (hit) {
        out[lower] = hit;
      } else if (hit === undefined) {
        missing.push(account);
      }
    }
    for (let start = 0; start < missing.length; start += BATCH_SIZE) {
      const chunk = missing.slice(start, start + BATCH_SIZE);
      const results = await this.sp.batchGet(
        chunk.map((account) => profileUrl(this.sp, account)),
        token
      );
      results.forEach((result, index) => {
        if (!result.ok || !result.body) {
          // A 404 means the account is gone; anything else may be temporary and is asked again next time.
          if (result.status === 404) {
            this.rememberMissing(chunk[index]);
          }
          return;
        }
        let parsed: IProfile | undefined;
        try {
          parsed = parseProfile(JSON.parse(result.body));
        } catch {
          parsed = undefined;
        }
        if (parsed) {
          this.remember(parsed);
          out[chunk[index].toLowerCase()] = parsed;
        } else {
          this.rememberMissing(chunk[index]);
        }
      });
    }
    return out;
  }

  /** One person, from the cache or one request. */
  public async one(account: string, token?: CancelToken): Promise<IProfile | undefined> {
    const hit = this.cached(account);
    if (hit !== undefined) {
      return hit || undefined;
    }
    const profile = parseProfile(await this.sp.getJson<unknown>(profileUrl(this.sp, account), { token }));
    if (profile) {
      this.remember(profile);
    } else {
      this.rememberMissing(account);
    }
    return profile;
  }
}

export function initials(name: string): string {
  const parts = name
    .replace(/\(.*?\)/g, '')
    .split(/\s+/)
    .filter((p) => /^[A-Za-z0-9\u00C0-\uFFFF]/.test(p));
  if (parts.length === 0) {
    return '?';
  }
  const first = parts[0].charAt(0);
  const last = parts.length > 1 ? parts[parts.length - 1].charAt(0) : '';
  return (first + last).toUpperCase();
}
