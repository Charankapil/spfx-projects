import { ScanCancelledError, ScanPausedError } from './errors';

/**
 * How hard a scan may push SharePoint. Throttling in SharePoint Online is
 * measured in resource units per minute per user and per app, and per tenant
 * over five minutes, so one busy scan can slow down other people and tools
 * on the same tenant. The default is deliberately gentle.
 */
export type ScanSpeed = 'gentle' | 'balanced' | 'fast';

export interface ISpeedProfile {
  /** Most requests in flight at once. */
  maxConcurrency: number;
  /** Least time between the start of two requests. */
  minGapMs: number;
  /** Libraries read at the same time. */
  libraryWorkers: number;
  /** Readers per large library (one ID range each). */
  rangeReaders: number;
}

export const SPEED_PROFILES: { [speed: string]: ISpeedProfile } = {
  gentle: { maxConcurrency: 2, minGapMs: 400, libraryWorkers: 2, rangeReaders: 2 },
  balanced: { maxConcurrency: 4, minGapMs: 200, libraryWorkers: 3, rangeReaders: 2 },
  fast: { maxConcurrency: 6, minGapMs: 100, libraryWorkers: 4, rangeReaders: 3 }
};

export const DEFAULT_SPEED: ScanSpeed = 'gentle';

export interface IGovernorClock {
  now(): number;
  /** Resolves after ms; rejects with ScanCancelledError if the scan is cancelled meanwhile. */
  sleep(ms: number): Promise<void>;
}

export interface IGovernorState {
  /** Requests currently allowed in flight (lowered after throttling, raised again after quiet spells). */
  concurrency: number;
  maxConcurrency: number;
  /** Epoch ms until which no request is sent; in the past when not paused. */
  pausedUntil: number;
  /** Times SharePoint has throttled this scan. */
  throttledCount: number;
  /** Total time spent waiting because of throttling. */
  waitedMs: number;
}

export interface IRequestOutcome {
  throttled: boolean;
  retryAfterMs?: number;
}

const MAX_BACKOFF_MS = 5 * 60 * 1000;
const MAX_GAP_MS = 2000;
/** Requests in a row that must succeed before the pace is raised a notch. */
const RECOVER_AFTER = 25;

/**
 * Every request goes through here. It spaces requests out, caps how many are
 * in flight, and reacts to throttling for the whole scan at once:
 *
 *  - a 429/503 pauses *all* requests until Retry-After has passed (jittered),
 *    instead of only the one that was refused while the others carry on;
 *  - each throttle halves the allowed concurrency and widens the gap between
 *    requests; a run of successes raises them again (additive increase,
 *    multiplicative decrease);
 *  - throttling is never a failure. After `patienceMs` of accumulated waiting
 *    the scan pauses with ScanPausedError so it can be resumed later.
 */
export class RequestGovernor {
  /** Called whenever the state changes visibly (a throttle, or recovery), so the UI can show it. */
  public onChange: (() => void) | undefined;

  private inFlight = 0;
  private limit: number;
  private gap: number;
  private nextStart = 0;
  private pausedUntil = 0;
  private successStreak = 0;
  private throttleStreak = 0;
  private throttledCount = 0;
  private waitedMs = 0;
  private exhausted = false;
  private cancelled = false;
  private slotWaiters: { resolve: () => void; reject: (e: Error) => void }[] = [];

  constructor(
    private profile: ISpeedProfile,
    private clock: IGovernorClock,
    private patienceMs: number,
    private random: () => number = Math.random
  ) {
    this.limit = profile.maxConcurrency;
    this.gap = profile.minGapMs;
  }

  public state(): IGovernorState {
    return {
      concurrency: this.limit,
      maxConcurrency: this.profile.maxConcurrency,
      pausedUntil: this.pausedUntil,
      throttledCount: this.throttledCount,
      waitedMs: this.waitedMs
    };
  }

  public cancel(): void {
    this.cancelled = true;
    const waiters = this.slotWaiters;
    this.slotWaiters = [];
    waiters.forEach((w) => w.reject(new ScanCancelledError('Scan cancelled.')));
  }

  /** Resolves when a request may start. Reserve it with release() when the request ends. */
  public async acquire(): Promise<void> {
    for (;;) {
      if (this.cancelled) {
        throw new ScanCancelledError('Scan cancelled.');
      }
      if (this.exhausted) {
        throw new ScanPausedError(
          'SharePoint kept throttling this scan, so it paused. Your progress is saved; resume it later, ideally outside working hours.'
        );
      }
      const now = this.clock.now();
      if (now < this.pausedUntil) {
        // In slices, so a scan that has run out of patience stops promptly instead of after the whole pause.
        await this.clock.sleep(Math.min(this.pausedUntil - now, 1000));
        continue;
      }
      if (this.inFlight >= this.limit) {
        await new Promise<void>((resolve, reject) => this.slotWaiters.push({ resolve, reject }));
        continue;
      }
      const wait = this.nextStart - now;
      if (wait > 0) {
        await this.clock.sleep(wait);
        continue;
      }
      this.inFlight++;
      this.nextStart = this.clock.now() + this.gap;
      return;
    }
  }

  public release(outcome: IRequestOutcome): void {
    this.inFlight--;
    if (outcome.throttled) {
      this.onThrottled(outcome.retryAfterMs);
    } else {
      this.onSucceeded();
    }
    // A freed slot (or a lowered limit that now allows nobody) is re-checked by the waiters themselves.
    const next = this.slotWaiters.shift();
    if (next) {
      next.resolve();
    }
  }

  private onThrottled(retryAfterMs: number | undefined): void {
    this.throttleStreak++;
    this.throttledCount++;
    this.successStreak = 0;
    // Slow down for the rest of the scan, not just for this request.
    this.limit = Math.max(1, Math.floor(this.limit / 2));
    this.gap = Math.min(MAX_GAP_MS, Math.round(this.gap * 1.5));

    const backoff = Math.min(MAX_BACKOFF_MS, 5000 * Math.pow(2, this.throttleStreak - 1));
    const base = retryAfterMs !== undefined && retryAfterMs >= 0 ? Math.min(retryAfterMs, MAX_BACKOFF_MS) : backoff;
    // Jitter (0-20%) so several waiting requests do not all return at the same instant.
    const delay = Math.round(base * (1 + this.random() * 0.2));
    const now = this.clock.now();
    const until = now + delay;
    if (until > this.pausedUntil) {
      // Count only the extension of the pause, so overlapping 429s are not double counted.
      this.waitedMs += until - Math.max(now, this.pausedUntil);
      this.pausedUntil = until;
    }
    if (this.waitedMs > this.patienceMs) {
      this.exhausted = true;
    }
    this.notify();
  }

  private onSucceeded(): void {
    this.throttleStreak = 0;
    if (++this.successStreak >= RECOVER_AFTER) {
      this.successStreak = 0;
      const recovered = this.limit < this.profile.maxConcurrency || this.gap > this.profile.minGapMs;
      this.limit = Math.min(this.profile.maxConcurrency, this.limit + 1);
      this.gap = Math.max(this.profile.minGapMs, Math.round(this.gap * 0.85));
      if (recovered) {
        this.notify();
      }
    }
  }

  private notify(): void {
    if (this.onChange) {
      this.onChange();
    }
  }
}

/** Retry-After as milliseconds: whole seconds, or an HTTP date. undefined if absent or unreadable. */
export function parseRetryAfter(header: string | null | undefined, now: number): number | undefined {
  if (!header) {
    return undefined;
  }
  const seconds = parseInt(header, 10);
  if (!isNaN(seconds) && /^\s*\d+\s*$/.test(header)) {
    return Math.max(0, seconds) * 1000;
  }
  const date = Date.parse(header);
  return isNaN(date) ? undefined : Math.max(0, date - now);
}
