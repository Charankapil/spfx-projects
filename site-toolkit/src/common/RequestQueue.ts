/**
 * One request queue for the whole page.
 *
 * Every SharePoint call made by any Site Toolkit component goes through here,
 * so the page as a whole stays far below SharePoint's per-user limit
 * (documented as 3,000 requests per 5 minutes):
 *
 *  - requests are sent strictly one at a time, with a gap between them;
 *  - a rolling 5-minute budget caps the total, whatever the components do;
 *  - a 429 or 503 pauses the whole queue for the Retry-After time, then the
 *    same request is retried, and the gap is doubled for the rest of the
 *    session.
 *
 * Each web part is bundled separately, so the state lives on `window` rather
 * than in a module variable: that is what lets two web parts on one page
 * share a single budget.
 */

export const BASE_GAP_MS = 500;
export const MAX_GAP_MS = 4000;
export const BUDGET_WINDOW_MS = 5 * 60 * 1000;
export const BUDGET_REQUESTS = 600;
export const MAX_THROTTLE_RETRIES = 5;
const DEFAULT_RETRY_AFTER_SEC = 10;
const MAX_RETRY_AFTER_SEC = 600;

export interface IThrottleAwareResponse {
  status: number;
  headers: { get(name: string): string | null };
}

export interface IQueueSnapshot {
  /** Epoch ms until which the queue is paused because SharePoint throttled it; 0 when running. */
  pausedUntil: number;
  gapMs: number;
  throttleCount: number;
  sentInWindow: number;
}

interface IQueueState {
  chain: Promise<void>;
  gapMs: number;
  lastFinishedAt: number;
  pausedUntil: number;
  sent: number[];
  throttleCount: number;
  listeners: Array<(snapshot: IQueueSnapshot) => void>;
}

export interface IToolkitError extends Error {
  code?: 'cancelled' | 'throttled' | 'http';
  status?: number;
}

/**
 * Errors are tagged with a code instead of subclassing Error: the build
 * targets ES5, where `instanceof` on an Error subclass is always false.
 */
export function toolkitError(message: string, code: IToolkitError['code'], status?: number): IToolkitError {
  const error: IToolkitError = new Error(message);
  error.code = code;
  error.status = status;
  return error;
}

export function isCancelled(err: unknown): boolean {
  return !!err && (err as IToolkitError).code === 'cancelled';
}

export function httpStatusOf(err: unknown): number | undefined {
  return err ? (err as IToolkitError).status : undefined;
}

export function errorText(err: unknown): string {
  if (err instanceof Error && err.message) {
    return err.message;
  }
  return 'Something went wrong.';
}

/** Lets a component stop a long-running scan; checked before every request. */
export class CancelToken {
  public cancelled = false;

  public cancel(): void {
    this.cancelled = true;
  }

  public throwIfCancelled(): void {
    if (this.cancelled) {
      throw toolkitError('Cancelled.', 'cancelled');
    }
  }
}

const STATE_KEY = '__siteToolkitRequestQueue_v1';

interface IClock {
  now(): number;
  sleep(ms: number): Promise<void>;
}

const realClock: IClock = {
  now: () => Date.now(),
  sleep: (ms) => new Promise<void>((resolve) => setTimeout(resolve, ms))
};

let clock: IClock = realClock;

/** Test hook: swaps real time for a fake clock. */
export function setClockForTests(fake: IClock | undefined): void {
  clock = fake || realClock;
}

function getState(): IQueueState {
  const host = (typeof window !== 'undefined' ? window : globalThis) as unknown as Record<string, IQueueState>;
  if (!host[STATE_KEY]) {
    host[STATE_KEY] = {
      chain: Promise.resolve(),
      gapMs: BASE_GAP_MS,
      lastFinishedAt: 0,
      pausedUntil: 0,
      sent: [],
      throttleCount: 0,
      listeners: []
    };
  }
  return host[STATE_KEY];
}

/** Test hook: forgets all queue state. */
export function resetQueueForTests(): void {
  const host = (typeof window !== 'undefined' ? window : globalThis) as unknown as Record<string, unknown>;
  delete host[STATE_KEY];
}

export function getQueueSnapshot(): IQueueSnapshot {
  const state = getState();
  const now = clock.now();
  return {
    pausedUntil: state.pausedUntil > now ? state.pausedUntil : 0,
    gapMs: state.gapMs,
    throttleCount: state.throttleCount,
    sentInWindow: state.sent.filter((t) => t > now - BUDGET_WINDOW_MS).length
  };
}

export function subscribeToQueue(listener: (snapshot: IQueueSnapshot) => void): () => void {
  const state = getState();
  state.listeners.push(listener);
  return () => {
    const index = state.listeners.indexOf(listener);
    if (index >= 0) {
      state.listeners.splice(index, 1);
    }
  };
}

function notify(state: IQueueState): void {
  const snapshot = getQueueSnapshot();
  for (const listener of state.listeners.slice()) {
    try {
      listener(snapshot);
    } catch {
      // A broken listener must never stop the queue.
    }
  }
}

function parseRetryAfter(response: IThrottleAwareResponse, attempt: number): number {
  let header: string | null = null;
  try {
    header = response.headers.get('Retry-After');
  } catch {
    header = null;
  }
  const seconds = header ? parseInt(header, 10) : NaN;
  if (!isNaN(seconds) && seconds > 0) {
    return Math.min(seconds, MAX_RETRY_AFTER_SEC);
  }
  // No usable header: back off exponentially instead.
  return Math.min(DEFAULT_RETRY_AFTER_SEC * Math.pow(2, attempt), MAX_RETRY_AFTER_SEC);
}

async function waitForTurn(state: IQueueState, token: CancelToken | undefined, weight: number): Promise<void> {
  for (;;) {
    if (token) {
      token.throwIfCancelled();
    }
    const now = clock.now();
    let readyAt = Math.max(state.pausedUntil, state.lastFinishedAt + state.gapMs);

    state.sent = state.sent.filter((t) => t > now - BUDGET_WINDOW_MS);
    if (state.sent.length > 0 && state.sent.length + weight > BUDGET_REQUESTS) {
      // Wait until enough of the oldest requests have left the window to make room.
      const mustLeave = Math.min(state.sent.length, state.sent.length + weight - BUDGET_REQUESTS);
      readyAt = Math.max(readyAt, state.sent[mustLeave - 1] + BUDGET_WINDOW_MS);
    }

    if (readyAt <= now) {
      return;
    }
    // Sleep in short steps so a cancel is noticed promptly during a long pause.
    await clock.sleep(Math.min(readyAt - now, 1000));
  }
}

async function execute<T extends IThrottleAwareResponse>(
  send: () => Promise<T>,
  token: CancelToken | undefined,
  weight: number
): Promise<T> {
  const state = getState();
  for (let attempt = 0; ; attempt++) {
    await waitForTurn(state, token, weight);

    // A batch carries many operations in one HTTP call; each one counts against the budget.
    for (let i = 0; i < weight; i++) {
      state.sent.push(clock.now());
    }
    let response: T;
    try {
      response = await send();
    } finally {
      state.lastFinishedAt = clock.now();
    }

    if (response.status !== 429 && response.status !== 503) {
      return response;
    }

    state.throttleCount++;
    state.gapMs = Math.min(state.gapMs * 2, MAX_GAP_MS);
    state.pausedUntil = clock.now() + parseRetryAfter(response, attempt) * 1000;
    notify(state);

    if (attempt >= MAX_THROTTLE_RETRIES) {
      throw toolkitError(
        'SharePoint is busy and asked this page to slow down several times. Please try again in a few minutes.',
        'throttled',
        response.status
      );
    }
  }
}

/**
 * Sends one request through the page-wide queue. `send` may be called more
 * than once if SharePoint throttles it, so it must be safe to repeat.
 * `weight` is the number of operations the request carries (1 unless it is a batch).
 */
export function enqueue<T extends IThrottleAwareResponse>(
  send: () => Promise<T>,
  token?: CancelToken,
  weight = 1
): Promise<T> {
  const state = getState();
  const run = state.chain.then(() => execute(send, token, Math.max(1, Math.min(weight, BUDGET_REQUESTS))));
  state.chain = run.then(
    () => undefined,
    () => undefined
  );
  return run;
}
