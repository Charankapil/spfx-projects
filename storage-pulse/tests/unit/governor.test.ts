import assert from 'node:assert/strict';
import { test } from 'node:test';

import { ScanCancelledError, ScanPausedError } from '../../src/webparts/storagePulse/services/errors';
import {
  IGovernorClock,
  ISpeedProfile,
  parseRetryAfter,
  RequestGovernor,
  SPEED_PROFILES
} from '../../src/webparts/storagePulse/services/RequestGovernor';

/** Virtual time: sleeping advances the clock instantly, so pacing can be checked exactly and fast. */
function fakeClock(): IGovernorClock & { t: number } {
  const clock = {
    t: 1000000,
    now: () => clock.t,
    sleep: (ms: number) => {
      clock.t += Math.max(1, ms);
      return Promise.resolve();
    }
  };
  return clock;
}

const PROFILE: ISpeedProfile = { maxConcurrency: 4, minGapMs: 100, libraryWorkers: 2, rangeReaders: 2 };
const noJitter = (): number => 0;

test('requests start at least minGapMs apart', async () => {
  const clock = fakeClock();
  const gov = new RequestGovernor(PROFILE, clock, 1e9, noJitter);
  const starts: number[] = [];
  for (let i = 0; i < 8; i++) {
    await gov.acquire();
    starts.push(clock.now());
    gov.release({ throttled: false });
  }
  for (let i = 1; i < starts.length; i++) {
    assert.ok(starts[i] - starts[i - 1] >= PROFILE.minGapMs, `gap ${starts[i] - starts[i - 1]}`);
  }
});

test('never more than the allowed number of requests in flight', async () => {
  const clock = fakeClock();
  const gov = new RequestGovernor({ ...PROFILE, minGapMs: 0 }, clock, 1e9, noJitter);
  let inFlight = 0;
  let max = 0;
  const work = async (): Promise<void> => {
    for (let i = 0; i < 20; i++) {
      await gov.acquire();
      inFlight++;
      max = Math.max(max, inFlight);
      await Promise.resolve();
      await Promise.resolve();
      inFlight--;
      gov.release({ throttled: false });
    }
  };
  await Promise.all([work(), work(), work(), work(), work(), work(), work(), work()]);
  assert.ok(max <= 4, `max in flight ${max}`);
  assert.ok(max >= 2, 'requests did not overlap at all');
});

test('a throttle holds every request until Retry-After has passed', async () => {
  const clock = fakeClock();
  const gov = new RequestGovernor({ ...PROFILE, minGapMs: 0 }, clock, 1e9, noJitter);
  await gov.acquire();
  const throttledAt = clock.now();
  gov.release({ throttled: true, retryAfterMs: 30000 });
  // Several workers all want to send: none may start before the hold ends.
  const starts: number[] = [];
  await Promise.all(
    [1, 2, 3].map(async () => {
      await gov.acquire();
      starts.push(clock.now());
      gov.release({ throttled: false });
    })
  );
  assert.ok(Math.min(...starts) >= throttledAt + 30000, `started ${Math.min(...starts) - throttledAt} ms after the throttle`);
});

test('throttling halves the allowed concurrency and widens the gap; a run of successes restores them', async () => {
  const clock = fakeClock();
  const gov = new RequestGovernor(PROFILE, clock, 1e9, noJitter);
  assert.equal(gov.state().concurrency, 4);
  await gov.acquire();
  gov.release({ throttled: true, retryAfterMs: 1000 });
  assert.equal(gov.state().concurrency, 2);
  await gov.acquire();
  gov.release({ throttled: true, retryAfterMs: 1000 });
  assert.equal(gov.state().concurrency, 1);
  assert.equal(gov.state().throttledCount, 2);

  let changes = 0;
  gov.onChange = () => changes++;
  for (let i = 0; i < 25 * 3 + 5; i++) {
    await gov.acquire();
    gov.release({ throttled: false });
  }
  assert.equal(gov.state().concurrency, 4, 'concurrency did not recover');
  assert.ok(changes >= 3, 'recovery was not announced');
});

test('without Retry-After the wait backs off exponentially, capped at five minutes', async () => {
  const clock = fakeClock();
  const gov = new RequestGovernor(PROFILE, clock, 1e12, noJitter);
  const waits: number[] = [];
  for (let i = 0; i < 9; i++) {
    await gov.acquire();
    const before = clock.now();
    gov.release({ throttled: true });
    waits.push(gov.state().pausedUntil - before);
  }
  assert.equal(waits[0], 5000);
  assert.equal(waits[1], 10000);
  assert.equal(waits[2], 20000);
  assert.equal(waits[8], 300000);
});

test('after the patience is used up the scan pauses instead of failing', async () => {
  const clock = fakeClock();
  const gov = new RequestGovernor(PROFILE, clock, 20000, noJitter);
  await gov.acquire();
  gov.release({ throttled: true, retryAfterMs: 15000 });
  await gov.acquire(); // still within patience: waits, then proceeds
  gov.release({ throttled: true, retryAfterMs: 15000 });
  await assert.rejects(gov.acquire(), ScanPausedError);
});

test('Cancel ends waiting requests at once', async () => {
  // Real clock so acquire() really waits.
  const clock: IGovernorClock = { now: () => Date.now(), sleep: (ms) => new Promise((r) => setTimeout(r, ms)) };
  const gov = new RequestGovernor(PROFILE, clock, 1e9, noJitter);
  await gov.acquire();
  gov.release({ throttled: true, retryAfterMs: 60000 });
  const waiting = gov.acquire();
  setTimeout(() => gov.cancel(), 50);
  const started = Date.now();
  await assert.rejects(waiting, ScanCancelledError);
  assert.ok(Date.now() - started < 3000);
});

test('the default profile is the gentlest', () => {
  assert.ok(SPEED_PROFILES.gentle.maxConcurrency < SPEED_PROFILES.balanced.maxConcurrency);
  assert.ok(SPEED_PROFILES.balanced.maxConcurrency < SPEED_PROFILES.fast.maxConcurrency);
  assert.ok(SPEED_PROFILES.gentle.minGapMs > SPEED_PROFILES.fast.minGapMs);
});

test('parseRetryAfter reads seconds and dates, and ignores nonsense', () => {
  const now = Date.parse('2026-09-29T10:00:00Z');
  assert.equal(parseRetryAfter('30', now), 30000);
  assert.equal(parseRetryAfter('0', now), 0);
  assert.equal(parseRetryAfter('Tue, 29 Sep 2026 10:01:00 GMT', now), 60000);
  assert.equal(parseRetryAfter('Tue, 29 Sep 2026 09:00:00 GMT', now), 0);
  assert.equal(parseRetryAfter('soon', now), undefined);
  assert.equal(parseRetryAfter('', now), undefined);
  assert.equal(parseRetryAfter(null, now), undefined);
});
