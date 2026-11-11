import assert from 'node:assert/strict';
import { test } from 'node:test';

import {
  BASE_GAP_MS,
  BUDGET_REQUESTS,
  BUDGET_WINDOW_MS,
  CancelToken,
  enqueue,
  getQueueSnapshot,
  isCancelled,
  MAX_THROTTLE_RETRIES,
  subscribeToQueue
} from '../src/common/RequestQueue';
import { response, useFakeClock } from './helpers';

test('requests never overlap, even when queued together', async () => {
  const clock = useFakeClock();
  let running = 0;
  let peak = 0;
  const send = async (): Promise<ReturnType<typeof response>> => {
    running++;
    peak = Math.max(peak, running);
    await clock.sleep(40);
    running--;
    return response(200);
  };
  await Promise.all([1, 2, 3, 4, 5, 6, 7, 8].map(() => enqueue(send)));
  assert.equal(peak, 1);
});

test('a gap is left between one request finishing and the next starting', async () => {
  const clock = useFakeClock();
  const started: number[] = [];
  const finished: number[] = [];
  const send = async (): Promise<ReturnType<typeof response>> => {
    started.push(clock.now());
    await clock.sleep(100);
    finished.push(clock.now());
    return response(200);
  };
  await Promise.all([enqueue(send), enqueue(send), enqueue(send)]);
  for (let i = 1; i < started.length; i++) {
    assert.ok(started[i] - finished[i - 1] >= BASE_GAP_MS, `gap before request ${i} was ${started[i] - finished[i - 1]} ms`);
  }
});

test('the first request is not delayed', async () => {
  const clock = useFakeClock();
  const before = clock.now();
  let startedAt = 0;
  await enqueue(async () => {
    startedAt = clock.now();
    return response(200);
  });
  assert.equal(startedAt, before);
});

test('a 429 waits for Retry-After, repeats the same request and slows the queue down', async () => {
  const clock = useFakeClock();
  const attempts: number[] = [];
  const result = await enqueue(async () => {
    attempts.push(clock.now());
    return attempts.length === 1 ? response(429, '', { 'Retry-After': '7' }) : response(200, { ok: true });
  });
  assert.equal(result.status, 200);
  assert.equal(attempts.length, 2);
  assert.ok(attempts[1] - attempts[0] >= 7000, `waited only ${attempts[1] - attempts[0]} ms`);
  const snapshot = getQueueSnapshot();
  assert.equal(snapshot.throttleCount, 1);
  assert.equal(snapshot.gapMs, BASE_GAP_MS * 2);
});

test('a 503 without Retry-After backs off, longer each time', async () => {
  const clock = useFakeClock();
  const attempts: number[] = [];
  await enqueue(async () => {
    attempts.push(clock.now());
    return attempts.length <= 2 ? response(503) : response(200);
  });
  const first = attempts[1] - attempts[0];
  const second = attempts[2] - attempts[1];
  assert.ok(first >= 10000);
  assert.ok(second >= 20000, `second wait was ${second} ms`);
});

test('while throttled, every other request on the page waits too', async () => {
  const clock = useFakeClock();
  let firstTry = true;
  let throttledAt = 0;
  const throttled = enqueue(async () => {
    if (firstTry) {
      firstTry = false;
      throttledAt = clock.now();
      return response(429, '', { 'Retry-After': '30' });
    }
    return response(200);
  });
  let otherAt = 0;
  const other = enqueue(async () => {
    otherAt = clock.now();
    return response(200);
  });
  await Promise.all([throttled, other]);
  assert.ok(otherAt - throttledAt >= 30000, 'the second request went out during the pause');
});

test('gives up with a clear error after repeated throttling', async () => {
  useFakeClock();
  let attempts = 0;
  await assert.rejects(
    enqueue(async () => {
      attempts++;
      return response(429, '', { 'Retry-After': '1' });
    }),
    (err: Error & { code?: string }) => err.code === 'throttled'
  );
  assert.equal(attempts, MAX_THROTTLE_RETRIES + 1);
});

test('never sends more than the budget in any five minutes', async () => {
  const clock = useFakeClock();
  const sentAt: number[] = [];
  const send = async (): Promise<ReturnType<typeof response>> => {
    sentAt.push(clock.now());
    return response(200);
  };
  for (let i = 0; i < BUDGET_REQUESTS + 50; i++) {
    await enqueue(send);
    // Pretend the caller is instant and the gap is the only brake, then jump ahead to stress the budget.
    if (i === 10) {
      clock.time += 1;
    }
  }
  for (let i = BUDGET_REQUESTS; i < sentAt.length; i++) {
    assert.ok(
      sentAt[i] - sentAt[i - BUDGET_REQUESTS] >= BUDGET_WINDOW_MS,
      `request ${i} was the ${BUDGET_REQUESTS + 1}th inside five minutes`
    );
  }
});

test('the budget holds even if the gap is somehow skipped', async () => {
  const clock = useFakeClock();
  const sentAt: number[] = [];
  // Requests that take "negative" time would defeat the gap; the rolling budget must still hold.
  for (let i = 0; i < BUDGET_REQUESTS + 5; i++) {
    await enqueue(async () => {
      sentAt.push(clock.now());
      clock.time -= BASE_GAP_MS;
      return response(200);
    });
  }
  const last = sentAt.length - 1;
  assert.ok(sentAt[last] - sentAt[last - BUDGET_REQUESTS] >= BUDGET_WINDOW_MS);
});

test('a batch counts as many operations against the budget', async () => {
  useFakeClock();
  await enqueue(async () => response(200), undefined, 50);
  assert.equal(getQueueSnapshot().sentInWindow, 50);
});

test('a batch that does not fit waits for room instead of overshooting', async () => {
  const clock = useFakeClock();
  await enqueue(async () => response(200), undefined, BUDGET_REQUESTS - 10);
  const start = clock.now();
  let sentAt = 0;
  await enqueue(
    async () => {
      sentAt = clock.now();
      return response(200);
    },
    undefined,
    50
  );
  assert.ok(sentAt - start >= BUDGET_WINDOW_MS - 1000, 'the batch went out although the budget was nearly spent');
});

test('cancelling rejects the request and leaves the queue working', async () => {
  useFakeClock();
  const token = new CancelToken();
  token.cancel();
  let sent = false;
  await assert.rejects(
    enqueue(async () => {
      sent = true;
      return response(200);
    }, token),
    (err: unknown) => isCancelled(err)
  );
  assert.equal(sent, false);
  const after = await enqueue(async () => response(200));
  assert.equal(after.status, 200);
});

test('cancelling during a throttle pause stops the retry', async () => {
  useFakeClock();
  const token = new CancelToken();
  let attempts = 0;
  await assert.rejects(
    enqueue(async () => {
      attempts++;
      token.cancel();
      return response(429, '', { 'Retry-After': '60' });
    }, token),
    (err: unknown) => isCancelled(err)
  );
  assert.equal(attempts, 1);
});

test('a request that fails outright does not block the ones behind it', async () => {
  useFakeClock();
  const failing = enqueue(async () => {
    throw new Error('network down');
  });
  const next = enqueue(async () => response(200));
  await assert.rejects(failing, /network down/);
  assert.equal((await next).status, 200);
});

test('listeners hear about a throttle, and a broken listener is ignored', async () => {
  useFakeClock();
  const seen: number[] = [];
  subscribeToQueue(() => {
    throw new Error('broken listener');
  });
  const unsubscribe = subscribeToQueue((snapshot) => seen.push(snapshot.pausedUntil));
  let first = true;
  await enqueue(async () => {
    if (first) {
      first = false;
      return response(429, '', { 'Retry-After': '5' });
    }
    return response(200);
  });
  unsubscribe();
  assert.equal(seen.length, 1);
  assert.ok(seen[0] > 0);
});

test('an absurd Retry-After is capped', async () => {
  const clock = useFakeClock();
  const attempts: number[] = [];
  await enqueue(async () => {
    attempts.push(clock.now());
    return attempts.length === 1 ? response(429, '', { 'Retry-After': '999999' }) : response(200);
  });
  assert.ok(attempts[1] - attempts[0] <= 601000);
});
