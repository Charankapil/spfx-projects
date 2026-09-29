/**
 * Runs `worker` over `items` with at most `concurrency` in flight. Unlike
 * Promise.all over a map, it waits for every running worker to settle before
 * reporting the first failure, so nothing is still writing to shared state
 * after the caller has moved on (important when a run is cancelled).
 */
export async function runPool<T>(
  items: T[],
  concurrency: number,
  worker: (item: T) => Promise<void>
): Promise<void> {
  let next = 0;
  let failure: unknown;
  let failed = false;

  const lane = async (): Promise<void> => {
    while (!failed && next < items.length) {
      const item = items[next++];
      try {
        await worker(item);
      } catch (err) {
        if (!failed) {
          failed = true;
          failure = err;
        }
      }
    }
  };

  const lanes: Promise<void>[] = [];
  for (let i = 0; i < Math.min(concurrency, items.length); i++) {
    lanes.push(lane());
  }
  await Promise.all(lanes);
  if (failed) {
    throw failure;
  }
}
