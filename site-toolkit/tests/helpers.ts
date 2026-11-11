import { resetQueueForTests, setClockForTests } from '../src/common/RequestQueue';
import { SpClient } from '../src/common/SpClient';

export interface IFakeResponse {
  ok: boolean;
  status: number;
  statusText: string;
  headers: { get(name: string): string | null };
  text(): Promise<string>;
  json(): Promise<unknown>;
}

export function response(status: number, body?: unknown, headers: Record<string, string> = {}): IFakeResponse {
  const text = body === undefined ? '' : typeof body === 'string' ? body : JSON.stringify(body);
  return {
    ok: status >= 200 && status < 300,
    status,
    statusText: `status ${status}`,
    headers: {
      get: (name: string): string | null => {
        for (const key of Object.keys(headers)) {
          if (key.toLowerCase() === name.toLowerCase()) {
            return headers[key];
          }
        }
        return null;
      }
    },
    text: () => Promise.resolve(text),
    json: () => Promise.resolve(JSON.parse(text))
  };
}

export interface ICall {
  method: 'GET' | 'POST';
  url: string;
  config: { name?: string };
  headers: Record<string, string>;
  body: string | undefined;
  at: number;
}

export type Handler = (call: ICall) => IFakeResponse | Promise<IFakeResponse>;

/** Time that only moves when something sleeps, so tests run instantly and exactly. */
export class FakeClock {
  public time = 1000000;
  public now = (): number => this.time;
  public sleep = (ms: number): Promise<void> => {
    this.time += ms;
    return Promise.resolve();
  };
}

export function useFakeClock(): FakeClock {
  const clock = new FakeClock();
  resetQueueForTests();
  setClockForTests(clock);
  return clock;
}

export function fakeSp(handler: Handler, clock: FakeClock): { sp: SpClient; calls: ICall[] } {
  const calls: ICall[] = [];
  const record = (method: 'GET' | 'POST', url: string, config: { name?: string }, options?: { headers?: Record<string, string>; body?: string }): Promise<IFakeResponse> => {
    const call: ICall = { method, url, config, headers: (options && options.headers) || {}, body: options && options.body, at: clock.now() };
    calls.push(call);
    return Promise.resolve(handler(call));
  };
  const http = {
    get: (url: string, config: { name?: string }, options?: { headers?: Record<string, string> }) => record('GET', url, config, options),
    post: (url: string, config: { name?: string }, options?: { headers?: Record<string, string>; body?: string }) => record('POST', url, config, options)
  };
  const sp = new SpClient(http as never, {
    web: { absoluteUrl: 'https://contoso.sharepoint.com/sites/team', serverRelativeUrl: '/sites/team' },
    site: { absoluteUrl: 'https://contoso.sharepoint.com/sites/team', serverRelativeUrl: '/sites/team' }
  });
  return { sp, calls };
}
