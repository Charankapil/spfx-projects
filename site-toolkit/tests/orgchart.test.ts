import assert from 'node:assert/strict';
import { test } from 'node:test';

import { buildBatchGetBody, parseBatchResponse } from '../src/common/batch';
import { initials, nearestManagers, parseProfile, ProfileStore, profileUrl } from '../src/webparts/orgChart/profiles';
import { fakeSp, response, useFakeClock } from './helpers';

const ACCOUNT = (name: string): string => `i:0#.f|membership|${name}@contoso.com`;

function raw(name: string, extra: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    AccountName: ACCOUNT(name),
    DisplayName: name.charAt(0).toUpperCase() + name.substring(1),
    Email: `${name}@contoso.com`,
    Title: 'Engineer',
    DirectReports: [],
    ExtendedManagers: [],
    Peers: [],
    UserProfileProperties: [{ Key: 'Department', Value: 'R&D' }, { Key: 'WorkPhone', Value: null }],
    ...extra
  };
}

function batchAnswer(parts: Array<{ status: number; body?: unknown }>): string {
  return (
    parts
      .map(
        (p) =>
          `--batchresponse_9\r\nContent-Type: application/http\r\nContent-Transfer-Encoding: binary\r\n\r\n` +
          `HTTP/1.1 ${p.status} X\r\nCONTENT-TYPE: application/json;odata=nometadata\r\n\r\n${p.body === undefined ? '' : JSON.stringify(p.body)}\r\n`
      )
      .join('') + '--batchresponse_9--\r\n'
  );
}

test('a profile is read from both answer shapes', () => {
  const modern = parseProfile(raw('alex', { ExtendedManagers: [ACCOUNT('ceo'), ACCOUNT('priya')], DirectReports: [ACCOUNT('tom')] }));
  assert.equal(modern?.name, 'Alex');
  assert.equal(modern?.department, 'R&D');
  assert.equal(modern?.phone, '');
  assert.deepEqual(modern?.managers, [ACCOUNT('ceo'), ACCOUNT('priya')]);

  const verbose = parseProfile({
    d: {
      AccountName: ACCOUNT('tom'),
      DisplayName: 'Tom',
      Email: null,
      Title: null,
      DirectReports: { results: [] },
      ExtendedManagers: { results: [ACCOUNT('alex')] },
      Peers: { results: [] },
      UserProfileProperties: { results: [{ Key: 'WorkEmail', Value: 'tom@contoso.com' }, { Key: 'SPS-JobTitle', Value: 'Analyst' }] }
    }
  });
  assert.equal(verbose?.email, 'tom@contoso.com');
  assert.equal(verbose?.jobTitle, 'Analyst');
  assert.deepEqual(verbose?.managers, [ACCOUNT('alex')]);
});

test('an account with no profile is recognised, not shown as an empty person', () => {
  assert.equal(parseProfile({ AccountName: null, DisplayName: null }), undefined);
  assert.equal(parseProfile({}), undefined);
  assert.equal(parseProfile(undefined), undefined);
});

test('the nearest managers are taken from the bottom of the chain', () => {
  const parsed = parseProfile(raw('alex', { ExtendedManagers: ['a', 'b', 'c', 'd'] }));
  if (!parsed) {
    throw new Error('profile expected');
  }
  const p = parsed;
  assert.deepEqual(nearestManagers(p, 2), ['c', 'd']);
  assert.deepEqual(nearestManagers(p, 10), ['a', 'b', 'c', 'd']);
  assert.deepEqual(nearestManagers(p, 0), []);
});

test('initials', () => {
  assert.equal(initials('Priya Shah'), 'PS');
  assert.equal(initials('Ada'), 'A');
  assert.equal(initials('Jean-Luc de la Cruz (Contractor)'), 'JC');
  assert.equal(initials('Émile Zola'), 'ÉZ');
  assert.equal(initials('   '), '?');
});

test("an account name with an apostrophe is quoted safely", () => {
  const clock = useFakeClock();
  const { sp } = fakeSp(() => response(200), clock);
  const url = profileUrl(sp, "i:0#.f|membership|o'brien@contoso.com");
  assert.ok(url.endsWith("?@v='i%3A0%23.f%7Cmembership%7Co''brien%40contoso.com'"), url);
});

test('a batch of reads is built without change sets and its bodies are read back', () => {
  const body = buildBatchGetBody(['https://x/a', 'https://x/b'], 'id1');
  assert.equal(body.split('GET https://x/').length - 1, 2);
  assert.ok(body.indexOf('changeset') < 0);
  const results = parseBatchResponse(batchAnswer([{ status: 200, body: { AccountName: 'a' } }, { status: 404, body: { error: { message: 'gone' } } }]));
  assert.deepEqual(JSON.parse(results[0].body || ''), { AccountName: 'a' });
  assert.equal(results[1].ok, false);
  assert.equal(results[1].error, 'gone');
});

test('many people are read in batches of 50, and remembered', async () => {
  const clock = useFakeClock();
  sessionStorage.clear();
  const accounts = Array.from({ length: 120 }, (_, i) => ACCOUNT(`p${i}`));
  const { sp, calls } = fakeSp((call) => {
    const urls = (call.body || '').split('\r\n').filter((l) => l.indexOf('GET ') === 0);
    return response(
      200,
      batchAnswer(
        urls.map((line) => {
          const account = decodeURIComponent(/@v='([^ ]+)'/.exec(line)?.[1] || '');
          // One person has left: the service answers with an empty profile.
          return account.indexOf('p7@') >= 0 ? { status: 200, body: { AccountName: null } } : { status: 200, body: raw(account.split('|')[2].split('@')[0]) };
        })
      )
    );
  }, clock);
  const store = new ProfileStore(sp);
  const people = await store.many(accounts.concat([accounts[0].toUpperCase()]));
  assert.equal(calls.length, 3, '120 people should take 3 batches');
  assert.ok(calls.every((c) => c.url.endsWith('/_api/$batch')));
  assert.equal(Object.keys(people).length, 119, 'the person without a profile is left out');
  assert.equal(people[ACCOUNT('p8')].name, 'P8');

  await store.many(accounts.slice(0, 60));
  assert.equal(calls.length, 3, 'cached people are not asked for again');
});

test('a throttled read inside a batch makes the whole batch wait and repeat', async () => {
  const clock = useFakeClock();
  sessionStorage.clear();
  let attempt = 0;
  const { sp, calls } = fakeSp(() => {
    attempt++;
    return response(
      200,
      attempt === 1
        ? batchAnswer([{ status: 200, body: raw('a') }]).replace('HTTP/1.1 200 X', 'HTTP/1.1 429 X\r\nRetry-After: 4')
        : batchAnswer([{ status: 200, body: raw('a') }])
    );
  }, clock);
  const people = await new ProfileStore(sp).many([ACCOUNT('a')]);
  assert.equal(calls.length, 2);
  assert.ok(calls[1].at - calls[0].at >= 4000);
  assert.equal(people[ACCOUNT('a')].name, 'A');
});
