import assert from 'node:assert/strict';
import { test } from 'node:test';
import { comparable, runTool, snapshotOf, startTenant } from './helpers.mjs';

// Random tenants: people, nested groups, sites with SharePoint groups, subsites, libraries with
// their own permissions. The tool's report must equal what the model says, for every person and
// group, in every mode. The model (test/mock-tenant.mjs) shares no code with the tool.

function rng(seed) {
  let s = seed >>> 0;
  return () => {
    s = (Math.imul(s ^ (s >>> 15), 2246822507) ^ Math.imul(s ^ (s >>> 13), 3266489909)) >>> 0;
    s = (s + 0x6d2b79f5) >>> 0;
    let t = Math.imul(s ^ (s >>> 15), 1 | s);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const ROLE_NAMES = ['Full Control', 'Design', 'Edit', 'Contribute', 'Read', 'View Only', 'Limited Access', 'Drop box'];
const hex = (n, len) => n.toString(16).padStart(len, '0');

function randomTenant(seed) {
  const r = rng(seed);
  const pick = (xs) => xs[Math.floor(r() * xs.length)];
  const chance = (p) => r() < p;
  const users = Array.from({ length: 6 }, (_, i) => ({ id: `00000000-0000-0000-0000-${hex(i + 1, 12)}`, upn: `u${i}@contoso.com`, displayName: `User ${i}`, userType: i === 5 ? 'Guest' : 'Member' }));
  const groups = Array.from({ length: 5 }, (_, i) => ({ id: `10000000-0000-0000-0000-${hex(i + 1, 12)}`, displayName: `Group ${i}`, unified: i >= 3, members: [], owners: [] }));
  groups.forEach((g, i) => {
    users.forEach((u) => chance(0.3) && g.members.push(u.id));
    // Security groups nest (only into later ones, so there are no cycles to worry the model).
    if (!g.unified && i > 0 && chance(0.5)) groups[i - 1].unified || g.members.push(groups[i - 1].id);
    if (g.unified) users.forEach((u) => chance(0.2) && g.owners.push(u.id));
  });
  const refs = () => {
    const kind = r();
    if (kind < 0.4) return 'user:' + pick(users).upn;
    if (kind < 0.75) return 'group:' + pick(groups).id;
    if (kind < 0.85) return 'owners:' + pick(groups.filter((g) => g.unified)).id;
    if (kind < 0.93) return 'eeeu';
    return 'everyone';
  };
  let listNo = 0;
  const sites = Array.from({ length: 5 }, (_, s) => {
    const path = `/sites/s${s}`;
    const spGroups = Array.from({ length: 3 }, (_, g) => ({ title: `S${s} group ${g}`, members: Array.from({ length: Math.floor(r() * 3) }, refs).filter((x, i, a) => a.indexOf(x) === i) }));
    const assignments = () =>
      Array.from({ length: 1 + Math.floor(r() * 4) }, () => ({ to: chance(0.5) ? 'sp:' + pick(spGroups).title : refs(), roles: [pick(ROLE_NAMES)].concat(chance(0.2) ? [pick(ROLE_NAMES)] : []) }))
        // One assignment per principal, as in SharePoint.
        .filter((a, i, all) => all.findIndex((b) => b.to === a.to) === i);
    const lists = (webPath) =>
      Array.from({ length: Math.floor(r() * 4) }, () => {
        const n = ++listNo;
        const l = { id: `aaaaaaaa-0000-0000-0000-${hex(n, 12)}`, title: `Lib ${n}`, path: `${webPath}/Lib ${n}`, baseType: chance(0.2) ? 0 : 1 };
        if (chance(0.5)) l.assignments = assignments();
        if (chance(0.1)) l.hidden = true;
        if (chance(0.1)) l.system = true;
        return l;
      });
    const web = (webPath, depth, isRoot) => {
      const w = { path: webPath, title: `Web ${webPath}`, lists: lists(webPath) };
      if (isRoot || chance(0.5)) w.assignments = assignments();
      if (depth < 2) w.webs = Array.from({ length: Math.floor(r() * 3) }, (_, i) => web(`${webPath}/w${i}`, depth + 1, false));
      return w;
    };
    return { path, title: `Site ${s}`, admins: chance(0.3) ? [refs()] : [], spGroups, web: web(path, 0, true), visitors: chance(0.5) ? ['user:' + pick(users).upn] : [] };
  });
  return { users, groups, sites };
}

for (let seed = 1; seed <= 25; seed++) {
  test(`random tenant ${seed}: report equals the permission model`, async (t) => {
    const model = randomTenant(seed * 7919);
    const tenant = await startTenant(model);
    t.after(() => tenant.stop());
    const principals = model.users.map((u) => 'user:' + u.upn).concat(model.groups.map((g) => 'group:' + g.id));
    const keys = model.users.map((u) => 'user:' + u.id).concat(model.groups.map((g) => 'group:' + g.id));
    const keyOf = new Map(principals.map((p, i) => [p, keys[i]]));
    // Three modes per tenant, rotating through the options.
    const modes = [{ scan: {} }, { scan: { thorough: false, includeLists: true } }, { scan: { expandGroups: false } }];
    for (const mode of modes) {
      const scope = { includeLists: !!mode.scan.includeLists };
      const usersOnly = mode.scan.expandGroups === false;
      const who = usersOnly ? principals.slice(0, model.users.length) : principals;
      const { result, error } = await runTool(tenant, { file: { principals: who, ...mode, }, opts: { period: '2026-11', force: true } });
      assert.ifError(error);
      assert.equal(result.sitesFailed, 0);
      const expected = tenant
        .oracle(who, scope)
        .map((r) => `${keyOf.get(r.who)} | ${r.object} | ${r.url} | ${r.access}`)
        .sort();
      const rows = snapshotOf(tenant, '2026-11').rows;
      assert.deepEqual(comparable(rows), expected, `seed ${seed} mode ${JSON.stringify(mode)}`);
      assert.ok(rows.every((r) => r.access !== 'No access' && r.how), 'every row says how');
      assert.equal(rows.filter((r) => r.note).length, 0, 'nothing unconfirmed when SharePoint and the directory agree');
      assert.equal(new Set(rows.map((r) => `${r.p}|${r.object}|${r.url}`)).size, rows.length, 'no duplicate rows');
    }
  });
}
