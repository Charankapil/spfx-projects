// A small, deliberately strict simulator of the Power Automate / Logic Apps workflow
// definition language - enough to run flow/definition.ts for real in the tests.
//
// It is stricter than the service wherever the service's behaviour is not certain, so a flow
// that passes here does not depend on the lenient reading:
//   - if / and / or / coalesce receive already-evaluated arguments (no laziness to rely on);
//   - int() and integer arithmetic throw outside 32 bits;
//   - concat / replace / split / startsWith ... throw on null or non-string input;
//   - substring throws when the range leaves the string;
//   - indexing a string or a number throws, even with "?";
//   - a function that is not on the list below throws.
// startsWith / endsWith ignore case and contains / equals / replace do not, as documented.

const INT_MAX = 2147483647;

class WdlError extends Error {}

function fail(message) {
  throw new WdlError(message);
}

// ---- expression parser --------------------------------------------------------------

function parse(text) {
  let i = 0;
  const ws = () => {
    while (i < text.length && /\s/.test(text[i])) i++;
  };
  const expr = () => {
    ws();
    let node;
    const c = text[i];
    if (c === "'") {
      let s = '';
      i++;
      for (;;) {
        if (i >= text.length) fail(`Unterminated string in: ${text}`);
        if (text[i] === "'") {
          if (text[i + 1] === "'") {
            s += "'";
            i += 2;
            continue;
          }
          i++;
          break;
        }
        s += text[i++];
      }
      node = { t: 'lit', v: s };
    } else if (/[-0-9]/.test(c)) {
      const m = /^-?\d+(\.\d+)?/.exec(text.slice(i));
      if (!m) fail(`Bad number at ${i} in: ${text}`);
      i += m[0].length;
      node = { t: 'lit', v: Number(m[0]) };
    } else if (/[A-Za-z_]/.test(c)) {
      const m = /^[A-Za-z_][A-Za-z0-9_]*/.exec(text.slice(i));
      i += m[0].length;
      ws();
      if (text[i] === '(') {
        i++;
        const args = [];
        ws();
        if (text[i] === ')') {
          i++;
        } else {
          for (;;) {
            args.push(expr());
            ws();
            if (text[i] === ',') {
              i++;
              continue;
            }
            if (text[i] === ')') {
              i++;
              break;
            }
            fail(`Expected , or ) at ${i} in: ${text}`);
          }
        }
        node = { t: 'call', name: m[0], args };
      } else if (m[0] === 'true' || m[0] === 'false' || m[0] === 'null') {
        node = { t: 'lit', v: m[0] === 'null' ? null : m[0] === 'true' };
      } else {
        fail(`Unexpected identifier ${m[0]} in: ${text}`);
      }
    } else {
      fail(`Unexpected "${c}" at ${i} in: ${text}`);
    }
    for (;;) {
      ws();
      if (text[i] === '?' && text[i + 1] === '[') {
        i += 2;
        const key = expr();
        ws();
        if (text[i] !== ']') fail(`Expected ] at ${i} in: ${text}`);
        i++;
        node = { t: 'index', of: node, key, safe: true };
      } else if (text[i] === '[') {
        i++;
        const key = expr();
        ws();
        if (text[i] !== ']') fail(`Expected ] at ${i} in: ${text}`);
        i++;
        node = { t: 'index', of: node, key, safe: false };
      } else {
        break;
      }
    }
    return node;
  };
  const node = expr();
  ws();
  if (i !== text.length) fail(`Unexpected trailing text at ${i} in: ${text}`);
  return node;
}

const parsed = new Map();
function parseCached(text) {
  let n = parsed.get(text);
  if (!n) {
    n = parse(text);
    parsed.set(text, n);
  }
  return n;
}

// ---- values -------------------------------------------------------------------------

const isObject = (v) => v !== null && typeof v === 'object' && !Array.isArray(v) && !(v instanceof Xml);
const str = (v, fn) => (typeof v === 'string' ? v : fail(`${fn}: expected a string, got ${describe(v)}`));
const arr = (v, fn) => (Array.isArray(v) ? v : fail(`${fn}: expected an array, got ${describe(v)}`));
const bool = (v, fn) => (typeof v === 'boolean' ? v : fail(`${fn}: expected a boolean, got ${describe(v)}`));
const int32 = (v, fn) => {
  if (typeof v !== 'number' || !Number.isInteger(v)) fail(`${fn}: expected an integer, got ${describe(v)}`);
  if (Math.abs(v) > INT_MAX) fail(`${fn}: ${v} does not fit a 32-bit integer`);
  return v;
};
function describe(v) {
  if (v === null || v === undefined) return 'null';
  if (Array.isArray(v)) return 'an array';
  if (v instanceof Xml) return 'xml';
  return `${typeof v} ${JSON.stringify(v).slice(0, 60)}`;
}
function deepEqual(a, b) {
  if (a === b) return true;
  if (a === null || b === null || typeof a !== typeof b) return false;
  if (typeof a !== 'object') return false;
  return JSON.stringify(a) === JSON.stringify(b);
}
function toText(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'boolean') return v ? 'True' : 'False';
  if (typeof v === 'number') return String(v);
  if (v instanceof Xml) fail('string: xml is not supported here');
  return JSON.stringify(v);
}

class Xml {
  constructor(value) {
    this.value = value;
  }
}

// Uri.EscapeDataString: everything but the unreserved characters.
const escapeData = (s) => encodeURIComponent(s).replace(/[!'()*]/g, (c) => '%' + c.charCodeAt(0).toString(16).toUpperCase());

function formatDate(iso, format) {
  const d = new Date(iso);
  if (isNaN(d.getTime())) fail(`formatDateTime: bad timestamp ${iso}`);
  const p = (n, w = 2) => String(n).padStart(w, '0');
  const known = { yyyy: p(d.getUTCFullYear(), 4), MM: p(d.getUTCMonth() + 1), dd: p(d.getUTCDate()), HH: p(d.getUTCHours()), mm: p(d.getUTCMinutes()), ss: p(d.getUTCSeconds()) };
  const out = format.replace(/yyyy|MM|dd|HH|mm|ss/g, (t) => known[t]);
  if (/[A-Za-z]/.test(out.replace(/[0-9]/g, '').replace(/UTC/g, ''))) fail(`formatDateTime: unsupported format ${format}`);
  return out;
}

const FUNCTIONS = {
  concat: (...a) => a.map((x) => (x === null ? '' : str(x, 'concat'))).join(''),
  replace: (s, a, b) => {
    if (str(a, 'replace') === '') fail('replace: the old value is empty');
    return str(s, 'replace').split(a).join(str(b, 'replace'));
  },
  split: (s, sep) => {
    if (str(sep, 'split') === '') fail('split: empty separator');
    return str(s, 'split').split(sep);
  },
  join: (a, sep) => arr(a, 'join').map(toText).join(str(sep, 'join')),
  first: (v) => (typeof v === 'string' ? v.charAt(0) : arr(v, 'first').length ? v[0] : null),
  last: (v) => (typeof v === 'string' ? v.charAt(v.length - 1) : arr(v, 'last').length ? v[v.length - 1] : null),
  length: (v) => (typeof v === 'string' || Array.isArray(v) ? v.length : fail(`length: expected a string or an array, got ${describe(v)}`)),
  empty: (v) => v === null || v === '' || (Array.isArray(v) && v.length === 0) || (isObject(v) && Object.keys(v).length === 0),
  if: (c, a, b) => (bool(c, 'if') ? a : b),
  equals: (a, b) => deepEqual(a, b),
  not: (a) => !bool(a, 'not'),
  and: (...a) => (a.length < 2 ? fail('and: needs two or more arguments') : a.every((x) => bool(x, 'and'))),
  or: (...a) => (a.length < 2 ? fail('or: needs two or more arguments') : a.some((x) => bool(x, 'or'))),
  less: (a, b) => (typeof a === typeof b && (typeof a === 'number' || typeof a === 'string') ? a < b : fail('less: mismatched types')),
  greater: (a, b) => (typeof a === typeof b && (typeof a === 'number' || typeof a === 'string') ? a > b : fail('greater: mismatched types')),
  contains: (c, v) => {
    if (typeof c === 'string') return c.indexOf(str(v, 'contains')) >= 0;
    if (Array.isArray(c)) return c.some((x) => deepEqual(x, v));
    if (isObject(c)) return Object.prototype.hasOwnProperty.call(c, str(v, 'contains'));
    return fail(`contains: expected a string, array or object, got ${describe(c)}`);
  },
  startsWith: (s, p) => str(s, 'startsWith').toLowerCase().startsWith(str(p, 'startsWith').toLowerCase()),
  endsWith: (s, p) => str(s, 'endsWith').toLowerCase().endsWith(str(p, 'endsWith').toLowerCase()),
  toLower: (s) => str(s, 'toLower').toLowerCase(),
  toUpper: (s) => str(s, 'toUpper').toUpperCase(),
  trim: (s) => str(s, 'trim').trim(),
  int: (s) => {
    if (typeof s === 'number') return int32(s, 'int');
    if (!/^-?\d+$/.test(str(s, 'int'))) fail(`int: "${s}" is not an integer`);
    return int32(Number(s), 'int');
  },
  string: (v) => toText(v),
  add: (a, b) => int32(int32(a, 'add') + int32(b, 'add'), 'add'),
  sub: (a, b) => int32(int32(a, 'sub') - int32(b, 'sub'), 'sub'),
  mul: (a, b) => int32(int32(a, 'mul') * int32(b, 'mul'), 'mul'),
  div: (a, b) => (int32(b, 'div') === 0 ? fail('div: division by zero') : Math.trunc(int32(a, 'div') / b)),
  mod: (a, b) => (int32(b, 'mod') === 0 ? fail('mod: division by zero') : int32(a, 'mod') % b),
  min: (a, b) => Math.min(int32(a, 'min'), int32(b, 'min')),
  substring: (s, start, len) => {
    str(s, 'substring');
    int32(start, 'substring');
    if (start < 0 || start > s.length) fail(`substring: start ${start} is outside a string of ${s.length}`);
    if (len === undefined) return s.substring(start);
    int32(len, 'substring');
    if (len < 0 || start + len > s.length) fail(`substring: ${start}+${len} is outside a string of ${s.length}`);
    return s.substring(start, start + len);
  },
  coalesce: (...a) => {
    for (const x of a) if (x !== null) return x;
    return null;
  },
  json: (s) => {
    try {
      return JSON.parse(str(s, 'json'));
    } catch (e) {
      return fail(`json: ${e.message} in ${String(s).slice(0, 80)}`);
    }
  },
  union: (...a) => {
    const out = [];
    a.forEach((x) => arr(x, 'union').forEach((v) => (out.some((o) => deepEqual(o, v)) ? 0 : out.push(v))));
    return out;
  },
  encodeUriComponent: (s) => escapeData(str(s, 'encodeUriComponent')),
  decodeUriComponent: (s) => decodeURIComponent(str(s, 'decodeUriComponent')),
  uriHost: (s) => new URL(str(s, 'uriHost')).host,
  uriPath: (s) => new URL(str(s, 'uriPath')).pathname,
  formatDateTime: (ts, format) => formatDate(str(ts, 'formatDateTime'), str(format, 'formatDateTime')),
  xml: (v) => {
    if (!isObject(v) || Object.keys(v).length !== 1) fail('xml: the JSON needs exactly one root property');
    return new Xml(v);
  },
  xpath: (x, path) => {
    if (!(x instanceof Xml)) fail('xpath: the first argument is not xml');
    // Only the one shape the flow uses: the Value of the first <c> whose Key is ..., as a string.
    const m = /^string\(\/(\w+)\/(\w+)\[(\w+)="([^"]*)"\]\/(\w+)\)$/.exec(str(path, 'xpath'));
    if (!m) fail(`xpath: unsupported expression ${path}`);
    const root = x.value[m[1]];
    if (root === undefined) return '';
    let nodes = isObject(root) ? root[m[2]] : undefined;
    if (nodes === undefined || nodes === null) return '';
    if (!Array.isArray(nodes)) nodes = [nodes];
    const hit = nodes.find((n) => isObject(n) && toText(n[m[3]]) === m[4]);
    return hit ? toText(hit[m[5]] === undefined ? null : hit[m[5]]) : '';
  }
};

// ---- evaluation ---------------------------------------------------------------------

function evaluate(node, ctx) {
  if (node.t === 'lit') return node.v;
  if (node.t === 'index') {
    const target = evaluate(node.of, ctx);
    const key = evaluate(node.key, ctx);
    if (target === null || target === undefined) {
      if (node.safe) return null;
      fail('Cannot index null (use ?[...])');
    }
    if (Array.isArray(target)) {
      if (!Number.isInteger(key)) fail(`Array index must be an integer, got ${describe(key)}`);
      if (key < 0 || key >= target.length) fail(`Array index ${key} is out of range (length ${target.length})`);
      return target[key];
    }
    if (isObject(target)) {
      const v = target[str(key, 'property access')];
      if (v === undefined) {
        if (node.safe) return null;
        fail(`Property ${key} does not exist (use ?[...])`);
      }
      return v;
    }
    return fail(`Cannot index ${describe(target)}`);
  }
  const args = node.args.map((a) => evaluate(a, ctx));
  switch (node.name) {
    case 'variables': {
      const v = ctx.run.variables[str(args[0], 'variables')];
      if (!v) fail(`Variable ${args[0]} is not initialized`);
      return v.value;
    }
    case 'outputs':
    case 'body': {
      const name = str(args[0], node.name);
      if (!ctx.run.known.has(name)) fail(`${node.name}('${name}'): no such action`);
      const a = ctx.run.results[name];
      const out = a && a.outputs !== undefined ? a.outputs : null;
      if (node.name === 'outputs') return out;
      return isObject(out) && out.body !== undefined ? out.body : null;
    }
    case 'item':
      if (ctx.item === undefined) fail('item() used outside a Select / Filter array');
      return ctx.item;
    case 'items': {
      const name = str(args[0], 'items');
      if (!(name in ctx.loops)) fail(`items('${name}'): not inside that loop`);
      return ctx.loops[name];
    }
    case 'parameters':
      return ctx.run.parameters[str(args[0], 'parameters')] === undefined ? fail(`Unknown parameter ${args[0]}`) : ctx.run.parameters[args[0]];
    case 'utcNow':
      return ctx.run.now();
    default: {
      const fn = FUNCTIONS[node.name];
      if (!fn) fail(`Function ${node.name} is not on the vetted list`);
      const v = fn(...args);
      return v === undefined ? null : v;
    }
  }
}

function evalExpression(text, ctx) {
  ctx.run.note(text);
  return evaluate(parseCached(text), ctx);
}

/** The @{...} segments of a string that is not one whole expression. */
function interpolate(text, ctx) {
  let out = '';
  let i = 0;
  while (i < text.length) {
    if (text[i] === '@' && text[i + 1] === '{') {
      let j = i + 2;
      let inString = false;
      for (; j < text.length; j++) {
        if (text[j] === "'") inString = !inString;
        else if (text[j] === '}' && !inString) break;
      }
      if (j >= text.length) fail(`Unterminated @{ in: ${text}`);
      out += toText(evalExpression(text.slice(i + 2, j), ctx));
      i = j + 1;
    } else if (text[i] === '@' && text[i + 1] === '@') {
      out += '@';
      i += 2;
    } else {
      out += text[i++];
    }
  }
  return out;
}

function evalTemplate(value, ctx) {
  if (typeof value === 'string') {
    if (value.startsWith('@@')) return interpolate(value, ctx);
    if (value.startsWith('@{')) return interpolate(value, ctx);
    if (value.startsWith('@')) return evalExpression(value.slice(1), ctx);
    return interpolate(value, ctx);
  }
  if (Array.isArray(value)) return value.map((v) => evalTemplate(v, ctx));
  if (value !== null && typeof value === 'object') {
    const out = {};
    Object.keys(value).forEach((k) => {
      out[k] = evalTemplate(value[k], ctx);
    });
    return out;
  }
  return value;
}

function evalCondition(c, ctx) {
  if (typeof c === 'string') return bool(evalTemplate(c, ctx), 'condition');
  const op = Object.keys(c)[0];
  const v = c[op];
  switch (op) {
    case 'and':
      return v.every((x) => evalCondition(x, ctx));
    case 'or':
      return v.some((x) => evalCondition(x, ctx));
    case 'not':
      return !evalCondition(v, ctx);
    case 'equals':
    case 'greater':
    case 'less':
      return FUNCTIONS[op](evalTemplate(v[0], ctx), evalTemplate(v[1], ctx));
    default:
      return fail(`Unsupported condition operator ${op}`);
  }
}

// ---- actions ------------------------------------------------------------------------

const DONE_BADLY = ['Failed', 'TimedOut'];

function namesIn(actions, out = []) {
  Object.keys(actions || {}).forEach((name) => {
    out.push(name);
    const a = actions[name];
    namesIn(a.actions, out);
    if (a.else) namesIn(a.else.actions, out);
  });
  return out;
}

/** Runs one scope; returns Succeeded or Failed by the "last action of every branch" rule. */
async function runScope(actions, ctx) {
  const run = ctx.run;
  const names = Object.keys(actions || {});
  namesIn(actions).forEach((n) => {
    run.results[n] = { status: 'Skipped', outputs: null };
  });
  const done = {};
  let left = names.slice();
  while (left.length) {
    const ready = left.filter((n) => Object.keys(actions[n].runAfter || {}).every((p) => p in done));
    if (!ready.length) fail(`Circular or dangling runAfter among: ${left.join(', ')}`);
    for (const name of ready) {
      const a = actions[name];
      const after = a.runAfter || {};
      Object.keys(after).forEach((p) => {
        if (names.indexOf(p) < 0) fail(`${name} runs after ${p}, which is not in the same scope`);
      });
      const go = Object.keys(after).every((p) => after[p].indexOf(done[p]) >= 0);
      if (!go) {
        done[name] = 'Skipped';
        run.results[name] = { status: 'Skipped', outputs: null };
        continue;
      }
      run.executed++;
      run.counts[name] = (run.counts[name] || 0) + 1;
      let status = 'Succeeded';
      let outputs = null;
      try {
        const r = await runAction(name, a, ctx);
        status = r.status;
        outputs = r.outputs === undefined ? null : r.outputs;
        if (r.error) run.errors.push(`${name}: ${r.error}`);
      } catch (e) {
        if (!(e instanceof WdlError)) throw e;
        status = 'Failed';
        run.errors.push(`${name}: ${e.message}`);
      }
      done[name] = status;
      run.results[name] = { status, outputs };
    }
    left = left.filter((n) => !(n in done));
  }
  // A skipped last action takes the status of what it would have run after.
  const effective = (n) => (done[n] !== 'Skipped' ? done[n] : Object.keys(actions[n].runAfter || {}).some((p) => DONE_BADLY.indexOf(effective(p)) >= 0) ? 'Failed' : 'Skipped');
  const leaves = names.filter((n) => !names.some((m) => Object.keys(actions[m].runAfter || {}).indexOf(n) >= 0));
  return leaves.some((n) => DONE_BADLY.indexOf(effective(n)) >= 0) ? 'Failed' : 'Succeeded';
}

async function runAction(name, a, ctx) {
  const run = ctx.run;
  switch (a.type) {
    case 'Compose':
      return { status: 'Succeeded', outputs: evalTemplate(a.inputs, ctx) };
    case 'InitializeVariable': {
      if (ctx.depth > 0) fail('Variables can only be initialized at the top level');
      a.inputs.variables.forEach((v) => {
        if (run.variables[v.name]) fail(`Variable ${v.name} is initialized twice`);
        run.variables[v.name] = { type: v.type, value: checkType(v.type, evalTemplate(v.value, ctx), v.name) };
      });
      return { status: 'Succeeded' };
    }
    case 'SetVariable': {
      const v = run.variables[a.inputs.name] || fail(`Variable ${a.inputs.name} is not initialized`);
      if (JSON.stringify(a.inputs.value).indexOf(`variables('${a.inputs.name}')`) >= 0) fail(`Set variable ${a.inputs.name} refers to itself`);
      v.value = checkType(v.type, evalTemplate(a.inputs.value, ctx), a.inputs.name);
      return { status: 'Succeeded' };
    }
    case 'AppendToStringVariable': {
      const v = run.variables[a.inputs.name] || fail(`Variable ${a.inputs.name} is not initialized`);
      if (v.type !== 'string') fail(`${a.inputs.name} is not a string variable`);
      v.value += str(evalTemplate(a.inputs.value, ctx), 'Append to string variable');
      return { status: 'Succeeded' };
    }
    case 'Select': {
      const from = arr(evalTemplate(a.inputs.from, ctx), 'Select from');
      return { status: 'Succeeded', outputs: { body: from.map((item) => evalTemplate(a.inputs.select, { ...ctx, item })) } };
    }
    case 'Query': {
      const from = arr(evalTemplate(a.inputs.from, ctx), 'Filter array from');
      return { status: 'Succeeded', outputs: { body: from.filter((item) => bool(evalTemplate(a.inputs.where, { ...ctx, item }), 'Filter array where')) } };
    }
    case 'Foreach': {
      const items = arr(evalTemplate(a.foreach, ctx), 'Apply to each');
      let status = 'Succeeded';
      for (const item of items) {
        const s = await runScope(a.actions, { ...ctx, depth: ctx.depth + 1, item: undefined, loops: { ...ctx.loops, [name]: item } });
        if (s !== 'Succeeded') status = 'Failed';
      }
      if (!items.length) namesIn(a.actions).forEach((n) => (run.results[n] = { status: 'Skipped', outputs: null }));
      return { status };
    }
    case 'Until': {
      for (let n = 0; ; n++) {
        if (n >= a.limit.count) return { status: 'Failed', error: `Until reached its limit of ${a.limit.count}` };
        const s = await runScope(a.actions, { ...ctx, depth: ctx.depth + 1 });
        const stop = bool(evalTemplate(a.expression, ctx), 'Until');
        if (s !== 'Succeeded') return { status: 'Failed' };
        if (stop) return { status: 'Succeeded' };
      }
    }
    case 'If': {
      const yes = evalCondition(a.expression, ctx);
      namesIn(a.actions).forEach((n) => (run.results[n] = { status: 'Skipped', outputs: null }));
      namesIn(a.else && a.else.actions).forEach((n) => (run.results[n] = { status: 'Skipped', outputs: null }));
      const s = await runScope(yes ? a.actions : (a.else && a.else.actions) || {}, { ...ctx, depth: ctx.depth + 1 });
      return { status: s };
    }
    case 'OpenApiConnection': {
      const inputs = evalTemplate(a.inputs, ctx);
      return run.connector(inputs, name);
    }
    default:
      return fail(`Unsupported action type ${a.type}`);
  }
}

function checkType(type, value, name) {
  const ok = type === 'string' ? typeof value === 'string' : type === 'integer' ? Number.isInteger(value) : type === 'array' ? Array.isArray(value) : type === 'boolean' ? typeof value === 'boolean' : fail(`Unsupported variable type ${type}`);
  if (!ok) fail(`Variable ${name} is ${type} but was given ${describe(value)}`);
  return value;
}

/**
 * Runs a workflow definition. `connector(inputs, actionName)` answers the connector actions
 * with { status, outputs }. Returns the run: status, variables, per-action results (last
 * iteration), every expression error, and how often each action ran.
 */
async function runFlow(definition, { connector, now = () => new Date().toISOString() }) {
  const expressions = new Set();
  const run = {
    variables: {},
    results: {},
    known: new Set(namesIn(definition.actions)),
    parameters: { $authentication: {}, $connections: {} },
    errors: [],
    counts: {},
    executed: 0,
    connector,
    now,
    note: (t) => expressions.add(t)
  };
  const status = await runScope(definition.actions, { run, depth: 0, loops: {} });
  return { status, variables: run.variables, results: run.results, errors: run.errors, counts: run.counts, executed: run.executed, expressions };
}

/** Evaluates one expression on its own (for testing a building block). */
function evalStandalone(expression, item) {
  const run = { variables: {}, results: {}, known: new Set(), parameters: {}, now: () => new Date().toISOString(), note: () => undefined };
  return evaluate(parse(expression), { run, loops: {}, item });
}

module.exports = { runFlow, evalStandalone, parse, namesIn, WdlError };
