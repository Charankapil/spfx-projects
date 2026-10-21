import { ExpressionError, fieldRefs, functionNames, Node, parse } from './parser';

/** What an expression can see while it runs. */
export interface IExpressionEnv {
  /** Current form value of a field, by internal name. Undefined for unknown fields. */
  getField(name: string): unknown;
  /** True when the name is a field on the list (lets the evaluator report typos). */
  hasField?(name: string): boolean;
  user: { name: string; email: string; loginName: string };
  /** SharePoint group names the current user belongs to. */
  groups: string[];
  mode: string;
  itemId?: number;
  now?: () => Date;
}

type Fn = (args: Node[], env: IExpressionEnv) => unknown;

export function isBlank(value: unknown): boolean {
  return (
    value === undefined ||
    value === null ||
    value === '' ||
    (typeof value === 'number' && isNaN(value)) ||
    (Array.isArray(value) && value.length === 0)
  );
}

/**
 * Lookup, person, link and term values are objects; in comparisons and text
 * they stand for their display text, so `[Customer] = "Contoso"` works
 * without `.Title`.
 */
export function toPrimitive(value: unknown): unknown {
  if (value === null || value === undefined) {
    return null;
  }
  if (Array.isArray(value)) {
    return value.map(toPrimitive);
  }
  if (value instanceof Date) {
    return value;
  }
  if (typeof value === 'object') {
    const o = value as { [key: string]: unknown };
    if (o.title !== undefined) {
      return o.title;
    }
    if (o.label !== undefined) {
      return o.label;
    }
    if (o.url !== undefined) {
      return o.url;
    }
    if (o.name !== undefined) {
      return o.name;
    }
  }
  return value;
}

export function toBool(value: unknown): boolean {
  const v = toPrimitive(value);
  if (isBlank(v)) {
    return false;
  }
  if (typeof v === 'string') {
    const lower = v.toLowerCase();
    return lower !== 'false' && lower !== '0' && lower !== 'no';
  }
  return !!v;
}

function toNumber(value: unknown): number | null {
  const v = toPrimitive(value);
  if (isBlank(v)) {
    return null;
  }
  if (typeof v === 'number') {
    return v;
  }
  if (typeof v === 'boolean') {
    return v ? 1 : 0;
  }
  if (v instanceof Date) {
    return v.getTime();
  }
  if (typeof v === 'string') {
    const n = Number(v.replace(/,/g, '').trim());
    return isNaN(n) ? null : n;
  }
  return null;
}

function toDate(value: unknown): Date | null {
  const v = toPrimitive(value);
  if (isBlank(v)) {
    return null;
  }
  if (v instanceof Date) {
    return v;
  }
  if (typeof v === 'number' || typeof v === 'string') {
    const d = new Date(v as string);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

function pad(n: number, width = 2): string {
  let s = String(n);
  while (s.length < width) {
    s = '0' + s;
  }
  return s;
}

/** .NET-style date tokens: yyyy yy MMMM MMM MM M dd d HH H hh h mm ss tt. */
export function formatDate(d: Date, format: string): string {
  return format.replace(/yyyy|yy|MMMM|MMM|MM|M|dd|d|HH|H|hh|h|mm|ss|tt/g, (token) => {
    const h12 = d.getHours() % 12 || 12;
    switch (token) {
      case 'yyyy': return String(d.getFullYear());
      case 'yy': return pad(d.getFullYear() % 100);
      case 'MMMM': return MONTHS[d.getMonth()];
      case 'MMM': return MONTHS[d.getMonth()].substring(0, 3);
      case 'MM': return pad(d.getMonth() + 1);
      case 'M': return String(d.getMonth() + 1);
      case 'dd': return pad(d.getDate());
      case 'd': return String(d.getDate());
      case 'HH': return pad(d.getHours());
      case 'H': return String(d.getHours());
      case 'hh': return pad(h12);
      case 'h': return String(h12);
      case 'mm': return pad(d.getMinutes());
      case 'ss': return pad(d.getSeconds());
      case 'tt': return d.getHours() < 12 ? 'AM' : 'PM';
      default: return token;
    }
  });
}

/** Number formats: "0", "0.00", "#,##0", "#,##0.00", optional trailing "%". */
export function formatNumber(n: number, format: string): string {
  let value = n;
  let suffix = '';
  let fmt = format;
  if (fmt.charAt(fmt.length - 1) === '%') {
    value = value * 100;
    suffix = '%';
    fmt = fmt.substring(0, fmt.length - 1);
  }
  const dot = fmt.indexOf('.');
  const decimals = dot < 0 ? 0 : fmt.length - dot - 1;
  let text = value.toFixed(decimals);
  if (fmt.indexOf(',') >= 0) {
    const parts = text.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    text = parts.join('.');
  }
  return text + suffix;
}

export function toText(value: unknown): string {
  const v = toPrimitive(value);
  if (isBlank(v)) {
    return '';
  }
  if (Array.isArray(v)) {
    return v.map(toText).join('; ');
  }
  if (v instanceof Date) {
    return v.toLocaleDateString();
  }
  return String(v);
}

function startOfDay(d: Date): Date {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/** Equality with blank-awareness and case-insensitive text comparison. */
export function looseEquals(a: unknown, b: unknown): boolean {
  const x = toPrimitive(a);
  const y = toPrimitive(b);
  if (isBlank(x) || isBlank(y)) {
    return isBlank(x) && isBlank(y);
  }
  if (Array.isArray(x) || Array.isArray(y)) {
    const xs = Array.isArray(x) ? x : [x];
    const ys = Array.isArray(y) ? y : [y];
    return xs.length === ys.length && xs.every((item) => ys.some((other) => looseEquals(item, other)));
  }
  if (x instanceof Date || y instanceof Date) {
    const dx = toDate(x);
    const dy = toDate(y);
    return !!dx && !!dy && dx.getTime() === dy.getTime();
  }
  if (typeof x === 'number' || typeof y === 'number' || typeof x === 'boolean' || typeof y === 'boolean') {
    if (typeof x === 'boolean' || typeof y === 'boolean') {
      return toBool(x) === toBool(y);
    }
    return toNumber(x) === toNumber(y);
  }
  return String(x).toLowerCase() === String(y).toLowerCase();
}

function compare(a: unknown, b: unknown): number | null {
  const x = toPrimitive(a);
  const y = toPrimitive(b);
  if (isBlank(x) || isBlank(y)) {
    return null;
  }
  if (x instanceof Date || y instanceof Date) {
    const dx = toDate(x);
    const dy = toDate(y);
    return dx && dy ? dx.getTime() - dy.getTime() : null;
  }
  const nx = toNumber(x);
  const ny = toNumber(y);
  if (nx !== null && ny !== null && (typeof x === 'number' || typeof y === 'number')) {
    return nx - ny;
  }
  const sx = String(x).toLowerCase();
  const sy = String(y).toLowerCase();
  return sx < sy ? -1 : sx > sy ? 1 : 0;
}

function contains(container: unknown, item: unknown): boolean {
  if (isBlank(container)) {
    return false;
  }
  if (Array.isArray(container)) {
    return container.some((c) => looseEquals(c, item));
  }
  const text = toText(container).toLowerCase();
  return text.indexOf(toText(item).toLowerCase()) >= 0;
}

function flatten(values: unknown[]): unknown[] {
  const out: unknown[] = [];
  values.forEach((v) => {
    if (Array.isArray(v)) {
      out.push(...flatten(v));
    } else {
      out.push(v);
    }
  });
  return out;
}

function addToDate(date: Date, amount: number, unit: string): Date {
  const d = new Date(date.getTime());
  switch (unit) {
    case 'year':
    case 'years':
      d.setFullYear(d.getFullYear() + amount);
      break;
    case 'month':
    case 'months':
      d.setMonth(d.getMonth() + amount);
      break;
    case 'week':
    case 'weeks':
      d.setDate(d.getDate() + amount * 7);
      break;
    case 'hour':
    case 'hours':
      d.setTime(d.getTime() + amount * 3600000);
      break;
    case 'minute':
    case 'minutes':
      d.setTime(d.getTime() + amount * 60000);
      break;
    case 'day':
    case 'days':
      d.setDate(d.getDate() + amount);
      break;
    default:
      throw new ExpressionError(`Unknown date unit "${unit}". Use days, weeks, months, years, hours or minutes`);
  }
  return d;
}

function diffDates(a: Date, b: Date, unit: string): number {
  switch (unit) {
    case 'year':
    case 'years':
      return b.getFullYear() - a.getFullYear();
    case 'month':
    case 'months':
      return (b.getFullYear() - a.getFullYear()) * 12 + b.getMonth() - a.getMonth();
    case 'week':
    case 'weeks':
      return Math.round((startOfDay(b).getTime() - startOfDay(a).getTime()) / 86400000 / 7 * 1000) / 1000;
    case 'hour':
    case 'hours':
      return Math.floor((b.getTime() - a.getTime()) / 3600000);
    case 'minute':
    case 'minutes':
      return Math.floor((b.getTime() - a.getTime()) / 60000);
    case 'day':
    case 'days':
      // Calendar days; rounding absorbs daylight-saving hour shifts.
      return Math.round((startOfDay(b).getTime() - startOfDay(a).getTime()) / 86400000);
    default:
      throw new ExpressionError(`Unknown date unit "${unit}". Use days, weeks, months, years, hours or minutes`);
  }
}

function arity(name: string, args: Node[], min: number, max = min): void {
  if (args.length < min || args.length > max) {
    const expected = min === max ? `${min}` : max === Infinity ? `at least ${min}` : `${min}-${max}`;
    throw new ExpressionError(`${name}() takes ${expected} argument${expected === '1' ? '' : 's'}, got ${args.length}`);
  }
}

function evalAll(args: Node[], env: IExpressionEnv): unknown[] {
  return args.map((a) => evaluateNode(a, env));
}

function regex(pattern: string, flags: string): RegExp {
  try {
    return new RegExp(pattern, flags);
  } catch {
    throw new ExpressionError(`"${pattern}" is not a valid regular expression`);
  }
}

/** name (lower case) -> [display name, implementation]. Display names feed the designer's function list. */
const FUNCTIONS: { [name: string]: { display: string; signature: string; help: string; fn: Fn } } = {};

function define(display: string, signature: string, help: string, fn: Fn): void {
  FUNCTIONS[display.toLowerCase()] = { display, signature, help, fn };
}

define('If', 'If(condition, then, [condition2, then2, ...], [else])', 'First result whose condition is true.', (args, env) => {
  arity('If', args, 2, Infinity);
  let i = 0;
  for (; i + 1 < args.length; i += 2) {
    if (toBool(evaluateNode(args[i], env))) {
      return evaluateNode(args[i + 1], env);
    }
  }
  return i < args.length ? evaluateNode(args[i], env) : null;
});
define('Switch', 'Switch(value, match1, result1, ..., [default])', 'Result for the first match.', (args, env) => {
  arity('Switch', args, 3, Infinity);
  const value = evaluateNode(args[0], env);
  let i = 1;
  for (; i + 1 < args.length; i += 2) {
    if (looseEquals(value, evaluateNode(args[i], env))) {
      return evaluateNode(args[i + 1], env);
    }
  }
  return i < args.length ? evaluateNode(args[i], env) : null;
});
define('And', 'And(a, b, ...)', 'True when every argument is true.', (args, env) => args.every((a) => toBool(evaluateNode(a, env))));
define('Or', 'Or(a, b, ...)', 'True when any argument is true.', (args, env) => args.some((a) => toBool(evaluateNode(a, env))));
define('Not', 'Not(value)', 'Opposite of a true / false value.', (args, env) => {
  arity('Not', args, 1);
  return !toBool(evaluateNode(args[0], env));
});
define('IsBlank', 'IsBlank(value)', 'True when the value is empty.', (args, env) => {
  arity('IsBlank', args, 1);
  return isBlank(toPrimitive(evaluateNode(args[0], env)));
});
define('Coalesce', 'Coalesce(a, b, ...)', 'First value that is not blank.', (args, env) => {
  for (const a of args) {
    const v = evaluateNode(a, env);
    if (!isBlank(v)) {
      return v;
    }
  }
  return null;
});
define('IfError', 'IfError(value, fallback)', 'Fallback when the value cannot be calculated.', (args, env) => {
  arity('IfError', args, 2);
  try {
    return evaluateNode(args[0], env);
  } catch {
    return evaluateNode(args[1], env);
  }
});
define('Len', 'Len(text or list)', 'Number of characters, or of items in a multi-value field.', (args, env) => {
  arity('Len', args, 1);
  const v = toPrimitive(evaluateNode(args[0], env));
  return Array.isArray(v) ? v.length : toText(v).length;
});
define('Count', 'Count(list)', 'Number of values in a multi-value field.', (args, env) => {
  arity('Count', args, 1);
  const v = evaluateNode(args[0], env);
  return isBlank(v) ? 0 : Array.isArray(v) ? v.length : 1;
});
define('Lower', 'Lower(text)', 'Lower-case text.', (args, env) => {
  arity('Lower', args, 1);
  return toText(evaluateNode(args[0], env)).toLowerCase();
});
define('Upper', 'Upper(text)', 'Upper-case text.', (args, env) => {
  arity('Upper', args, 1);
  return toText(evaluateNode(args[0], env)).toUpperCase();
});
define('Trim', 'Trim(text)', 'Text without leading or trailing spaces.', (args, env) => {
  arity('Trim', args, 1);
  return toText(evaluateNode(args[0], env)).trim();
});
define('Left', 'Left(text, count)', 'First characters of the text.', (args, env) => {
  arity('Left', args, 2);
  const [t, n] = evalAll(args, env);
  return toText(t).substring(0, toNumber(n) || 0);
});
define('Right', 'Right(text, count)', 'Last characters of the text.', (args, env) => {
  arity('Right', args, 2);
  const [t, n] = evalAll(args, env);
  const text = toText(t);
  return text.substring(Math.max(0, text.length - (toNumber(n) || 0)));
});
define('Mid', 'Mid(text, start, [count])', 'Part of the text; start is 1-based.', (args, env) => {
  arity('Mid', args, 2, 3);
  const [t, s, n] = evalAll(args, env);
  const start = Math.max(0, (toNumber(s) || 1) - 1);
  return args.length === 3 ? toText(t).substr(start, toNumber(n) || 0) : toText(t).substring(start);
});
define('Find', 'Find(find, within)', '1-based position of text, or blank.', (args, env) => {
  arity('Find', args, 2);
  const [f, w] = evalAll(args, env);
  const index = toText(w).indexOf(toText(f));
  return index < 0 ? null : index + 1;
});
define('Replace', 'Replace(text, find, replacement)', 'Replaces every occurrence.', (args, env) => {
  arity('Replace', args, 3);
  const [t, f, r] = evalAll(args, env);
  const find = toText(f);
  return find ? toText(t).split(find).join(toText(r)) : toText(t);
});
define('Concat', 'Concat(a, b, ...)', 'Joins values into one text.', (args, env) => evalAll(args, env).map(toText).join(''));
define('Join', 'Join(list, separator)', 'Joins the values of a multi-value field.', (args, env) => {
  arity('Join', args, 1, 2);
  const [list, sep] = evalAll(args, env);
  const items = toPrimitive(list);
  return (Array.isArray(items) ? items : isBlank(items) ? [] : [items]).map(toText).join(args.length > 1 ? toText(sep) : ', ');
});
define('Split', 'Split(text, separator)', 'Splits text into a list.', (args, env) => {
  arity('Split', args, 2);
  const [t, s] = evalAll(args, env);
  const text = toText(t);
  return text ? text.split(toText(s)).map((x) => x.trim()).filter((x) => x !== '') : [];
});
define('Contains', 'Contains(text or list, value)', 'True when the text contains the value (any case), or the list has it.', (args, env) => {
  arity('Contains', args, 2);
  const [c, v] = evalAll(args, env);
  return contains(toPrimitive(c), v);
});
define('StartsWith', 'StartsWith(text, start)', 'Case-insensitive.', (args, env) => {
  arity('StartsWith', args, 2);
  const [t, s] = evalAll(args, env);
  return toText(t).toLowerCase().indexOf(toText(s).toLowerCase()) === 0;
});
define('EndsWith', 'EndsWith(text, end)', 'Case-insensitive.', (args, env) => {
  arity('EndsWith', args, 2);
  const [t, s] = evalAll(args, env);
  const text = toText(t).toLowerCase();
  const end = toText(s).toLowerCase();
  return text.length >= end.length && text.substring(text.length - end.length) === end;
});
define('IsMatch', 'IsMatch(text, pattern, [flags])', 'True when the whole text matches a regular expression.', (args, env) => {
  arity('IsMatch', args, 2, 3);
  const [t, p, f] = evalAll(args, env);
  return regex(`^(?:${toText(p)})$`, args.length > 2 ? toText(f) : '').test(toText(t));
});
define('IsEmail', 'IsEmail(text)', 'True for a valid e-mail address.', (args, env) => {
  arity('IsEmail', args, 1);
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(toText(evaluateNode(args[0], env)));
});
define('IsNumber', 'IsNumber(value)', 'True when the value is a number.', (args, env) => {
  arity('IsNumber', args, 1);
  return toNumber(evaluateNode(args[0], env)) !== null;
});
define('Text', 'Text(value, [format])', 'Formats a number ("#,##0.00") or date ("dd/MM/yyyy").', (args, env) => {
  arity('Text', args, 1, 2);
  const [v, f] = evalAll(args, env);
  const value = toPrimitive(v);
  if (isBlank(value)) {
    return '';
  }
  if (args.length > 1) {
    const format = toText(f);
    if (value instanceof Date) {
      return formatDate(value, format);
    }
    const n = toNumber(value);
    if (n !== null) {
      return formatNumber(n, format);
    }
  }
  return toText(value);
});
define('Value', 'Value(text)', 'Converts text to a number.', (args, env) => {
  arity('Value', args, 1);
  return toNumber(evaluateNode(args[0], env));
});
define('Round', 'Round(number, [decimals])', 'Rounds to the given decimals.', (args, env) => {
  arity('Round', args, 1, 2);
  const [v, d] = evalAll(args, env);
  const n = toNumber(v);
  if (n === null) {
    return null;
  }
  const factor = Math.pow(10, toNumber(d) || 0);
  return Math.round(n * factor) / factor;
});
define('Abs', 'Abs(number)', 'Absolute value.', (args, env) => {
  arity('Abs', args, 1);
  const n = toNumber(evaluateNode(args[0], env));
  return n === null ? null : Math.abs(n);
});
define('Mod', 'Mod(number, divisor)', 'Remainder after division.', (args, env) => {
  arity('Mod', args, 2);
  const [a, b] = evalAll(args, env).map(toNumber);
  return a === null || !b ? null : a % b;
});
define('Sum', 'Sum(a, b, ...)', 'Total of the numbers; blanks count as 0.', (args, env) =>
  flatten(evalAll(args, env)).reduce((total: number, v) => total + (toNumber(v) || 0), 0)
);
define('Min', 'Min(a, b, ...)', 'Smallest number.', (args, env) => {
  const nums = flatten(evalAll(args, env)).map(toNumber).filter((n): n is number => n !== null);
  return nums.length ? Math.min(...nums) : null;
});
define('Max', 'Max(a, b, ...)', 'Largest number.', (args, env) => {
  const nums = flatten(evalAll(args, env)).map(toNumber).filter((n): n is number => n !== null);
  return nums.length ? Math.max(...nums) : null;
});
define('Today', 'Today()', "Today's date, at midnight.", (args, env) => {
  arity('Today', args, 0);
  return startOfDay(env.now ? env.now() : new Date());
});
define('Now', 'Now()', 'Current date and time.', (args, env) => {
  arity('Now', args, 0);
  return env.now ? env.now() : new Date();
});
define('Date', 'Date(year, month, day)', 'Builds a date; month is 1-12.', (args, env) => {
  arity('Date', args, 3);
  const [y, m, d] = evalAll(args, env).map(toNumber);
  return new Date(y || 0, (m || 1) - 1, d || 1);
});
define('DateValue', 'DateValue(text)', 'Converts text such as "2026-12-31" to a date.', (args, env) => {
  arity('DateValue', args, 1);
  return toDate(evaluateNode(args[0], env));
});
define('DateAdd', 'DateAdd(date, amount, [unit])', 'Adds days (default), weeks, months, years, hours or minutes.', (args, env) => {
  arity('DateAdd', args, 2, 3);
  const [d, n, u] = evalAll(args, env);
  const date = toDate(d);
  if (!date) {
    return null;
  }
  return addToDate(date, toNumber(n) || 0, args.length > 2 ? toText(u).toLowerCase() : 'days');
});
define('DateDiff', 'DateDiff(start, end, [unit])', 'Difference in days (default), weeks, months, years, hours or minutes.', (args, env) => {
  arity('DateDiff', args, 2, 3);
  const [a, b, u] = evalAll(args, env);
  const start = toDate(a);
  const end = toDate(b);
  if (!start || !end) {
    return null;
  }
  return diffDates(start, end, args.length > 2 ? toText(u).toLowerCase() : 'days');
});
define('Year', 'Year(date)', 'Year number.', (args, env) => {
  arity('Year', args, 1);
  const d = toDate(evaluateNode(args[0], env));
  return d ? d.getFullYear() : null;
});
define('Month', 'Month(date)', 'Month number, 1-12.', (args, env) => {
  arity('Month', args, 1);
  const d = toDate(evaluateNode(args[0], env));
  return d ? d.getMonth() + 1 : null;
});
define('Day', 'Day(date)', 'Day of the month.', (args, env) => {
  arity('Day', args, 1);
  const d = toDate(evaluateNode(args[0], env));
  return d ? d.getDate() : null;
});
define('Weekday', 'Weekday(date)', 'Day of the week, 1 (Sunday) - 7 (Saturday).', (args, env) => {
  arity('Weekday', args, 1);
  const d = toDate(evaluateNode(args[0], env));
  return d ? d.getDay() + 1 : null;
});
define('UserName', 'UserName()', 'Display name of the person filling in the form.', (args, env) => {
  arity('UserName', args, 0);
  return env.user.name;
});
define('UserEmail', 'UserEmail()', 'E-mail of the person filling in the form.', (args, env) => {
  arity('UserEmail', args, 0);
  return env.user.email;
});
define('UserLogin', 'UserLogin()', 'Login name of the person filling in the form.', (args, env) => {
  arity('UserLogin', args, 0);
  return env.user.loginName;
});
define('InGroup', 'InGroup("Group name")', 'True when the current user is in the SharePoint group.', (args, env) => {
  arity('InGroup', args, 1);
  const name = toText(evaluateNode(args[0], env)).toLowerCase();
  return env.groups.some((g) => g.toLowerCase() === name);
});
define('FormMode', 'FormMode()', '"New", "Edit" or "Display".', (args, env) => {
  arity('FormMode', args, 0);
  return env.mode;
});
define('ItemId', 'ItemId()', 'Id of the item being edited; blank on a new item.', (args, env) => {
  arity('ItemId', args, 0);
  return env.itemId === undefined ? null : env.itemId;
});

export function functionCatalog(): { name: string; signature: string; help: string }[] {
  return Object.keys(FUNCTIONS)
    .sort()
    .map((key) => ({ name: FUNCTIONS[key].display, signature: FUNCTIONS[key].signature, help: FUNCTIONS[key].help }));
}

export function isKnownFunction(name: string): boolean {
  return Object.prototype.hasOwnProperty.call(FUNCTIONS, name.toLowerCase());
}

function member(obj: unknown, prop: string): unknown {
  const p = prop.toLowerCase();
  if (isBlank(obj)) {
    return null;
  }
  if (Array.isArray(obj)) {
    if (p === 'count' || p === 'length') {
      return obj.length;
    }
    // [Reviewers].Email -> list of e-mails.
    return obj.map((o) => member(o, prop));
  }
  if (typeof obj === 'string' && p === 'length') {
    return obj.length;
  }
  if (typeof obj === 'object' && !(obj instanceof Date)) {
    const o = obj as { [key: string]: unknown };
    for (const key of Object.keys(o)) {
      if (key.toLowerCase() === p) {
        return o[key];
      }
    }
    return null;
  }
  throw new ExpressionError(`A ${typeof obj} value has no property "${prop}"`);
}

function arithmetic(op: string, a: unknown, b: unknown): unknown {
  const x = toPrimitive(a);
  const y = toPrimitive(b);
  // Date + number adds days; date - date is days between.
  if (x instanceof Date && op === '+' && !(y instanceof Date)) {
    return addToDate(x, toNumber(y) || 0, 'days');
  }
  if (x instanceof Date && op === '-') {
    if (y instanceof Date) {
      return (x.getTime() - y.getTime()) / 86400000;
    }
    return addToDate(x, -(toNumber(y) || 0), 'days');
  }
  const nx = toNumber(x);
  const ny = toNumber(y);
  if (nx === null && ny === null) {
    return null;
  }
  const l = nx || 0;
  const r = ny || 0;
  switch (op) {
    case '+': return l + r;
    case '-': return l - r;
    case '*': return l * r;
    default:
      if (r === 0) {
        throw new ExpressionError('Division by zero');
      }
      return l / r;
  }
}

export function evaluateNode(node: Node, env: IExpressionEnv): unknown {
  switch (node.k) {
    case 'lit':
      return node.v;
    case 'field':
      if (env.hasField && !env.hasField(node.name)) {
        throw new ExpressionError(`There is no field with the internal name [${node.name}]`);
      }
      return env.getField(node.name);
    case 'member':
      return member(evaluateNode(node.obj, env), node.prop);
    case 'un': {
      const v = evaluateNode(node.e, env);
      if (node.op === '!') {
        return !toBool(v);
      }
      const n = toNumber(v);
      return n === null ? null : node.op === '-' ? -n : n;
    }
    case 'bin': {
      if (node.op === '&&') {
        return toBool(evaluateNode(node.l, env)) && toBool(evaluateNode(node.r, env));
      }
      if (node.op === '||') {
        return toBool(evaluateNode(node.l, env)) || toBool(evaluateNode(node.r, env));
      }
      const l = evaluateNode(node.l, env);
      const r = evaluateNode(node.r, env);
      switch (node.op) {
        case '=':
          return looseEquals(l, r);
        case '!=':
          return !looseEquals(l, r);
        case '<': {
          const c = compare(l, r);
          return c !== null && c < 0;
        }
        case '>': {
          const c = compare(l, r);
          return c !== null && c > 0;
        }
        case '<=': {
          const c = compare(l, r);
          return c !== null && c <= 0;
        }
        case '>=': {
          const c = compare(l, r);
          return c !== null && c >= 0;
        }
        case 'in':
          return contains(toPrimitive(r), l);
        case '&':
          return toText(l) + toText(r);
        default:
          return arithmetic(node.op, l, r);
      }
    }
    case 'call': {
      const fn = FUNCTIONS[node.name];
      if (!fn) {
        throw new ExpressionError(`Unknown function ${node.name}()`);
      }
      return fn.fn(node.args, env);
    }
  }
  return null;
}

export function evaluate(expression: string, env: IExpressionEnv): unknown {
  return evaluateNode(parse(expression), env);
}

/**
 * Evaluates an optional expression, falling back when it is empty or fails.
 * Used for visibility / required / read-only rules: a broken rule must never
 * stop someone from filling in the form, so errors resolve to the fallback
 * (the designer shows the error to the form author instead).
 */
export function evaluateOr<T>(expression: string | undefined, env: IExpressionEnv, fallback: T, convert: (v: unknown) => T): T {
  if (!expression || !expression.trim()) {
    return fallback;
  }
  try {
    return convert(evaluate(expression, env));
  } catch {
    return fallback;
  }
}

/**
 * Replaces `{{ expression }}` placeholders. `escape` lets callers make the
 * values safe for where the text ends up (URL component, OData literal).
 */
export function renderTemplate(template: string | undefined, env: IExpressionEnv, escape?: (value: string) => string): string {
  if (!template) {
    return '';
  }
  return template.replace(/\{\{([\s\S]+?)\}\}/g, (_match, expr: string) => {
    let text: string;
    try {
      text = toText(evaluate(expr.trim(), env));
    } catch {
      text = '';
    }
    return escape ? escape(text) : text;
  });
}

export interface IExpressionCheck {
  error?: string;
  /** Field names referenced that are not on the list. */
  unknownFields: string[];
}

/** Design-time check: syntax, unknown functions and unknown fields. */
export function checkExpression(expression: string, knownFields: string[]): IExpressionCheck {
  if (!expression || !expression.trim()) {
    return { unknownFields: [] };
  }
  try {
    const node = parse(expression);
    const unknownFn = functionNames(node).filter((n) => !isKnownFunction(n));
    if (unknownFn.length) {
      return { error: `Unknown function ${unknownFn[0]}()`, unknownFields: [] };
    }
    const refs = fieldRefs(node);
    return { unknownFields: refs.filter((r) => knownFields.indexOf(r) < 0) };
  } catch (e) {
    return { error: (e as Error).message, unknownFields: [] };
  }
}
