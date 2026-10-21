/**
 * Parser for the form expression language - a small, Excel / Power Fx
 * flavoured syntax that replaces the custom JavaScript Plumsail runs and
 * the Power Fx formulas Power Apps runs:
 *
 *   [Status] = "Approved" && [Amount] > 1000
 *   If(IsBlank([DueDate]), "No due date", Text([DueDate], "dd MMM yyyy"))
 *   [Manager].Email & ";" & UserEmail()
 *
 * Field values are referenced by internal name in square brackets. Parsing
 * produces a plain AST that the evaluator walks; nothing is ever compiled to
 * or run as JavaScript, so a form definition cannot execute code.
 */

export type Node =
  | { k: 'lit'; v: unknown }
  | { k: 'field'; name: string }
  | { k: 'call'; name: string; args: Node[] }
  | { k: 'un'; op: string; e: Node }
  | { k: 'bin'; op: string; l: Node; r: Node }
  | { k: 'member'; obj: Node; prop: string };

export class ExpressionError extends Error {
  constructor(message: string, public position?: number) {
    super(message);
    this.name = 'ExpressionError';
  }
}

type TokenType = 'num' | 'str' | 'field' | 'ident' | 'op' | 'eof';

interface IToken {
  t: TokenType;
  v: string;
  pos: number;
}

const OPERATORS = ['&&', '||', '==', '!=', '<>', '<=', '>=', '=', '<', '>', '+', '-', '*', '/', '&', '!', '(', ')', ',', '.'];

function isDigit(ch: string): boolean {
  return ch >= '0' && ch <= '9';
}

function isIdentStart(ch: string): boolean {
  return /[A-Za-z_]/.test(ch);
}

function isIdentPart(ch: string): boolean {
  return /[A-Za-z0-9_]/.test(ch);
}

export function tokenize(src: string): IToken[] {
  const tokens: IToken[] = [];
  let i = 0;
  while (i < src.length) {
    const ch = src.charAt(i);
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    const start = i;
    if (isDigit(ch) || (ch === '.' && isDigit(src.charAt(i + 1)))) {
      while (i < src.length && (isDigit(src.charAt(i)) || src.charAt(i) === '.')) {
        i++;
      }
      tokens.push({ t: 'num', v: src.substring(start, i), pos: start });
      continue;
    }
    if (ch === '"' || ch === "'") {
      const quote = ch;
      let value = '';
      i++;
      let closed = false;
      while (i < src.length) {
        const c = src.charAt(i);
        if (c === '\\' && i + 1 < src.length) {
          const next = src.charAt(i + 1);
          value += next === 'n' ? '\n' : next === 't' ? '\t' : next;
          i += 2;
          continue;
        }
        if (c === quote) {
          // Doubled quote is an escaped quote, as in Excel / Power Fx.
          if (src.charAt(i + 1) === quote) {
            value += quote;
            i += 2;
            continue;
          }
          closed = true;
          i++;
          break;
        }
        value += c;
        i++;
      }
      if (!closed) {
        throw new ExpressionError('Text is missing its closing quote', start);
      }
      tokens.push({ t: 'str', v: value, pos: start });
      continue;
    }
    if (ch === '[') {
      const end = src.indexOf(']', i);
      if (end < 0) {
        throw new ExpressionError('Field reference is missing its closing ]', start);
      }
      const name = src.substring(i + 1, end).trim();
      if (!name) {
        throw new ExpressionError('Empty field reference []', start);
      }
      tokens.push({ t: 'field', v: name, pos: start });
      i = end + 1;
      continue;
    }
    if (isIdentStart(ch)) {
      while (i < src.length && isIdentPart(src.charAt(i))) {
        i++;
      }
      tokens.push({ t: 'ident', v: src.substring(start, i), pos: start });
      continue;
    }
    let matched = '';
    for (const op of OPERATORS) {
      if (src.substr(i, op.length) === op) {
        matched = op;
        break;
      }
    }
    if (!matched) {
      throw new ExpressionError(`Unexpected character "${ch}"`, start);
    }
    tokens.push({ t: 'op', v: matched, pos: start });
    i += matched.length;
  }
  tokens.push({ t: 'eof', v: '', pos: src.length });
  return tokens;
}

/** Word operators, case-insensitive: And, Or, Not, In. */
const WORD_OPS: { [word: string]: string } = { and: '&&', or: '||', not: '!', in: 'in' };

class Parser {
  private i = 0;

  constructor(private tokens: IToken[]) {}

  private peek(): IToken {
    return this.tokens[this.i];
  }

  private next(): IToken {
    return this.tokens[this.i++];
  }

  /** Current token as an operator, mapping And/Or/Not/In words to their symbols. */
  private op(): string | undefined {
    const tok = this.peek();
    if (tok.t === 'op') {
      return tok.v;
    }
    if (tok.t === 'ident') {
      const word = tok.v.toLowerCase();
      // "Not(...)" is the Not function; And / Or / In are always operators
      // here because they cannot start an operand.
      const after = this.tokens[this.i + 1];
      if (word === 'not' && after && after.t === 'op' && after.v === '(') {
        return undefined;
      }
      return WORD_OPS[word];
    }
    return undefined;
  }

  private expect(value: string): void {
    const tok = this.next();
    if (tok.t !== 'op' || tok.v !== value) {
      throw new ExpressionError(`Expected "${value}" but found ${describe(tok)}`, tok.pos);
    }
  }

  public parse(): Node {
    const node = this.parseOr();
    const tok = this.peek();
    if (tok.t !== 'eof') {
      throw new ExpressionError(`Unexpected ${describe(tok)}`, tok.pos);
    }
    return node;
  }

  private parseOr(): Node {
    let left = this.parseAnd();
    while (this.op() === '||') {
      this.next();
      left = { k: 'bin', op: '||', l: left, r: this.parseAnd() };
    }
    return left;
  }

  private parseAnd(): Node {
    let left = this.parseNot();
    while (this.op() === '&&') {
      this.next();
      left = { k: 'bin', op: '&&', l: left, r: this.parseNot() };
    }
    return left;
  }

  private parseNot(): Node {
    if (this.op() === '!') {
      this.next();
      return { k: 'un', op: '!', e: this.parseNot() };
    }
    return this.parseComparison();
  }

  private parseComparison(): Node {
    let left = this.parseConcat();
    for (;;) {
      const op = this.op();
      if (op === '=' || op === '==' || op === '!=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=' || op === 'in') {
        this.next();
        const normalized = op === '==' ? '=' : op === '<>' ? '!=' : op;
        left = { k: 'bin', op: normalized, l: left, r: this.parseConcat() };
      } else {
        return left;
      }
    }
  }

  private parseConcat(): Node {
    let left = this.parseAdditive();
    while (this.op() === '&') {
      this.next();
      left = { k: 'bin', op: '&', l: left, r: this.parseAdditive() };
    }
    return left;
  }

  private parseAdditive(): Node {
    let left = this.parseMultiplicative();
    for (;;) {
      const op = this.op();
      if (op === '+' || op === '-') {
        this.next();
        left = { k: 'bin', op, l: left, r: this.parseMultiplicative() };
      } else {
        return left;
      }
    }
  }

  private parseMultiplicative(): Node {
    let left = this.parseUnary();
    for (;;) {
      const op = this.op();
      if (op === '*' || op === '/') {
        this.next();
        left = { k: 'bin', op, l: left, r: this.parseUnary() };
      } else {
        return left;
      }
    }
  }

  private parseUnary(): Node {
    const op = this.op();
    if (op === '-' || op === '+') {
      this.next();
      return { k: 'un', op, e: this.parseUnary() };
    }
    if (op === '!') {
      this.next();
      return { k: 'un', op: '!', e: this.parseUnary() };
    }
    return this.parsePostfix();
  }

  private parsePostfix(): Node {
    let node = this.parsePrimary();
    while (this.op() === '.') {
      this.next();
      const tok = this.next();
      if (tok.t !== 'ident') {
        throw new ExpressionError(`Expected a property name after "." but found ${describe(tok)}`, tok.pos);
      }
      node = { k: 'member', obj: node, prop: tok.v };
    }
    return node;
  }

  private parsePrimary(): Node {
    const tok = this.next();
    switch (tok.t) {
      case 'num': {
        const value = Number(tok.v);
        if (isNaN(value)) {
          throw new ExpressionError(`"${tok.v}" is not a number`, tok.pos);
        }
        return { k: 'lit', v: value };
      }
      case 'str':
        return { k: 'lit', v: tok.v };
      case 'field':
        return { k: 'field', name: tok.v };
      case 'ident': {
        const lower = tok.v.toLowerCase();
        const after = this.peek();
        if (after.t === 'op' && after.v === '(') {
          this.next();
          const args: Node[] = [];
          if (!(this.peek().t === 'op' && this.peek().v === ')')) {
            args.push(this.parseOr());
            while (this.peek().t === 'op' && this.peek().v === ',') {
              this.next();
              args.push(this.parseOr());
            }
          }
          this.expect(')');
          return { k: 'call', name: lower, args };
        }
        if (lower === 'true') {
          return { k: 'lit', v: true };
        }
        if (lower === 'false') {
          return { k: 'lit', v: false };
        }
        if (lower === 'null' || lower === 'blank') {
          return { k: 'lit', v: null };
        }
        throw new ExpressionError(
          `Unknown name "${tok.v}". Put field names in square brackets, e.g. [${tok.v}]`,
          tok.pos
        );
      }
      case 'op':
        if (tok.v === '(') {
          const inner = this.parseOr();
          this.expect(')');
          return inner;
        }
        throw new ExpressionError(`Unexpected "${tok.v}"`, tok.pos);
      default:
        throw new ExpressionError('The expression ended unexpectedly', tok.pos);
    }
  }
}

function describe(tok: IToken): string {
  switch (tok.t) {
    case 'eof':
      return 'the end of the expression';
    case 'str':
      return `text "${tok.v}"`;
    case 'field':
      return `field [${tok.v}]`;
    default:
      return `"${tok.v}"`;
  }
}

const cache = new Map<string, Node>();

export function parse(src: string): Node {
  const cached = cache.get(src);
  if (cached) {
    return cached;
  }
  const node = new Parser(tokenize(src)).parse();
  if (cache.size > 2000) {
    cache.clear();
  }
  cache.set(src, node);
  return node;
}

/** Internal names of every field the expression reads. */
export function fieldRefs(node: Node, into: string[] = []): string[] {
  switch (node.k) {
    case 'field':
      if (into.indexOf(node.name) < 0) {
        into.push(node.name);
      }
      break;
    case 'call':
      node.args.forEach((a) => fieldRefs(a, into));
      break;
    case 'un':
      fieldRefs(node.e, into);
      break;
    case 'bin':
      fieldRefs(node.l, into);
      fieldRefs(node.r, into);
      break;
    case 'member':
      fieldRefs(node.obj, into);
      break;
  }
  return into;
}

export function functionNames(node: Node, into: string[] = []): string[] {
  switch (node.k) {
    case 'call':
      if (into.indexOf(node.name) < 0) {
        into.push(node.name);
      }
      node.args.forEach((a) => functionNames(a, into));
      break;
    case 'un':
      functionNames(node.e, into);
      break;
    case 'bin':
      functionNames(node.l, into);
      functionNames(node.r, into);
      break;
    case 'member':
      functionNames(node.obj, into);
      break;
  }
  return into;
}
