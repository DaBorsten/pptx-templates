/**
 * Small, side-effect free expression language for template tags
 * (`{{IF user.age >= 18}}`, `{{price.toFixed(2)}}`,
 * `{{FOR p IN projects.filter(p => p.active)}}`, …).
 *
 * Deliberately no `eval`/`vm`: a template comes from a .pptx file, so it is
 * treated as untrusted input. Supported: literals (incl. object and array
 * literals), identifiers, member access (`a.b`, `a[0]`, `a['b']`, `a?.b`),
 * calls of functions reachable from the data, arrow functions
 * (`x => x.name`, `(a, b) => a - b`), unary `! - +`, `* / %`, `+ -`,
 * comparisons, `== != === !==`, `&& || ??` and the ternary `? :`.
 * Prototype-reaching keys are rejected.
 */

export type TemplateScope = {
  lookup(name: string): { found: boolean; value?: unknown };
};

type Token =
  | { type: 'num'; value: number }
  | { type: 'str'; value: string }
  | { type: 'id'; value: string }
  | { type: 'op'; value: string };

type Node =
  | { type: 'lit'; value: unknown }
  | { type: 'id'; name: string }
  | { type: 'member'; object: Node; property: Node; optional?: boolean }
  | { type: 'call'; callee: Node; args: Node[] }
  | { type: 'unary'; op: string; arg: Node }
  | { type: 'binary'; op: string; left: Node; right: Node }
  | { type: 'cond'; test: Node; then: Node; else: Node }
  | { type: 'array'; items: Node[] }
  | { type: 'object'; entries: { key: Node; value: Node }[] }
  | { type: 'arrow'; params: string[]; body: Node };

const OPERATORS = [
  '===',
  '!==',
  '==',
  '!=',
  '=>',
  '<=',
  '>=',
  '&&',
  '||',
  '??',
  '?.',
  '<',
  '>',
  '+',
  '-',
  '*',
  '/',
  '%',
  '!',
  '?',
  ':',
  '.',
  ',',
  '(',
  ')',
  '[',
  ']',
  '{',
  '}',
];

/** PowerPoint's AutoCorrect turns typed quotes into typographic ones. */
export const QUOTES: Record<string, string> = {
  '"': '"',
  "'": "'",
  '“': '”',
  '„': '“',
  '‘': '’',
  '‚': '‘',
};

const LITERALS: Record<string, unknown> = {
  true: true,
  false: false,
  null: null,
  undefined: undefined,
};

const FORBIDDEN_KEYS = ['constructor', 'prototype'];

export class TemplateExpressionError extends Error {
  constructor(message: string, expression: string) {
    super(`${message} in template expression "${expression}"`);
    this.name = 'TemplateExpressionError';
  }
}

class ArrowScope implements TemplateScope {
  constructor(
    private readonly vars: Record<string, unknown>,
    private readonly parent: TemplateScope,
  ) {}

  lookup(name: string): { found: boolean; value?: unknown } {
    if (Object.prototype.hasOwnProperty.call(this.vars, name)) {
      return { found: true, value: this.vars[name] };
    }
    return this.parent.lookup(name);
  }
}

export class TemplateExpression {
  readonly source: string;
  private readonly ast: Node;
  private tokens: Token[];
  private pos = 0;

  constructor(source: string) {
    this.source = source;
    this.tokens = this.tokenize(source);
    if (!this.tokens.length) this.fail('Empty expression');
    this.ast = this.parseExpression();
    if (this.pos < this.tokens.length) {
      this.fail(`Unexpected "${this.tokens[this.pos].value}"`);
    }
    this.tokens = undefined;
  }

  evaluate(scope: TemplateScope, strict = false): unknown {
    return this.evaluateNode(this.ast, scope, strict);
  }

  private fail(message: string): never {
    throw new TemplateExpressionError(message, this.source);
  }

  /*
   * Tokenizer
   */

  private tokenize(source: string): Token[] {
    const tokens: Token[] = [];
    let i = 0;
    while (i < source.length) {
      const char = source[i];
      if (/\s/.test(char)) {
        i++;
        continue;
      }
      if (QUOTES[char]) {
        const close = QUOTES[char];
        let value = '';
        i++;
        while (i < source.length && source[i] !== close) {
          if (source[i] === '\\' && i + 1 < source.length) i++;
          value += source[i++];
        }
        if (i >= source.length) this.fail('Unterminated string');
        i++;
        tokens.push({ type: 'str', value });
        continue;
      }
      const num = /^\d+(\.\d+)?/.exec(source.slice(i));
      if (num) {
        tokens.push({ type: 'num', value: Number(num[0]) });
        i += num[0].length;
        continue;
      }
      const id = /^[\p{L}_$][\p{L}\p{N}_$]*/u.exec(source.slice(i));
      if (id) {
        tokens.push({ type: 'id', value: id[0] });
        i += id[0].length;
        continue;
      }
      const op = OPERATORS.find((o) => source.startsWith(o, i));
      if (!op) this.fail(`Unexpected character "${char}"`);
      tokens.push({ type: 'op', value: op });
      i += op.length;
    }
    return tokens;
  }

  /*
   * Parser (recursive descent)
   */

  private peekOp(...ops: string[]): string | undefined {
    return this.opAt(this.pos, ...ops);
  }

  private opAt(pos: number, ...ops: string[]): string | undefined {
    const token = this.tokens[pos];
    return token?.type === 'op' && ops.includes(token.value)
      ? token.value
      : undefined;
  }

  private expectOp(op: string): void {
    if (!this.peekOp(op)) this.fail(`Expected "${op}"`);
    this.pos++;
  }

  private parseExpression(): Node {
    return this.parseArrow() ?? this.parseTernary();
  }

  /** `x => body` or `(a, b) => body`; undefined when not an arrow. */
  private parseArrow(): Node | undefined {
    const token = this.tokens[this.pos];
    if (token?.type === 'id' && this.opAt(this.pos + 1, '=>')) {
      this.pos += 2;
      return { type: 'arrow', params: [token.value], body: this.parseExpression() };
    }
    if (!this.peekOp('(')) return undefined;

    const params: string[] = [];
    let p = this.pos + 1;
    while (!this.opAt(p, ')')) {
      const param = this.tokens[p];
      if (param?.type !== 'id') return undefined;
      params.push(param.value);
      p++;
      if (this.opAt(p, ',')) p++;
      else if (!this.opAt(p, ')')) return undefined;
    }
    if (!this.opAt(p + 1, '=>')) return undefined;
    this.pos = p + 2;
    return { type: 'arrow', params, body: this.parseExpression() };
  }

  private parseTernary(): Node {
    const test = this.parseBinary(0);
    if (!this.peekOp('?')) return test;
    this.pos++;
    const thenNode = this.parseExpression();
    this.expectOp(':');
    return { type: 'cond', test, then: thenNode, else: this.parseExpression() };
  }

  private static readonly PRECEDENCE = [
    ['??'],
    ['||'],
    ['&&'],
    ['==', '!=', '===', '!=='],
    ['<', '>', '<=', '>='],
    ['+', '-'],
    ['*', '/', '%'],
  ];

  private parseBinary(level: number): Node {
    if (level >= TemplateExpression.PRECEDENCE.length) return this.parseUnary();
    let left = this.parseBinary(level + 1);
    let op: string;
    while ((op = this.peekOp(...TemplateExpression.PRECEDENCE[level]))) {
      this.pos++;
      left = { type: 'binary', op, left, right: this.parseBinary(level + 1) };
    }
    return left;
  }

  private parseUnary(): Node {
    const op = this.peekOp('!', '-', '+');
    if (op) {
      this.pos++;
      return { type: 'unary', op, arg: this.parseUnary() };
    }
    return this.parsePostfix();
  }

  private parsePostfix(): Node {
    let node = this.parsePrimary();
    for (;;) {
      const op = this.peekOp('.', '?.', '[', '(');
      if (op === '.' || op === '?.') {
        this.pos++;
        if (op === '?.' && this.peekOp('[')) {
          this.pos++;
          const property = this.parseExpression();
          this.expectOp(']');
          node = { type: 'member', object: node, property, optional: true };
          continue;
        }
        const token = this.tokens[this.pos++];
        if (token?.type !== 'id') this.fail('Expected property name');
        node = {
          type: 'member',
          object: node,
          property: { type: 'lit', value: token.value },
          optional: op === '?.',
        };
      } else if (op === '[') {
        this.pos++;
        const property = this.parseExpression();
        this.expectOp(']');
        node = { type: 'member', object: node, property };
      } else if (op === '(') {
        this.pos++;
        node = { type: 'call', callee: node, args: this.parseList(')') };
      } else {
        return node;
      }
    }
  }

  /** Comma separated expressions up to (and consuming) `close`. */
  private parseList(close: string): Node[] {
    const items: Node[] = [];
    while (!this.peekOp(close)) {
      items.push(this.parseExpression());
      if (!this.peekOp(close)) this.expectOp(',');
    }
    this.pos++;
    return items;
  }

  private parsePrimary(): Node {
    const token = this.tokens[this.pos++];
    if (!token) this.fail('Unexpected end');
    if (token.type === 'num' || token.type === 'str') {
      return { type: 'lit', value: token.value };
    }
    if (token.type === 'id') {
      return token.value in LITERALS
        ? { type: 'lit', value: LITERALS[token.value] }
        : { type: 'id', name: token.value };
    }
    switch (token.value) {
      case '(': {
        const node = this.parseExpression();
        this.expectOp(')');
        return node;
      }
      case '[':
        return { type: 'array', items: this.parseList(']') };
      case '{':
        return this.parseObject();
    }
    this.fail(`Unexpected "${token.value}"`);
  }

  private parseObject(): Node {
    const entries: { key: Node; value: Node }[] = [];
    while (!this.peekOp('}')) {
      const token = this.tokens[this.pos++];
      let key: Node;
      if (token?.type === 'id' || token?.type === 'str' || token?.type === 'num') {
        key = { type: 'lit', value: String(token.value) };
      } else if (token?.type === 'op' && token.value === '[') {
        key = this.parseExpression();
        this.expectOp(']');
      } else {
        this.fail('Expected property name');
      }
      // Shorthand `{ name }`
      const value: Node =
        token.type === 'id' && (this.peekOp(',') || this.peekOp('}'))
          ? { type: 'id', name: token.value }
          : (this.expectOp(':'), this.parseExpression());
      entries.push({ key, value });
      if (!this.peekOp('}')) this.expectOp(',');
    }
    this.pos++;
    return { type: 'object', entries };
  }

  /*
   * Evaluation
   */

  private evaluateNode(
    node: Node,
    scope: TemplateScope,
    strict: boolean,
  ): unknown {
    const evaluate = (n: Node) => this.evaluateNode(n, scope, strict);
    switch (node.type) {
      case 'lit':
        return node.value;
      case 'id': {
        const result = scope.lookup(node.name);
        if (!result.found && strict) this.fail(`"${node.name}" is not defined`);
        return result.value;
      }
      case 'member': {
        const object = evaluate(node.object);
        if (node.optional && (object === null || object === undefined)) {
          return undefined;
        }
        return this.getMember(object, evaluate(node.property));
      }
      case 'call': {
        let thisArg: unknown;
        let fn: unknown;
        if (node.callee.type === 'member') {
          thisArg = evaluate(node.callee.object);
          if (
            node.callee.optional &&
            (thisArg === null || thisArg === undefined)
          ) {
            return undefined;
          }
          fn = this.getMember(thisArg, evaluate(node.callee.property));
        } else {
          fn = evaluate(node.callee);
        }
        if (typeof fn !== 'function') this.fail('Not a function');
        return fn.apply(thisArg, node.args.map(evaluate));
      }
      case 'unary': {
        const arg = evaluate(node.arg) as number;
        if (node.op === '!') return !arg;
        return node.op === '-' ? -arg : +arg;
      }
      case 'cond':
        return evaluate(node.test) ? evaluate(node.then) : evaluate(node.else);
      case 'binary':
        return this.evaluateBinary(node, evaluate);
      case 'array':
        return node.items.map(evaluate);
      case 'object': {
        const result: Record<string, unknown> = {};
        node.entries.forEach((entry) => {
          const key = this.checkKey(evaluate(entry.key));
          result[key] = evaluate(entry.value);
        });
        return result;
      }
      case 'arrow':
        return (...args: unknown[]) => {
          const vars: Record<string, unknown> = {};
          node.params.forEach((param, i) => (vars[param] = args[i]));
          return this.evaluateNode(
            node.body,
            new ArrowScope(vars, scope),
            strict,
          );
        };
    }
  }

  private evaluateBinary(
    node: Extract<Node, { type: 'binary' }>,
    evaluate: (n: Node) => unknown,
  ): unknown {
    // Typed as number only to satisfy the compiler: the operators keep their
    // JavaScript semantics at runtime ('a' + 'b', 'a' < 'b', …).
    const left = evaluate(node.left) as number;
    if (node.op === '&&') return left && evaluate(node.right);
    if (node.op === '||') return left || evaluate(node.right);
    if (node.op === '??') return left ?? evaluate(node.right);
    const right = evaluate(node.right) as number;
    switch (node.op) {
      case '===':
        return left === right;
      case '!==':
        return left !== right;
      case '==':
        return left == right;
      case '!=':
        return left != right;
      case '<':
        return left < right;
      case '>':
        return left > right;
      case '<=':
        return left <= right;
      case '>=':
        return left >= right;
      case '+':
        return left + right;
      case '-':
        return left - right;
      case '*':
        return left * right;
      case '/':
        return left / right;
      case '%':
        return left % right;
    }
  }

  private checkKey(property: unknown): string {
    const key = String(property);
    if (key.startsWith('__') || FORBIDDEN_KEYS.includes(key)) {
      this.fail(`Access to "${key}" is not allowed`);
    }
    return key;
  }

  private getMember(object: unknown, property: unknown): unknown {
    const key = this.checkKey(property);
    if (object === null || object === undefined) return undefined;
    return (object as Record<string, unknown>)[key];
  }
}

/**
 * Splits `a = 1; b = a + 1` at top-level semicolons (outside strings and
 * brackets) and separates assignments from plain expressions.
 */
export const parseStatements = (
  source: string,
): { assign?: string; expr: TemplateExpression }[] => {
  const parts: string[] = [];
  let depth = 0;
  let quote: string | undefined;
  let start = 0;
  for (let i = 0; i < source.length; i++) {
    const char = source[i];
    if (quote) {
      if (char === '\\') i++;
      else if (char === quote) quote = undefined;
    } else if (QUOTES[char]) {
      quote = QUOTES[char];
    } else if ('([{'.includes(char)) {
      depth++;
    } else if (')]}'.includes(char)) {
      depth--;
    } else if (char === ';' && depth === 0) {
      parts.push(source.slice(start, i));
      start = i + 1;
    }
  }
  parts.push(source.slice(start));

  return parts
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part) => {
      const m =
        /^(?:(?:let|const|var)\s+)?([\p{L}_$][\p{L}\p{N}_$]*)\s*=(?![=>])([\s\S]*)$/u.exec(
          part,
        );
      return m
        ? { assign: m[1], expr: new TemplateExpression(m[2].trim()) }
        : { expr: new TemplateExpression(part) };
    });
};
