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

/** Character offsets in the expression source; `end` is exclusive. */
export type SourceRange = { start: number; end: number };

type Token = (
  | { type: 'num'; value: number }
  | { type: 'str'; value: string }
  | { type: 'id'; value: string }
  | { type: 'op'; value: string }
) &
  SourceRange;

type Node =
  | { type: 'lit'; value: unknown; range?: SourceRange }
  | { type: 'id'; name: string; range?: SourceRange }
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

/**
 * Strings use straight quotes only (`'Ada'`, `"Ada"`). Typographic quotes
 * from PowerPoint's AutoCorrect (`’Ada’`, `“Ada”`) are a syntax error; turn
 * off "Straight quotes" with "smart quotes" in the AutoCorrect options.
 */
export const isQuote = (char: string): boolean => char === "'" || char === '"';

/** Whether `char` ends a string opened with `open`. */
export const closesQuote = (open: string, char: string): boolean =>
  char === open;

/** A Map: a plain object would also "contain" `constructor`, `__proto__`, … */
const LITERALS = new Map<string, unknown>([
  ['true', true],
  ['false', false],
  ['null', null],
  ['undefined', undefined],
]);

/** `\n`, `\t`, `\r` in a string; any other escaped character is itself. */
const ESCAPES: Record<string, string> = { n: '\n', t: '\t', r: '\r' };

const FORBIDDEN_KEYS = [
  'constructor',
  'prototype',
  '__proto__',
  '__defineGetter__',
  '__defineSetter__',
  '__lookupGetter__',
  '__lookupSetter__',
];

/** Keys that could reach prototypes; expressions may not access them. */
export const isForbiddenKey = (key: string): boolean =>
  FORBIDDEN_KEYS.includes(key);

/** Functions that compile code from a string; never callable. */
const CODE_CONSTRUCTORS: unknown[] = [
  Function,
  eval,
  Object.getPrototypeOf(async (): Promise<void> => undefined).constructor,
  Object.getPrototypeOf(function* (): Generator<number> {
    yield 0;
  }).constructor,
  Object.getPrototypeOf(async function* (): AsyncGenerator<number> {
    yield 0;
  }).constructor,
];

/** `fn.call(…)` and friends would hide the real callee from the checks. */
const INDIRECT_CALLS: unknown[] = [
  Function.prototype.call,
  Function.prototype.apply,
  Function.prototype.bind,
];

/** Compile their string argument to a RegExp, which may backtrack forever. */
const REGEX_FROM_STRING: unknown[] = [
  String.prototype.match,
  String.prototype.matchAll,
  String.prototype.search,
];

/** Longest string `repeat`, `padStart` and `padEnd` may build. */
const MAX_STRING_LENGTH = 1_000_000;

/** A name read by an expression, plus the static property path after it. */
export type ExpressionReference = {
  name: string;
  path: string[];
  /** Where the name (index 0) and each key of `path` (index i + 1) are. */
  ranges: (SourceRange | undefined)[];
  /**
   * The value itself is used: it is the result, passed to a function or put
   * into a list or object. False where it is only tested or compared
   * (`a > 1`, `!a`, the condition of `a ? b : c`).
   */
  whole?: boolean;
};

const rangeOf = (token: Token): SourceRange => ({
  start: token.start,
  end: token.end,
});

/** `a`, `a.b`, `a['b']`, `a[0]`, `a?.b` → name + keys; else undefined. */
const staticPath = (node: Node): ExpressionReference | undefined => {
  if (node.type === 'id') {
    return { name: node.name, path: [], ranges: [node.range] };
  }
  if (node.type !== 'member' || node.property.type !== 'lit') return undefined;
  const parent = staticPath(node.object);
  return (
    parent && {
      name: parent.name,
      path: [...parent.path, String(node.property.value)],
      ranges: [...parent.ranges, node.property.range],
    }
  );
};

export type EvaluateOptions = {
  /** Throw on unknown names instead of treating them as undefined. */
  strict?: boolean;
  /** Let `match`, `matchAll` and `search` compile a string to a RegExp. */
  allowStringPatterns?: boolean;
};

export class TemplateExpressionError extends Error {
  /** The expression that failed. */
  readonly expression: string;
  /** Where in `expression` it failed, for syntax errors. */
  readonly range?: SourceRange;

  constructor(message: string, expression: string, range?: SourceRange) {
    super(`${message} in template expression "${expression}"`);
    this.name = 'TemplateExpressionError';
    this.expression = expression;
    this.range = range;
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

  /** `options` as a boolean is `strict`. */
  evaluate(
    scope: TemplateScope,
    options: EvaluateOptions | boolean = {},
  ): unknown {
    return this.evaluateNode(
      this.ast,
      scope,
      typeof options === 'boolean' ? { strict: options } : options,
    );
  }

  /**
   * The names the expression reads from its scope, each with the static
   * property path after it (`user.name.toUpperCase()` → `user`,
   * `['name', 'toUpperCase']`). Arrow parameters are not included; a path
   * stops at the first computed key (`a[i].b` → `a`, plus `i`).
   *
   * @param asValue Whether the expression's result is used as a value
   * (output, assignment) rather than tested (`IF`); sets `whole`.
   */
  references(asValue = true): ExpressionReference[] {
    const refs: ExpressionReference[] = [];
    const visit = (node: Node, bound: string[], whole: boolean): void => {
      const path = staticPath(node);
      if (path) {
        if (!bound.includes(path.name)) refs.push({ ...path, whole });
        return;
      }
      switch (node.type) {
        case 'member':
          // `a[i]`: which entry is read is open, so all of `a` may be
          visit(node.object, bound, true);
          visit(node.property, bound, false);
          return;
        case 'call':
          visit(node.callee, bound, false);
          node.args.forEach((arg) => visit(arg, bound, true));
          return;
        case 'unary':
          visit(node.arg, bound, false);
          return;
        case 'binary': {
          // `a || b`, `a ?? b`: the result is one of the operands
          const passes = ['&&', '||', '??'].includes(node.op) && whole;
          visit(node.left, bound, passes);
          visit(node.right, bound, passes);
          return;
        }
        case 'cond':
          visit(node.test, bound, false);
          visit(node.then, bound, whole);
          visit(node.else, bound, whole);
          return;
        case 'array':
          node.items.forEach((item) => visit(item, bound, true));
          return;
        case 'object':
          node.entries.forEach((entry) => {
            visit(entry.key, bound, false);
            visit(entry.value, bound, true);
          });
          return;
        case 'arrow':
          visit(node.body, [...bound, ...node.params], true);
          return;
      }
    };
    visit(this.ast, [], asValue);
    return refs;
  }

  /**
   * Blocked keys the expression uses with a fixed name, wherever they are:
   * `a.constructor`, `'x'.constructor`, `x => x.__proto__`,
   * `{ constructor: 1 }`. Computed keys (`a[k]`) show only when evaluated.
   */
  forbiddenKeys(): { key: string; range?: SourceRange }[] {
    const found: { key: string; range?: SourceRange }[] = [];
    const check = (node: Node): void => {
      if (node.type !== 'lit') return visit(node);
      const key = String(node.value);
      if (isForbiddenKey(key)) found.push({ key, range: node.range });
    };
    const visit = (node: Node): void => {
      switch (node.type) {
        case 'member':
          visit(node.object);
          check(node.property);
          return;
        case 'call':
          visit(node.callee);
          node.args.forEach(visit);
          return;
        case 'unary':
          visit(node.arg);
          return;
        case 'binary':
          visit(node.left);
          visit(node.right);
          return;
        case 'cond':
          visit(node.test);
          visit(node.then);
          visit(node.else);
          return;
        case 'array':
          node.items.forEach(visit);
          return;
        case 'object':
          node.entries.forEach((entry) => {
            check(entry.key);
            visit(entry.value);
          });
          return;
        case 'arrow':
          visit(node.body);
          return;
      }
    };
    visit(this.ast);
    return found;
  }

  /** The expression as a plain path (`order.items`), else undefined. */
  path(): ExpressionReference | undefined {
    return staticPath(this.ast);
  }

  /**
   * While parsing, the error points at the current token (or the end of
   * the source); the tokenizer passes its own range.
   */
  private fail(message: string, range?: SourceRange): never {
    const end = this.source.length;
    const at =
      range ??
      (this.tokens
        ? (this.tokens[this.pos] ?? { start: end, end })
        : undefined);
    throw new TemplateExpressionError(
      message,
      this.source,
      at && { start: at.start, end: at.end },
    );
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
      const start = i;
      if (isQuote(char)) {
        let value = '';
        i++;
        while (i < source.length && !closesQuote(char, source[i])) {
          if (source[i] === '\\' && i + 1 < source.length) {
            const escaped = source[i + 1];
            value += Object.prototype.hasOwnProperty.call(ESCAPES, escaped)
              ? ESCAPES[escaped]
              : escaped;
            i += 2;
          } else {
            value += source[i++];
          }
        }
        if (i >= source.length) {
          this.fail('Unterminated string', { start, end: source.length });
        }
        i++;
        tokens.push({ type: 'str', value, start, end: i });
        continue;
      }
      const num = /^(?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(
        source.slice(i),
      );
      if (num) {
        i += num[0].length;
        tokens.push({ type: 'num', value: Number(num[0]), start, end: i });
        continue;
      }
      const id = /^[\p{L}_$][\p{L}\p{N}_$]*/u.exec(source.slice(i));
      if (id) {
        i += id[0].length;
        tokens.push({ type: 'id', value: id[0], start, end: i });
        continue;
      }
      const op = OPERATORS.find((o) => source.startsWith(o, i));
      if (!op) {
        this.fail(`Unexpected character "${char}"`, { start, end: start + 1 });
      }
      i += op.length;
      tokens.push({ type: 'op', value: op, start, end: i });
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
      return {
        type: 'arrow',
        params: [token.value],
        body: this.parseExpression(),
      };
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
        if (token?.type !== 'id') {
          this.fail('Expected property name', token && rangeOf(token));
        }
        node = {
          type: 'member',
          object: node,
          property: { type: 'lit', value: token.value, range: rangeOf(token) },
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
    const range = rangeOf(token);
    if (token.type === 'num' || token.type === 'str') {
      return { type: 'lit', value: token.value, range };
    }
    if (token.type === 'id') {
      return LITERALS.has(token.value)
        ? { type: 'lit', value: LITERALS.get(token.value), range }
        : { type: 'id', name: token.value, range };
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
    this.fail(`Unexpected "${token.value}"`, range);
  }

  private parseObject(): Node {
    const entries: { key: Node; value: Node }[] = [];
    while (!this.peekOp('}')) {
      const token = this.tokens[this.pos++];
      let key: Node;
      if (
        token?.type === 'id' ||
        token?.type === 'str' ||
        token?.type === 'num'
      ) {
        key = {
          type: 'lit',
          value: String(token.value),
          range: rangeOf(token),
        };
      } else if (token?.type === 'op' && token.value === '[') {
        key = this.parseExpression();
        this.expectOp(']');
      } else {
        this.fail('Expected property name', token && rangeOf(token));
      }
      // Shorthand `{ name }`
      const value: Node =
        token.type === 'id' && (this.peekOp(',') || this.peekOp('}'))
          ? { type: 'id', name: token.value, range: rangeOf(token) }
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
    options: EvaluateOptions,
  ): unknown {
    const evaluate = (n: Node) => this.evaluateNode(n, scope, options);
    switch (node.type) {
      case 'lit':
        return node.value;
      case 'id': {
        const result = scope.lookup(node.name);
        if (!result.found && options.strict) {
          this.fail(`"${node.name}" is not defined`);
        }
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
        const args = node.args.map(evaluate);
        this.checkCall(fn, thisArg, args, options);
        return fn.apply(thisArg, args);
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
            options,
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

  /**
   * Calls a template may not make: compiling code, hiding the callee behind
   * call/apply/bind, building a RegExp from a string (catastrophic
   * backtracking blocks the process; `allowStringPatterns` permits it) and
   * building huge strings.
   */
  private checkCall(
    fn: unknown,
    thisArg: unknown,
    args: unknown[],
    options: EvaluateOptions,
  ): void {
    if (CODE_CONSTRUCTORS.includes(fn)) {
      this.fail('Calling the Function constructor is not allowed');
    }
    if (INDIRECT_CALLS.includes(fn)) {
      this.fail('call, apply and bind are not allowed');
    }
    if (
      !options.allowStringPatterns &&
      REGEX_FROM_STRING.includes(fn) &&
      !(args[0] instanceof RegExp)
    ) {
      this.fail(
        'match, matchAll and search need a RegExp from the data, or the ' +
          'allowStringPatterns option (for plain text use includes, ' +
          'indexOf or startsWith)',
      );
    }
    // Legacy, but still there: recompiles a RegExp from the data in place
    if (!options.allowStringPatterns && fn === RegExp.prototype.compile) {
      this.fail(
        'compile is not allowed: it rebuilds a RegExp from a string ' +
          '(the allowStringPatterns option permits it)',
      );
    }
    const length =
      fn === String.prototype.repeat
        ? String(thisArg).length * Number(args[0])
        : fn === String.prototype.padStart || fn === String.prototype.padEnd
          ? Number(args[0])
          : 0;
    if (length > MAX_STRING_LENGTH) {
      this.fail(`A string longer than ${MAX_STRING_LENGTH} characters`);
    }
  }

  private checkKey(property: unknown): string {
    const key = String(property);
    if (isForbiddenKey(key)) {
      this.fail(`Access to "${key}" is not allowed`);
    }
    return key;
  }

  private getMember(object: unknown, property: unknown): unknown {
    const key = this.checkKey(property);
    if (object === null || object === undefined) return undefined;
    const value = (object as Record<string, unknown>)[key];
    // Rejected when read, not only when called: as a callback
    // (`list.map(fn.call, fn)`) they would never pass checkCall.
    if (INDIRECT_CALLS.includes(value)) {
      this.fail('call, apply and bind are not allowed');
    }
    if (CODE_CONSTRUCTORS.includes(value)) {
      this.fail('Calling the Function constructor is not allowed');
    }
    return value;
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
      else if (closesQuote(quote, char)) quote = undefined;
    } else if (isQuote(char)) {
      quote = char;
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
