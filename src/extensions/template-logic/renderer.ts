import { XmlDocument, XmlElement } from '../../types/xml-types';
import { XmlHelper } from '../../helper/xml-helper';
import { XmlRelationshipHelper } from '../../helper/xml-relationship-helper';
import { MultiTextHelper } from '../../helper/multitext-helper';
import { HtmlToMultiTextHelper } from '../../helper/html-to-multitext-helper';
import { RenderTemplateOptions, TemplateData } from './types';
import {
  parseStatements,
  TemplateExpression,
  TemplateScope,
} from './expression';
import { parseCommand, resolveAlias, tagSyntax } from './commands';
import { ImageTarget, loadImage, placeImages } from './images';

/** Access to the slide beyond its XML, resolved on first use. */
export type RenderContext = {
  /** The slide's <Relationships> element (LINK, HTML links, IMAGE). */
  relations?: () => Promise<XmlElement | undefined>;
  /** Archive + root template to write media files to (IMAGE). */
  images?: () => Promise<ImageTarget | undefined>;
};

/**
 * Renders docx-templates style logic written directly into slide text:
 *
 *   {{name}}  {{= expr}}  {{INS expr}}
 *   {{IF expr}} … {{ELSE-IF expr}} … {{ELSE}} … {{END-IF}}
 *   {{FOR item IN items}} … {{$item.name}} … {{$idx}} … {{END-FOR}}
 *   {{EXEC total = 0}}  {{! total = total + 1}}
 *   {{ALIAS name INS user.firstName}} … {{*name}}
 *   {{HTML expr}}  {{LINK ({ url, label })}}
 *
 * A text body is flattened into one token stream (runs, line breaks,
 * paragraph boundaries, tags), so blocks may sit inside a paragraph or span
 * several paragraphs. Paragraphs that only hold control tags are dropped.
 * In tables, a block that is not closed within its cell spans table rows and
 * repeats/removes whole rows.
 */

type Expr = Pick<TemplateExpression, 'source' | 'evaluate'>;
type Statement = { assign?: string; expr: Expr };

type Directive =
  | { type: 'insert' | 'html' | 'link' | 'image'; expr: Expr }
  | { type: 'if' | 'elseif'; expr: Expr }
  | { type: 'for'; variable: string; expr: Expr }
  | { type: 'exec'; statements: Statement[] }
  | { type: 'else' | 'endif' | 'endfor' | 'alias' };

type DirectiveType = Directive['type'];

type TagToken = {
  kind: 'tag';
  run: XmlElement;
  tag: string;
  directive: Directive;
};

type TextToken =
  | { kind: 'text'; run: XmlElement; text: string }
  | { kind: 'node'; node: XmlElement }
  | { kind: 'break'; paragraph: XmlElement }
  | TagToken;

type Block<T> =
  | { kind: 'item'; item: T }
  | {
      kind: 'if';
      branches: { expr?: Expr; tag: string; body: Block<T>[] }[];
    }
  | {
      kind: 'for';
      variable: string;
      expr: Expr;
      tag: string;
      body: Block<T>[];
    };

type BodyTokens = {
  tokens: TextToken[];
  /** Source paragraphs that contain structural tags and nothing else. */
  controlOnly: Set<XmlElement>;
};

type RowToken = {
  row?: XmlElement;
  cells?: (BodyTokens | undefined)[];
  tag?: string;
  directive?: Directive;
};

type OutputParagraph = {
  source: XmlElement;
  nodes: XmlElement[];
  /** Ready-made paragraph (HTML output); written as is. */
  raw?: XmlElement;
  /** Remainder of a paragraph split by HTML; dropped when empty. */
  continuation?: boolean;
};

type Link = { target: string; internal: boolean };

/** Tags that open, split or close a block. */
const BLOCK_TYPES: DirectiveType[] = [
  'if',
  'elseif',
  'else',
  'endif',
  'for',
  'endfor',
];
/** Tags that leave nothing in their own paragraph. */
const STRUCTURAL_TYPES: DirectiveType[] = [
  ...BLOCK_TYPES,
  'exec',
  'alias',
  'html',
  'image',
];
/** Tags that produce visible content (keep a table row). */
const CONTENT_TYPES: DirectiveType[] = ['insert', 'html', 'link'];

/** Children of <a:rPr>, schema order (CT_TextCharacterProperties). */
const RPR_CHILD_ORDER = [
  'a:ln',
  'a:noFill',
  'a:solidFill',
  'a:gradFill',
  'a:blipFill',
  'a:pattFill',
  'a:grpFill',
  'a:effectLst',
  'a:effectDag',
  'a:highlight',
  'a:uLnTx',
  'a:uLn',
  'a:uFillTx',
  'a:uFill',
  'a:latin',
  'a:ea',
  'a:cs',
  'a:sym',
  'a:hlinkClick',
  'a:hlinkMouseOver',
  'a:rtl',
  'a:extLst',
];

const REL_TYPE_HYPERLINK =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const REL_TYPE_SLIDE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';

const textControl = (token: TextToken) =>
  token.kind === 'tag' ? token : undefined;

const isThenable = (value: unknown): value is PromiseLike<unknown> =>
  !!value && typeof (value as PromiseLike<unknown>).then === 'function';

class Scope implements TemplateScope {
  constructor(
    private readonly vars: Record<string, unknown>,
    private readonly parent?: Scope,
    private readonly writable = true,
  ) {}

  lookup(name: string): { found: boolean; value?: unknown } {
    if (this.vars && Object.prototype.hasOwnProperty.call(this.vars, name)) {
      return { found: true, value: this.vars[name] };
    }
    return this.parent ? this.parent.lookup(name) : { found: false };
  }

  /**
   * Updates the nearest writable scope that defines `name`, else defines it
   * here. The user's data and additionalJsContext are never written.
   */
  assign(name: string, value: unknown): void {
    if (!this.assignExisting(name, value)) this.vars[name] = value;
  }

  private assignExisting(name: string, value: unknown): boolean {
    if (
      this.writable &&
      Object.prototype.hasOwnProperty.call(this.vars, name)
    ) {
      this.vars[name] = value;
      return true;
    }
    return this.parent?.assignExisting(name, value) ?? false;
  }
}

export class TemplateRenderError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'TemplateRenderError';
  }
}

export class TemplateRenderer {
  private readonly opening: string;
  private readonly tagPattern: RegExp;
  private readonly strict: boolean;
  private readonly root: Scope;
  private readonly aliases = new Map<string, string>();
  private readonly errors: string[] = [];
  private readonly links = new Map<string, string>();
  private relationsPromise?: Promise<XmlElement | undefined>;
  private imageTargetPromise?: Promise<ImageTarget | undefined>;
  /** IMAGE values per shape text body, placed once the body is written. */
  private readonly pendingImages = new Map<
    XmlElement,
    ReturnType<typeof loadImage>[]
  >();

  constructor(
    data: TemplateData,
    private readonly options: RenderTemplateOptions = {},
    private readonly context: RenderContext = {},
  ) {
    const syntax = tagSyntax(options);
    this.opening = syntax.opening;
    this.tagPattern = syntax.pattern;
    this.strict = options.strict ?? false;
    const jsContext = new Scope(options.additionalJsContext ?? {}, undefined, false);
    this.root = new Scope({}, new Scope(data ?? {}, jsContext, false));
  }

  async render(element: XmlDocument | XmlElement): Promise<void> {
    this.collectAliases(element);

    for (const target of this.findTargets(element)) {
      if (target.tagName === 'a:tbl') {
        await this.renderTable(target);
      } else if (this.containsTag(target)) {
        const bodyTokens = this.tokenizeBody(target);
        await this.writeBody(
          target,
          bodyTokens,
          this.parse(bodyTokens.tokens, textControl),
          this.root,
        );
        await this.placePendingImages(target);
      }
    }

    if (this.errors.length) {
      throw new TemplateRenderError(
        `Template rendering failed with ${this.errors.length} error(s):\n` +
          this.errors.map((e) => `- ${e}`).join('\n'),
      );
    }
  }

  /** Shape text bodies and tables, in document order. */
  private findTargets(element: XmlDocument | XmlElement): XmlElement[] {
    const targets: XmlElement[] = [];
    const isTarget = (el: XmlElement) =>
      el.tagName === 'a:tbl' || el.tagName === 'p:txBody';
    const visit = (parent: XmlDocument | XmlElement) => {
      this.childElements(parent as XmlElement).forEach((child) => {
        if (isTarget(child)) targets.push(child);
        else visit(child);
      });
    };
    if (isTarget(element as XmlElement)) targets.push(element as XmlElement);
    else visit(element);
    return targets;
  }

  private containsTag(element: XmlElement): boolean {
    return (element.textContent ?? '').includes(this.opening);
  }

  /** ALIAS definitions apply everywhere, so collect them up front. */
  private collectAliases(element: XmlDocument | XmlElement): void {
    XmlHelper.collectionToArray(element.getElementsByTagName('a:p')).forEach(
      (paragraph) => {
        const text = XmlHelper.collectionToArray(
          paragraph.getElementsByTagName('a:t'),
        )
          .map((t) => t.textContent)
          .join('');
        for (const match of text.matchAll(this.tagPattern)) {
          const command = parseCommand(match[1], match[0]);
          if (command.type === 'ALIAS') {
            this.aliases.set(command.name, command.code);
          }
        }
      },
    );
  }

  /*
   * Tokenizing
   */

  private tokenizeBody(body: XmlElement): BodyTokens {
    const tokens: TextToken[] = [];
    const controlOnly = new Set<XmlElement>();

    this.childElements(body, 'a:p').forEach((paragraph, index) => {
      if (index > 0) tokens.push({ kind: 'break', paragraph });
      const paragraphTokens = this.tokenizeParagraph(paragraph);
      tokens.push(...paragraphTokens);

      const isStructural = (t: TextToken) =>
        t.kind === 'tag' && STRUCTURAL_TYPES.includes(t.directive.type);
      const onlyStructural = paragraphTokens.every(
        (t) => isStructural(t) || (t.kind === 'text' && t.text.trim() === ''),
      );
      if (paragraphTokens.some(isStructural) && onlyStructural) {
        controlOnly.add(paragraph);
      }
    });

    return { tokens, controlOnly };
  }

  /**
   * PowerPoint splits text into runs at will (spell check, edits, styling),
   * so a tag like `{{user.name}}` often spans several <a:r>. Consecutive runs
   * are joined; each tag takes the formatting of the run it starts in.
   */
  private tokenizeParagraph(paragraph: XmlElement): TextToken[] {
    const tokens: TextToken[] = [];
    let segment: XmlElement[] = [];
    const flush = () => {
      tokens.push(...this.tokenizeRuns(segment));
      segment = [];
    };

    this.childElements(paragraph).forEach((child) => {
      if (child.tagName === 'a:r') {
        segment.push(child);
      } else if (
        child.tagName !== 'a:pPr' &&
        child.tagName !== 'a:endParaRPr'
      ) {
        flush();
        tokens.push({ kind: 'node', node: child });
      }
    });
    flush();

    return tokens;
  }

  private tokenizeRuns(runs: XmlElement[]): TextToken[] {
    const texts = runs.map((run) => this.runText(run));
    const full = texts.join('');
    const starts: number[] = [];
    texts.reduce((offset, text) => {
      starts.push(offset);
      return offset + text.length;
    }, 0);
    const runAt = (offset: number) => {
      let r = 0;
      while (r + 1 < runs.length && starts[r + 1] <= offset) r++;
      return r;
    };

    const tokens: TextToken[] = [];
    const pushText = (from: number, to: number) => {
      // Plain text keeps its original run boundaries and formatting.
      while (from < to) {
        const r = runAt(from);
        const end = Math.min(to, starts[r] + texts[r].length);
        tokens.push({
          kind: 'text',
          run: runs[r],
          text: full.slice(from, end),
        });
        from = end;
      }
    };

    let last = 0;
    for (const match of full.matchAll(this.tagPattern)) {
      pushText(last, match.index);
      tokens.push({
        kind: 'tag',
        run: runs[runAt(match.index)],
        tag: match[0],
        directive: this.parseDirective(match[1], match[0]),
      });
      last = match.index + match[0].length;
    }
    pushText(last, full.length);

    // An empty run (<a:t/>) still carries formatting; keep it.
    if (!tokens.length && runs.length) {
      runs.forEach((run) => tokens.push({ kind: 'node', node: run }));
    }
    return tokens;
  }

  private parseDirective(inner: string, tag: string): Directive {
    const command = parseCommand(resolveAlias(inner, this.aliases, tag), tag);
    const compile = (source: string) => this.compile(source, tag);

    switch (command.type) {
      case 'IF':
        return { type: 'if', expr: compile(command.code) };
      case 'ELSE-IF':
        return { type: 'elseif', expr: compile(command.code) };
      case 'ELSE':
        return { type: 'else' };
      case 'END-IF':
        return { type: 'endif' };
      case 'FOR':
        return {
          type: 'for',
          variable: command.name,
          expr: compile(command.code),
        };
      case 'END-FOR':
        return { type: 'endfor' };
      case 'ALIAS':
        return { type: 'alias' };
      case 'EXEC':
        return { type: 'exec', statements: this.compileStatements(command.code, tag) };
      case 'HTML':
        return { type: 'html', expr: compile(command.code) };
      case 'LINK':
        return { type: 'link', expr: compile(command.code) };
      case 'IMAGE':
        return { type: 'image', expr: compile(command.code) };
      case 'INS':
        return { type: 'insert', expr: compile(command.code) };
    }
  }

  /**
   * A syntax error fails right away - unless errors are handled
   * (errorHandler, failFast: false); then it fails when evaluated.
   */
  private compile(source: string, tag: string): Expr {
    try {
      return new TemplateExpression(source);
    } catch (error) {
      if (!this.handlesErrors) throw this.tagged(error, tag);
      return {
        source,
        evaluate: () => {
          throw error;
        },
      };
    }
  }

  private compileStatements(source: string, tag: string): Statement[] {
    try {
      return parseStatements(source);
    } catch (error) {
      if (!this.handlesErrors) throw this.tagged(error, tag);
      return [{ expr: this.compile(source, tag) }];
    }
  }

  private get handlesErrors(): boolean {
    return !!this.options.errorHandler || this.options.failFast === false;
  }

  private tagged(error: unknown, tag: string): Error {
    const message = error instanceof Error ? error.message : String(error);
    return new TemplateRenderError(`${tag}: ${message}`);
  }

  /**
   * Evaluates one command. Promises are awaited, so data functions may be
   * async. Errors go to errorHandler, are collected (failFast: false) or
   * reject the render.
   */
  private async value(tag: string, fn: () => unknown): Promise<unknown> {
    try {
      let result = fn();
      if (isThenable(result)) result = await result;
      return result;
    } catch (error) {
      const err = error instanceof Error ? error : new Error(String(error));
      if (this.options.errorHandler) {
        return this.options.errorHandler(err, tag);
      }
      if (this.options.failFast === false) {
        this.errors.push(`${tag}: ${err.message}`);
        return undefined;
      }
      throw this.tagged(err, tag);
    }
  }

  /*
   * Parsing & evaluating blocks (shared by text bodies and table rows)
   */

  private parse<T>(
    items: T[],
    control: (item: T) => { tag: string; directive: Directive } | undefined,
  ): Block<T>[] {
    const root: Block<T>[] = [];
    type Open = { block: Block<T>; tag: string; body: Block<T>[] };
    const stack: Open[] = [];
    const current = () => (stack.length ? stack[stack.length - 1].body : root);

    items.forEach((item) => {
      const tag = control(item);
      const directive = tag?.directive;
      if (!directive || !BLOCK_TYPES.includes(directive.type)) {
        current().push({ kind: 'item', item });
        return;
      }

      const top = stack[stack.length - 1];
      switch (directive.type) {
        case 'if': {
          const body: Block<T>[] = [];
          const block: Block<T> = {
            kind: 'if',
            branches: [{ expr: directive.expr, tag: tag.tag, body }],
          };
          current().push(block);
          stack.push({ block, tag: tag.tag, body });
          break;
        }
        case 'for': {
          const body: Block<T>[] = [];
          const block: Block<T> = {
            kind: 'for',
            variable: directive.variable,
            expr: directive.expr,
            tag: tag.tag,
            body,
          };
          current().push(block);
          stack.push({ block, tag: tag.tag, body });
          break;
        }
        case 'elseif':
        case 'else': {
          if (top?.block.kind !== 'if') {
            throw new TemplateRenderError(
              `${tag.tag} without a matching {{IF …}}`,
            );
          }
          const branches = top.block.branches;
          if (!branches[branches.length - 1].expr) {
            throw new TemplateRenderError(
              `${tag.tag} after {{ELSE}} in ${top.tag}`,
            );
          }
          const body: Block<T>[] = [];
          branches.push({
            expr: directive.type === 'elseif' ? directive.expr : undefined,
            tag: tag.tag,
            body,
          });
          top.body = body;
          break;
        }
        case 'endif':
        case 'endfor': {
          const expected = directive.type === 'endif' ? 'if' : 'for';
          if (top?.block.kind !== expected) {
            throw new TemplateRenderError(
              `${tag.tag} without a matching ${expected.toUpperCase()}` +
                (top ? ` (open: ${top.tag})` : ''),
            );
          }
          stack.pop();
          break;
        }
      }
    });

    if (stack.length) {
      throw new TemplateRenderError(
        `Unclosed template block ${stack[stack.length - 1].tag}`,
      );
    }
    return root;
  }

  private async evaluate<T>(
    blocks: Block<T>[],
    scope: Scope,
    emit: (item: T, scope: Scope) => Promise<void>,
  ): Promise<void> {
    for (const block of blocks) {
      if (block.kind === 'item') {
        await emit(block.item, scope);
      } else if (block.kind === 'if') {
        for (const branch of block.branches) {
          const pass =
            !branch.expr ||
            (await this.value(branch.tag, () =>
              branch.expr.evaluate(scope, this.strict),
            ));
          if (pass) {
            await this.evaluate(branch.body, scope, emit);
            break;
          }
        }
      } else {
        const value = await this.value(block.tag, () => {
          const list = block.expr.evaluate(scope, this.strict);
          if (isThenable(list)) return list.then((v) => this.toList(v, block.expr));
          return this.toList(list, block.expr);
        });
        const entries = Array.isArray(value) ? value : [];
        for (let idx = 0; idx < entries.length; idx++) {
          const loopScope = new Scope(
            {
              [block.variable]: entries[idx],
              [`$${block.variable}`]: entries[idx],
              $idx: idx,
            },
            scope,
          );
          await this.evaluate(block.body, loopScope, emit);
        }
      }
    }
  }

  private toList(value: unknown, expr: Expr): unknown[] {
    if (value === null || value === undefined) return [];
    if (typeof (value as Iterable<unknown>)[Symbol.iterator] !== 'function') {
      throw new Error(`"${expr.source}" is not a list`);
    }
    return Array.from(value as Iterable<unknown>);
  }

  /*
   * Writing text bodies
   */

  private async writeBody(
    body: XmlElement,
    bodyTokens: BodyTokens,
    blocks: Block<TextToken>[],
    scope: Scope,
  ): Promise<void> {
    const sourceParagraphs = this.childElements(body, 'a:p');
    const first = sourceParagraphs[0];
    if (!first) return;

    const paragraphs: OutputParagraph[] = [{ source: first, nodes: [] }];
    const current = () => paragraphs[paragraphs.length - 1];

    await this.evaluate(blocks, scope, async (token, tokenScope) => {
      switch (token.kind) {
        case 'break':
          paragraphs.push({ source: token.paragraph, nodes: [] });
          break;
        case 'node':
          current().nodes.push(token.node.cloneNode(true) as XmlElement);
          break;
        case 'text':
          current().nodes.push(...this.createRuns(token.run, token.text));
          break;
        case 'tag':
          await this.emitTag(token, tokenScope, body, paragraphs);
          break;
      }
    });

    let kept = paragraphs.filter(
      (p) =>
        p.raw ||
        !(bodyTokens.controlOnly.has(p.source) || p.continuation) ||
        p.nodes.some(
          (n) => n.tagName !== 'a:r' || this.runText(n).trim() !== '',
        ),
    );
    // A text body needs at least one paragraph.
    if (!kept.length) kept = [{ source: first, nodes: [] }];

    sourceParagraphs.forEach((p) => body.removeChild(p));
    kept.forEach((p) =>
      body.appendChild(p.raw ?? this.createParagraph(p)),
    );
  }

  private async emitTag(
    token: TagToken,
    scope: Scope,
    body: XmlElement,
    paragraphs: OutputParagraph[],
  ): Promise<void> {
    const { directive, tag, run } = token;
    const current = paragraphs[paragraphs.length - 1];
    const evaluate = (expr: Expr) =>
      this.value(tag, () => expr.evaluate(scope, this.strict));

    switch (directive.type) {
      case 'insert': {
        const value = await this.value(tag, async () => {
          const result = await directive.expr.evaluate(scope, this.strict);
          if (
            this.options.rejectNullish &&
            (result === null || result === undefined)
          ) {
            throw new Error(`"${directive.expr.source}" is ${result}`);
          }
          return result;
        });
        current.nodes.push(...this.createRuns(run, this.format(value)));
        break;
      }
      case 'exec':
        for (const statement of directive.statements) {
          const value = await evaluate(statement.expr);
          if (statement.assign) scope.assign(statement.assign, value);
        }
        break;
      case 'html': {
        const html = this.format(await evaluate(directive.expr));
        if (!html) break;
        const generated = this.htmlParagraphs(
          body,
          current.source,
          run,
          html,
          await this.relations(),
        );
        generated.forEach((raw) =>
          paragraphs.push({ source: current.source, nodes: [], raw }),
        );
        paragraphs.push({ source: current.source, nodes: [], continuation: true });
        break;
      }
      case 'link': {
        const value = await evaluate(directive.expr);
        if (value === null || value === undefined || value === '') break;
        const runs = await this.value(tag, async () => {
          const { link, label } = this.toLink(value);
          const relations = await this.relations();
          if (!relations) {
            throw new Error('LINK needs the slide relationships');
          }
          const relId = this.linkRelation(relations, link);
          const created = this.createRuns(run, label);
          created
            .filter((n) => n.tagName === 'a:r')
            .forEach((r) => this.setHyperlink(r, relId, link.internal));
          return created;
        });
        if (Array.isArray(runs)) current.nodes.push(...(runs as XmlElement[]));
        break;
      }
      case 'image': {
        const value = await evaluate(directive.expr);
        if (value === null || value === undefined || value === '') break;
        const image = await this.value(tag, async () => {
          if (body.tagName !== 'p:txBody') {
            throw new Error('IMAGE is not supported inside table cells');
          }
          if (!(await this.imageTarget())) {
            throw new Error(
              'IMAGE needs slide.modify(modify.renderTemplate(…)); ' +
                'modifyElement() callbacks cannot add media files',
            );
          }
          return loadImage(value);
        });
        if (!image || typeof image !== 'object' || !('content' in image)) break;
        const pending = this.pendingImages.get(body) ?? [];
        pending.push(image as ReturnType<typeof loadImage>);
        this.pendingImages.set(body, pending);
        break;
      }
    }
  }

  private imageTarget(): Promise<ImageTarget | undefined> {
    this.imageTargetPromise ??= this.context.images
      ? this.context.images()
      : Promise.resolve(undefined);
    return this.imageTargetPromise;
  }

  /** Pictures go where the shape is; a shape left without text is removed. */
  private async placePendingImages(body: XmlElement): Promise<void> {
    const images = this.pendingImages.get(body);
    if (!images?.length) return;
    this.pendingImages.delete(body);
    const shape = body.parentNode as XmlElement;
    const text = XmlHelper.collectionToArray(body.getElementsByTagName('a:t'))
      .map((t) => t.textContent)
      .join('');
    try {
      await placeImages(
        await this.imageTarget(),
        shape,
        images,
        text.trim() === '',
      );
    } catch (error) {
      throw this.tagged(error, 'IMAGE');
    }
  }

  private createParagraph(output: OutputParagraph): XmlElement {
    const paragraph = output.source.cloneNode(false) as XmlElement;
    const pPr = this.childElements(output.source, 'a:pPr')[0];
    const endParaRPr = this.childElements(output.source, 'a:endParaRPr')[0];
    if (pPr) paragraph.appendChild(pPr.cloneNode(true));
    output.nodes.forEach((node) => paragraph.appendChild(node));
    if (endParaRPr) paragraph.appendChild(endParaRPr.cloneNode(true));
    return paragraph;
  }

  /** Clone a run with new text; line breaks become <a:br/>. */
  private createRuns(template: XmlElement, text: string): XmlElement[] {
    if (text === '') return [];
    const doc = template.ownerDocument;
    const rPr = this.childElements(template, 'a:rPr')[0];
    const nodes: XmlElement[] = [];
    const lines =
      this.options.processLineBreaks === false
        ? [text]
        : text.split(/\r\n|\n|\v/); // \v: PowerPoint's soft line break

    lines.forEach((line, index) => {
      if (index > 0) {
        const br = doc.createElement('a:br');
        if (rPr) br.appendChild(rPr.cloneNode(true));
        nodes.push(br as unknown as XmlElement);
      }
      if (line === '') return;
      const run = template.cloneNode(true) as XmlElement;
      const t = run.getElementsByTagName('a:t')[0];
      while (t.firstChild) t.removeChild(t.firstChild);
      t.appendChild(doc.createTextNode(XmlHelper.sanitizeText(line)));
      nodes.push(run);
    });
    return nodes;
  }

  private format(value: unknown): string {
    if (value === null || value === undefined) return '';
    if (Array.isArray(value))
      return value.map((v) => this.format(v)).join(', ');
    return String(value);
  }

  /*
   * HTML and hyperlinks
   */

  private relations(): Promise<XmlElement | undefined> {
    this.relationsPromise ??= this.context.relations
      ? this.context.relations()
      : Promise.resolve(undefined);
    return this.relationsPromise;
  }

  /**
   * Converts HTML with the library's html → MultiText pipeline. The tag's
   * run provides the default size and color.
   */
  private htmlParagraphs(
    body: XmlElement,
    source: XmlElement,
    run: XmlElement,
    html: string,
    relations?: XmlElement,
  ): XmlElement[] {
    const doc = body.ownerDocument;
    const host = doc.createElement('p:sp') as unknown as XmlElement;
    const scratch = body.cloneNode(false) as XmlElement;
    ['a:bodyPr', 'a:lstStyle'].forEach((name) =>
      this.childElements(body, name).forEach((child) =>
        scratch.appendChild(child.cloneNode(true)),
      ),
    );
    const paragraph = source.cloneNode(false) as XmlElement;
    paragraph.appendChild(run.cloneNode(true));
    scratch.appendChild(paragraph);
    host.appendChild(scratch);

    // The html pipeline wants a <body>; template values are usually fragments.
    const document = /<body[\s>]/i.test(html)
      ? html
      : `<html><body>${html}</body></html>`;
    new MultiTextHelper(host, relations).run(
      new HtmlToMultiTextHelper().run(document),
    );
    return this.childElements(scratch, 'a:p');
  }

  private toLink(value: unknown): { link: Link; label: string } {
    const spec =
      typeof value === 'object'
        ? (value as { url?: unknown; label?: unknown })
        : { url: value };
    const url = spec.url;
    if (url === null || url === undefined || url === '') {
      throw new Error('LINK needs a url, e.g. {{LINK ({ url, label })}}');
    }
    const internal = typeof url === 'number';
    const target = internal ? `../slides/slide${url}.xml` : String(url);
    const label = this.format(spec.label ?? url);
    return { link: { target, internal }, label };
  }

  private linkRelation(relations: XmlElement, link: Link): string {
    const key = `${link.internal}:${link.target}`;
    const known = this.links.get(key);
    if (known) return known;

    const root =
      relations.tagName === 'Relationships'
        ? relations
        : (relations.ownerDocument.documentElement as unknown as XmlElement);
    const id = `rId${XmlRelationshipHelper.getMaxId(
      root.getElementsByTagName('Relationship'),
      'Id',
      true,
    )}`;
    const rel = root.ownerDocument.createElement('Relationship');
    rel.setAttribute('Id', id);
    rel.setAttribute('Type', link.internal ? REL_TYPE_SLIDE : REL_TYPE_HYPERLINK);
    rel.setAttribute('Target', link.target);
    if (!link.internal) rel.setAttribute('TargetMode', 'External');
    root.appendChild(rel);

    this.links.set(key, id);
    return id;
  }

  private setHyperlink(run: XmlElement, relId: string, internal: boolean): void {
    const doc = run.ownerDocument;
    let rPr = this.childElements(run, 'a:rPr')[0];
    if (!rPr) {
      rPr = doc.createElement('a:rPr') as unknown as XmlElement;
      rPr.setAttribute('lang', 'en-US');
      run.insertBefore(rPr, run.firstChild);
    }
    this.childElements(rPr, 'a:hlinkClick').forEach((h) => rPr.removeChild(h));
    const hlink = doc.createElement('a:hlinkClick') as unknown as XmlElement;
    hlink.setAttribute('r:id', relId);
    if (internal) hlink.setAttribute('action', 'ppaction://hlinksldjump');
    XmlHelper.insertInSchemaOrder(rPr, hlink, RPR_CHILD_ORDER);
  }

  /*
   * Tables: blocks left open inside a cell span (and repeat) table rows.
   */

  private async renderTable(table: XmlElement): Promise<void> {
    if (!this.containsTag(table)) return;
    const rows = this.childElements(table, 'a:tr');
    const rowTokens: RowToken[] = [];

    rows.forEach((row) => {
      const openers: RowToken[] = [];
      const closers: RowToken[] = [];
      let hasContent = false;

      const cells = this.childElements(row, 'a:tc').map((cell) => {
        const body = this.childElements(cell, 'a:txBody')[0];
        if (!body) return undefined;
        const bodyTokens = this.tokenizeBody(body);
        const rowLevel = this.extractRowLevelTags(bodyTokens.tokens);
        rowLevel.openers.forEach((t) =>
          openers.push({ tag: t.tag, directive: t.directive }),
        );
        rowLevel.closers.forEach((t) =>
          closers.push({ tag: t.tag, directive: t.directive }),
        );
        const extracted = new Set<TextToken>([
          ...rowLevel.openers,
          ...rowLevel.closers,
        ]);
        bodyTokens.tokens = bodyTokens.tokens.filter((t) => !extracted.has(t));
        hasContent =
          hasContent ||
          bodyTokens.tokens.some(
            (t) =>
              (t.kind === 'tag' && CONTENT_TYPES.includes(t.directive.type)) ||
              (t.kind === 'text' && t.text.trim() !== '') ||
              (t.kind === 'node' && t.node.tagName === 'a:fld'),
          );
        return bodyTokens;
      });

      const isControlRow = !hasContent && (openers.length || closers.length);
      rowTokens.push(...openers);
      if (!isControlRow) rowTokens.push({ row, cells });
      rowTokens.push(...closers);
    });

    const output: XmlElement[] = [];
    await this.evaluate(
      this.parse(rowTokens, (t) =>
        t.directive ? { tag: t.tag, directive: t.directive } : undefined,
      ),
      this.root,
      async (token, scope) => {
        const row = token.row.cloneNode(true) as XmlElement;
        const cells = this.childElements(row, 'a:tc');
        for (let c = 0; c < cells.length; c++) {
          const bodyTokens = token.cells[c];
          const body = this.childElements(cells[c], 'a:txBody')[0];
          if (!bodyTokens || !body) continue;
          await this.writeBody(
            body,
            bodyTokens,
            this.parse(bodyTokens.tokens, textControl),
            scope,
          );
        }
        output.push(row);
      },
    );

    const anchor = rows[rows.length - 1].nextSibling;
    rows.forEach((row) => table.removeChild(row));
    output.forEach((row, r) => {
      table.insertBefore(row, anchor);
      const rowId = row.getElementsByTagName('a16:rowId')[0];
      if (rowId) rowId.setAttribute('val', String(10000 + r));
    });
    this.updateFrameHeight(table, output);
  }

  /** Block tags that are not closed (or opened) within the same cell. */
  private extractRowLevelTags(tokens: TextToken[]) {
    const stack: TagToken[] = [];
    const closers: TagToken[] = [];

    tokens.forEach((token) => {
      if (token.kind !== 'tag') return;
      const type = token.directive.type;
      if (type === 'if' || type === 'for') {
        stack.push(token);
      } else if (type === 'endif' || type === 'endfor') {
        if (stack.length) stack.pop();
        else closers.push(token);
      } else if ((type === 'else' || type === 'elseif') && !stack.length) {
        throw new TemplateRenderError(
          `${token.tag} cannot span table rows; use it inside a cell`,
        );
      }
    });

    return { openers: stack, closers };
  }

  private updateFrameHeight(table: XmlElement, rows: XmlElement[]): void {
    let frame = table.parentNode as XmlElement;
    while (frame && frame.tagName !== 'p:graphicFrame') {
      frame = frame.parentNode as XmlElement;
    }
    const ext = frame
      ?.getElementsByTagName('p:xfrm')[0]
      ?.getElementsByTagName('a:ext')[0];
    if (!ext) return;
    const height = rows.reduce(
      (sum, row) => sum + Number(row.getAttribute('h') || 0),
      0,
    );
    if (height > 0) ext.setAttribute('cy', String(height));
  }

  /*
   * DOM utilities
   */

  private childElements(parent: XmlElement, tagName?: string): XmlElement[] {
    const children: XmlElement[] = [];
    for (let node = parent.firstChild; node; node = node.nextSibling) {
      if (node.nodeType !== 1) continue;
      const element = node as XmlElement;
      if (!tagName || element.tagName === tagName) children.push(element);
    }
    return children;
  }

  private runText(run: XmlElement): string {
    return run.getElementsByTagName('a:t')[0]?.textContent ?? '';
  }
}
