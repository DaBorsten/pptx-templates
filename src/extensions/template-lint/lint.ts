import fs from 'fs';
import JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';
import { XmlDocument, XmlElement } from '../../types/xml-types';
import {
  fixSmartQuotes,
  isTableMarker,
  parseCommand,
  resolveAlias,
  tagSyntax,
} from '../template-logic/commands';
import {
  parseStatements,
  SourceRange,
  TemplateExpression,
  TemplateExpressionError,
} from '../template-logic/expression';
import { DataModel, Known, suggest } from './data-model';
import {
  LintTemplateOptions,
  TemplateLintIssue,
  TemplateLintResult,
} from './types';
import { autocorrectHint } from './hints';

/** One stretch of runs: a tag must sit within it (renderer tokenizing). */
type Segment = { text: string; paragraph: XmlElement; cell?: number };
type Container = { kind: 'shape' | 'table'; segments: Segment[] };
type OpenBlock = {
  type: 'IF' | 'FOR';
  tag: string;
  cell?: number;
  shape?: string;
  hasElse: boolean;
  vars: [string, Known][];
};

type IssueInput = Omit<TemplateLintIssue, 'part' | 'slide'>;

const ANY: Known = { kind: 'any' };

const children = (
  parent: XmlElement | XmlDocument,
  tagName?: string,
): XmlElement[] => {
  const result: XmlElement[] = [];
  for (let node = parent?.firstChild; node; node = node.nextSibling) {
    if (
      node.nodeType === 1 &&
      (!tagName || (node as XmlElement).tagName === tagName)
    ) {
      result.push(node as XmlElement);
    }
  }
  return result;
};

const descendants = (
  parent: XmlElement | XmlDocument,
  tagName: string,
): XmlElement[] =>
  Array.from(
    parent.getElementsByTagName(tagName) as unknown as ArrayLike<XmlElement>,
  );

const parseXml = (xml: string): XmlDocument =>
  new DOMParser().parseFromString(
    xml,
    'application/xml',
  ) as unknown as XmlDocument;

const shapeOf = (node: XmlElement): XmlElement | undefined => {
  for (
    let n = node.parentNode as XmlElement;
    n;
    n = n.parentNode as XmlElement
  ) {
    if (n.tagName === 'p:sp' || n.tagName === 'p:graphicFrame') return n;
  }
  return undefined;
};

const shapeName = (shape?: XmlElement): string | undefined =>
  shape && descendants(shape, 'p:cNvPr')[0]?.getAttribute('name');

/** Same rule as renderTemplate's IMAGE placement (`shapeFrame` in images.ts). */
const hasOwnFrame = (shape?: XmlElement): boolean => {
  const xfrm = children(children(shape, 'p:spPr')[0], 'a:xfrm')[0];
  return !!children(xfrm, 'a:off')[0] && !!children(xfrm, 'a:ext')[0];
};

/** Paragraph → stretches of consecutive runs, as the renderer tags them. */
const segmentsOf = (paragraph: XmlElement, cell?: number): Segment[] => {
  const segments: Segment[] = [];
  let text = '';
  const flush = () => {
    if (text) segments.push({ text, paragraph, cell });
    text = '';
  };
  children(paragraph).forEach((child) => {
    if (child.tagName === 'a:r') {
      text += descendants(child, 'a:t')
        .map((t) => t.textContent)
        .join('');
    } else if (child.tagName !== 'a:pPr' && child.tagName !== 'a:endParaRPr') {
      flush();
    }
  });
  flush();
  return segments;
};

/** Shape text bodies and tables in document order (renderer `findTargets`). */
const containersOf = (doc: XmlDocument): Container[] => {
  const containers: Container[] = [];
  const visit = (parent: XmlElement | XmlDocument) => {
    children(parent).forEach((child) => {
      if (child.tagName === 'p:txBody') {
        containers.push({
          kind: 'shape',
          segments: children(child, 'a:p').flatMap((p) => segmentsOf(p)),
        });
      } else if (child.tagName === 'a:tbl') {
        let cell = 0;
        const segments = children(child, 'a:tr').flatMap((row) =>
          children(row, 'a:tc').flatMap((tc) => {
            cell++;
            return children(tc, 'a:txBody').flatMap((body) =>
              children(body, 'a:p').flatMap((p) => segmentsOf(p, cell)),
            );
          }),
        );
        containers.push({ kind: 'table', segments });
      } else {
        visit(child);
      }
    });
  };
  visit(doc);
  return containers;
};

const resolvePath = (base: string, target: string): string => {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  target.split('/').forEach((segment) => {
    if (segment === '..') parts.pop();
    else if (segment !== '.') parts.push(segment);
  });
  return parts.join('/');
};

const OTHER_PARTS: {
  pattern: RegExp;
  describe: (n: string) => string;
  hint: string;
}[] = [
  {
    pattern: /^ppt\/slideLayouts\/slideLayout(\d+)\.xml$/,
    describe: (n) => `slide layout ${n}`,
    hint:
      'renderTemplate renders the slide XML only. Text of the layout is not ' +
      'part of it and stays as it is. Put the tag on the slide itself.',
  },
  {
    pattern: /^ppt\/slideMasters\/slideMaster(\d+)\.xml$/,
    describe: (n) => `slide master ${n}`,
    hint:
      'renderTemplate renders the slide XML only. Text of the master is not ' +
      'part of it and stays as it is. Put the tag on the slide itself.',
  },
  {
    pattern: /^ppt\/notesSlides\/notesSlide(\d+)\.xml$/,
    describe: () => 'speaker notes',
    hint: 'renderTemplate does not render speaker notes.',
  },
  {
    pattern: /^ppt\/charts\/chart(\d+)\.xml$/,
    describe: (n) => `chart ${n}`,
    hint:
      'Chart titles, labels and data are not rendered. Fill charts with ' +
      'modify.setChartData and friends.',
  },
  {
    pattern: /^ppt\/diagrams\/data(\d+)\.xml$/,
    describe: (n) => `SmartArt ${n}`,
    hint: 'SmartArt text is not rendered.',
  },
];

class TemplateLinter {
  readonly issues: TemplateLintIssue[] = [];
  private readonly syntax: ReturnType<typeof tagSyntax>;
  private readonly model: DataModel;
  private readonly seen = new Set<string>();
  private readonly fixQuotes: boolean;

  constructor(options: LintTemplateOptions) {
    this.syntax = tagSyntax(options);
    this.fixQuotes = options.fixSmartQuotes ?? false;
    this.model = new DataModel(
      options.data,
      options.schema,
      options.additionalJsContext,
    );
  }

  private report(
    part: string,
    slide: number | undefined,
    issue: IssueInput,
  ): void {
    // The same finding in the same shape once is enough (e.g. in loops)
    const key = [part, issue.shape, issue.rule, issue.message].join('\u0000');
    if (this.seen.has(key)) return;
    this.seen.add(key);
    this.issues.push({
      severity: issue.severity,
      rule: issue.rule,
      message: issue.message,
      ...(issue.hint && { hint: issue.hint }),
      part,
      ...(slide !== undefined && { slide }),
      ...(issue.shape !== undefined && { shape: issue.shape }),
      ...(issue.tag !== undefined && { tag: issue.tag }),
      ...(issue.range && { range: issue.range }),
    });
  }

  /** The current tag as the renderer sees it (same length as the raw tag). */
  private tagText = '';
  /** Where the command's code starts in `tagText`; -1 if not in it (alias). */
  private codeStart = -1;

  private enterTag(raw: string): void {
    // Non-breaking spaces count as spaces, as in parseCommand
    const text = raw.split(String.fromCharCode(0xa0)).join(' ');
    this.tagText = this.fixQuotes ? fixSmartQuotes(text) : text;
    this.codeStart = -1;
  }

  private enterCode(code: string): void {
    this.codeStart = code ? this.tagText.lastIndexOf(code) : -1;
  }

  /** A range in an expression's source → the same characters in the tag. */
  private locate(source: string, range?: SourceRange): SourceRange | undefined {
    if (!range || this.codeStart < 0) return undefined;
    const offset = this.tagText.indexOf(source, this.codeStart);
    if (offset < 0) return undefined;
    return {
      start: offset + range.start,
      end: offset + Math.max(range.end, range.start + 1),
    };
  }

  /** The renderer's steps: resolve an alias, fix quotes if asked, classify. */
  private parse(inner: string, aliases: Map<string, string>, raw: string) {
    const resolved = resolveAlias(inner, aliases, raw);
    return parseCommand(
      this.fixQuotes ? fixSmartQuotes(resolved) : resolved,
      raw,
    );
  }

  /** Template tags in `text`; setTable's `{{each:…}}` markers are not. */
  private tagsIn(text: string): { raw: string; inner: string }[] {
    return Array.from(text.matchAll(this.syntax.pattern), (m) => ({
      raw: m[0],
      inner: m[1],
    })).filter((tag) => !isTableMarker(tag.inner));
  }

  lintSlide(part: string, slide: number, doc: XmlDocument): void {
    const report = (issue: IssueInput) => this.report(part, slide, issue);
    const containers = containersOf(doc);
    const allTags = containers.flatMap((c) =>
      c.segments.flatMap((s) => this.tagsIn(s.text)),
    );

    // ALIAS and EXEC names apply to the whole render call: collect up front
    const aliases = new Map<string, string>();
    const execNames = new Set<string>();
    allTags.forEach(({ raw, inner }) => {
      try {
        const command = this.parse(inner, aliases, raw);
        if (command.type === 'ALIAS') aliases.set(command.name, command.code);
        if (command.type === 'EXEC') {
          parseStatements(command.code).forEach(
            (s) => s.assign && execNames.add(s.assign),
          );
        }
      } catch {
        // reported below, at the tag
      }
    });

    containers.forEach((container) => {
      const stack: OpenBlock[] = [];
      const where = container.kind === 'table' ? 'table' : 'shape';

      container.segments.forEach((segment) => {
        const shape = shapeOf(segment.paragraph);
        const name = shapeName(shape);
        this.checkDelimiters(segment.text, (issue) =>
          report({ ...issue, shape: name }),
        );

        for (const { raw, inner } of this.tagsIn(segment.text)) {
          // Without a more precise range, the whole tag is marked
          const tagReport = (issue: Omit<IssueInput, 'shape' | 'tag'>) =>
            report({
              ...issue,
              shape: name,
              tag: raw,
              range: issue.range ?? { start: 0, end: raw.length },
            });
          this.enterTag(raw);
          let command: ReturnType<typeof parseCommand>;
          try {
            command = this.parse(inner, aliases, raw);
            this.enterCode(command.code);
          } catch (error) {
            const unknownAlias = /unknown alias/.test((error as Error).message);
            tagReport({
              severity: 'error',
              rule: unknownAlias ? 'unknown-alias' : 'syntax',
              message: (error as Error).message,
              hint: unknownAlias
                ? 'Define it on the same slide with {{ALIAS name command}}.'
                : undefined,
            });
            continue;
          }

          const loopVars = new Map(stack.flatMap((block) => block.vars));
          const check = (code: string) =>
            this.checkExpression(code, loopVars, execNames, tagReport);

          switch (command.type) {
            case 'IF': {
              check(command.code);
              stack.push({
                type: 'IF',
                tag: raw,
                cell: segment.cell,
                shape: name,
                hasElse: false,
                vars: [],
              });
              break;
            }
            case 'FOR': {
              const expr = check(command.code);
              const element = this.loopElement(expr, loopVars, execNames);
              if (!element) {
                tagReport({
                  severity: 'warning',
                  rule: 'not-a-list',
                  message: `"${command.code}" is not a list in the example data`,
                  range: this.locate(command.code, {
                    start: 0,
                    end: command.code.length,
                  }),
                });
              }
              const known = element ?? ANY;
              stack.push({
                type: 'FOR',
                tag: raw,
                cell: segment.cell,
                shape: name,
                hasElse: false,
                vars: [
                  [command.name, known],
                  [`$${command.name}`, known],
                  ['$idx', { kind: 'value', value: 0 }],
                ],
              });
              break;
            }
            case 'ELSE-IF':
            case 'ELSE': {
              if (command.type === 'ELSE-IF') check(command.code);
              const top = stack[stack.length - 1];
              if (top?.type !== 'IF') {
                tagReport({
                  severity: 'error',
                  rule: 'unmatched-tag',
                  message: top
                    ? `${raw} inside ${top.tag}, but without a matching {{IF …}}`
                    : `${raw} without a matching {{IF …}}`,
                });
              } else if (
                container.kind === 'table' &&
                top.cell !== segment.cell
              ) {
                tagReport({
                  severity: 'error',
                  rule: 'unsupported-location',
                  message: `${raw} cannot span table rows; use it inside a cell`,
                  hint: `Put ${top.tag}, ${raw} and {{END-IF}} into the same cell.`,
                });
              } else if (top.hasElse) {
                tagReport({
                  severity: 'error',
                  rule: 'unmatched-tag',
                  message: `${raw} after {{ELSE}} in ${top.tag}`,
                });
              } else if (command.type === 'ELSE') {
                top.hasElse = true;
              }
              break;
            }
            case 'END-IF':
            case 'END-FOR': {
              this.close(
                stack,
                command.type === 'END-IF' ? 'IF' : 'FOR',
                raw,
                where,
                tagReport,
              );
              break;
            }
            case 'EXEC': {
              try {
                parseStatements(command.code).forEach((statement) =>
                  this.checkReferences(
                    statement.expr,
                    loopVars,
                    execNames,
                    tagReport,
                  ),
                );
              } catch (error) {
                this.syntaxError(command.code, error, tagReport);
              }
              break;
            }
            case 'ALIAS':
              break;
            case 'IMAGE': {
              check(command.code);
              if (segment.cell !== undefined) {
                tagReport({
                  severity: 'error',
                  rule: 'unsupported-location',
                  message: 'IMAGE is not supported inside table cells',
                  hint: 'Put the tag into a text box or shape on top of the table.',
                });
              } else if (!hasOwnFrame(shape)) {
                tagReport({
                  severity: 'error',
                  rule: 'unsupported-location',
                  message:
                    'IMAGE needs a shape with its own position and size; ' +
                    'this one inherits them from the layout',
                  hint:
                    'Move or resize the placeholder once in PowerPoint (that ' +
                    'stores its own position), or use a text box.',
                });
              }
              break;
            }
            default:
              check(command.code);
          }
        }
      });

      stack.forEach((block) =>
        report({
          severity: 'error',
          rule: 'unclosed-block',
          message: `${block.tag} is not closed with {{END-${block.type}}} in this ${where}`,
          hint:
            where === 'table'
              ? `Add {{END-${block.type}}} later in the same table.`
              : `Add {{END-${block.type}}} in the same shape: a block cannot continue in another shape.`,
          shape: block.shape,
          tag: block.tag,
          range: { start: 0, end: block.tag.length },
        }),
      );
    });
  }

  /** Pops the innermost block of `type`; reports what does not fit. */
  private close(
    stack: OpenBlock[],
    type: 'IF' | 'FOR',
    raw: string,
    where: string,
    report: (issue: Omit<IssueInput, 'shape' | 'tag'>) => void,
  ): void {
    const index = stack.map((b) => b.type).lastIndexOf(type);
    if (index === -1) {
      const top = stack[stack.length - 1];
      report({
        severity: 'error',
        rule: 'unmatched-tag',
        message:
          `${raw} without a matching {{${type} …}} in this ${where}` +
          (top ? ` (open: ${top.tag})` : ''),
      });
      return;
    }
    stack.splice(index + 1).forEach((inner) =>
      report({
        severity: 'error',
        rule: 'unclosed-block',
        message: `${inner.tag} is not closed before ${raw}`,
        hint: `Add {{END-${inner.type}}} before ${raw}.`,
      }),
    );
    stack.pop();
  }

  /** `{{` without `}}` and vice versa: such text is left as it is. */
  private checkDelimiters(
    text: string,
    report: (issue: IssueInput) => void,
  ): void {
    const { opening, closing } = this.syntax;
    // Complete tags become … so they do not count (and show as … in excerpts)
    const rest = text.replace(this.syntax.pattern, '…');
    const hint =
      'A tag has to be in one line: it cannot span paragraphs or line ' +
      'breaks (Shift+Enter). Check both delimiters.';
    // Some text around the stray delimiter, with the delimiter marked
    const excerpt = (index: number, delimiter: string) => {
      const from = Math.max(0, index - 15);
      const slice = rest.slice(from, index + 25);
      const start = index - from - (slice.length - slice.trimStart().length);
      return {
        tag: slice.trim(),
        range: { start, end: start + delimiter.length },
      };
    };

    const open = rest.indexOf(opening);
    if (open !== -1) {
      report({
        severity: 'error',
        rule: 'unterminated-tag',
        message:
          opening === closing
            ? `Unpaired "${opening}"; the text stays as it is`
            : `"${opening}" without "${closing}"; the text stays as it is`,
        hint,
        ...excerpt(open, opening),
      });
      return;
    }
    const close = rest.indexOf(closing);
    if (close !== -1 && opening !== closing) {
      report({
        severity: 'error',
        rule: 'unterminated-tag',
        message: `"${closing}" without "${opening}"; the text stays as it is`,
        hint,
        ...excerpt(close, closing),
      });
    }
  }

  private checkExpression(
    code: string,
    loopVars: Map<string, Known>,
    execNames: Set<string>,
    report: (issue: Omit<IssueInput, 'shape' | 'tag'>) => void,
  ): TemplateExpression | undefined {
    let expr: TemplateExpression;
    try {
      expr = new TemplateExpression(code);
    } catch (error) {
      this.syntaxError(code, error, report);
      return undefined;
    }
    this.checkReferences(expr, loopVars, execNames, report);
    return expr;
  }

  private syntaxError(
    code: string,
    error: unknown,
    report: (issue: Omit<IssueInput, 'shape' | 'tag'>) => void,
  ): void {
    const message = (error as Error).message;
    const failed = error instanceof TemplateExpressionError ? error : undefined;
    report({
      severity: 'error',
      rule: 'syntax',
      message,
      hint: autocorrectHint(code, message, this.fixQuotes),
      range: failed && this.locate(failed.expression, failed.range),
    });
  }

  private lookup(
    name: string,
    loopVars: Map<string, Known>,
    execNames: Set<string>,
  ) {
    if (loopVars.has(name)) return loopVars.get(name);
    if (execNames.has(name)) return ANY;
    return this.model.root(name);
  }

  private checkReferences(
    expr: TemplateExpression,
    loopVars: Map<string, Known>,
    execNames: Set<string>,
    report: (issue: Omit<IssueInput, 'shape' | 'tag'>) => void,
  ): void {
    // Rejected by the renderer only when evaluated, so check without data too
    for (const { key, range } of expr.forbiddenKeys()) {
      report({
        severity: 'error',
        rule: 'syntax',
        message: `Access to "${key}" is not allowed`,
        hint: 'Keys like constructor, prototype or __proto__ are blocked for safety.',
        range: this.locate(expr.source, range),
      });
    }
    if (!this.model.enabled) return;
    for (const ref of expr.references()) {
      let known = this.lookup(ref.name, loopVars, execNames);
      if (known.kind === 'missing') {
        const guess = suggest(ref.name, known.candidates);
        report({
          severity: 'warning',
          rule: 'unknown-variable',
          message:
            `"${ref.name}" is not in the data` +
            (guess ? ` (did you mean "${guess}"?)` : ''),
          hint:
            'Unknown names evaluate to undefined. Add the name to the data, ' +
            'or define it with {{EXEC …}} or {{FOR … IN …}}.',
          range: this.locate(expr.source, ref.ranges[0]),
        });
        continue;
      }
      for (let i = 0; i < ref.path.length; i++) {
        const next = this.model.member(known, ref.path[i]);
        if (next.kind === 'missing') {
          const owner = [ref.name, ...ref.path.slice(0, i)].join('.');
          const guess = suggest(ref.path[i], next.candidates);
          report({
            severity: 'warning',
            rule: 'unknown-property',
            message:
              `"${ref.path[i]}" does not exist on "${owner}" (${next.owner})` +
              (guess ? `; did you mean "${guess}"?` : ''),
            range: this.locate(expr.source, ref.ranges[i + 1]),
          });
          break;
        }
        known = next;
      }
    }
  }

  /** Entry type of a FOR loop; undefined if the data says "not a list". */
  private loopElement(
    expr: TemplateExpression | undefined,
    loopVars: Map<string, Known>,
    execNames: Set<string>,
  ): Known | undefined {
    const path = expr?.path();
    if (!path || !this.model.enabled) return ANY;
    let known = this.lookup(path.name, loopVars, execNames);
    for (const key of path.path) {
      if (known.kind === 'missing') return ANY;
      known = this.model.member(known, key);
    }
    return known.kind === 'missing' ? ANY : this.model.element(known);
  }

  lintOtherPart(
    part: string,
    slide: number | undefined,
    texts: string[],
  ): void {
    const kind = OTHER_PARTS.find((p) => p.pattern.test(part));
    if (!kind) return;
    const tags = texts.flatMap((text) => this.tagsIn(text).map((t) => t.raw));
    if (!tags.length) return;
    const n = kind.pattern.exec(part)[1];
    const shown = tags.slice(0, 3).join(', ') + (tags.length > 3 ? ', …' : '');
    this.report(part, slide, {
      severity: 'warning',
      rule: 'unsupported-location',
      message: `${tags.length} tag(s) in ${kind.describe(n)} are not rendered: ${shown}`,
      hint: kind.hint,
      tag: tags[0],
    });
  }
}

/**
 * Checks the template tags of a .pptx before rendering: syntax errors,
 * unclosed or unmatched IF/FOR blocks, tags that cannot be rendered where
 * they are (layouts, masters, notes, charts, IMAGE in tables), and - with
 * example `data` or a JSON `schema` - names and property paths that do not
 * exist in the data.
 *
 * @param pptx Path to a .pptx file or its contents.
 */
export const lintTemplate = async (
  pptx: string | Buffer | Uint8Array | ArrayBuffer,
  options: LintTemplateOptions = {},
): Promise<TemplateLintResult> => {
  const content = typeof pptx === 'string' ? fs.readFileSync(pptx) : pptx;
  const zip = await JSZip.loadAsync(content);
  const linter = new TemplateLinter(options);

  const slides = Object.keys(zip.files)
    .map((name) => ({
      name,
      match: /^ppt\/slides\/slide(\d+)\.xml$/.exec(name),
    }))
    .filter((s) => s.match)
    .map((s) => ({ name: s.name, slide: Number(s.match[1]) }))
    .sort((a, b) => a.slide - b.slide);

  // Parts a slide refers to (notes, charts, SmartArt) → its number
  const slideOfPart = new Map<string, number>();
  for (const { name, slide } of slides) {
    const relsPath = name.replace(
      /slides\/(slide\d+\.xml)$/,
      'slides/_rels/$1.rels',
    );
    const rels = await zip.file(relsPath)?.async('string');
    if (!rels) continue;
    descendants(parseXml(rels), 'Relationship').forEach((rel) => {
      if (rel.getAttribute('TargetMode') === 'External') return;
      // A relationship without Target is broken, but not the linter's topic
      if (!rel.getAttribute('Target')) return;
      const target = resolvePath(name, rel.getAttribute('Target'));
      if (!slideOfPart.has(target)) slideOfPart.set(target, slide);
    });
  }

  for (const { name, slide } of slides) {
    linter.lintSlide(
      name,
      slide,
      parseXml(await zip.file(name).async('string')),
    );
  }

  const others = Object.keys(zip.files)
    .filter((name) => OTHER_PARTS.some((p) => p.pattern.test(name)))
    .sort((a, b) => a.localeCompare(b, 'en', { numeric: true }));
  for (const name of others) {
    const doc = parseXml(await zip.file(name).async('string'));
    const texts = [...descendants(doc, 'a:t'), ...descendants(doc, 'c:v')].map(
      (t) => t.textContent,
    );
    linter.lintOtherPart(name, slideOfPart.get(name), texts);
  }

  const issues = linter.issues;
  return {
    issues,
    errorCount: issues.filter((i) => i.severity === 'error').length,
    warningCount: issues.filter((i) => i.severity === 'warning').length,
  };
};
