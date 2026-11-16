import { XMLSerializer } from '@xmldom/xmldom';
import { XmlElement } from '../../types/xml-types';
import { log } from '../../helper/logger';
import { FontRegistry } from './fonts';
import { measureParagraph, ParagraphMetrics, stackHeight } from './layout';
import { BoxGeometry, ShapeStyles, SlideStyles } from './styles';

/** What the template renderer provides to `flow()`. */
export type FlowHost = {
  /**
   * Renders a detached copy of a text body in which `flow()` returns
   * `entries`, without side effects on the slide; returns its paragraphs.
   */
  renderCopy(body: XmlElement, entries: unknown[]): Promise<XmlElement[]>;
  /** Layout, master, theme and presentation defaults of the slide. */
  styles(): Promise<SlideStyles>;
};

export type FlowOptions = {
  /** Font files or folders, searched before the system fonts. */
  fonts?: string[];
  /** Entries that fit nowhere: log a warning (default) or throw. */
  flowOverflow?: 'warn' | 'error';
  /**
   * 0-based: which of the slides the list needs is rendered. Without it,
   * the first one, and entries that need more slides count as overflow.
   */
  flowPage?: number;
  /** Called with the number of slides the lists of the slide need. */
  onFlowPages?: (pages: number) => void;
};

type Box = {
  body: XmlElement;
  name: string;
  geometry: BoxGeometry;
  /** Paragraphs that are not part of the loop, before and after it. */
  head: ParagraphMetrics[];
  tail: ParagraphMetrics[];
  /** Paragraphs of each entry. */
  entries: ParagraphMetrics[][];
};

type Plan = {
  list: unknown[];
  /** Per slide (page): the entries of each box. */
  pages: Map<XmlElement, number[]>[];
  /** Per page: the index of the first entry after it. */
  ends: number[];
};

const TOLERANCE = 0.01;

/** A `flow(…)` call: its argument as written, and without whitespace. */
export type FlowCall = {
  source: string;
  key: string;
  /** Source of the second argument: at most this many entries per box. */
  limit?: string;
};

/** Calls of `flow(…)` (or `flowRest(…)`) in a text. */
export const flowArguments = (text: string, name = 'flow'): FlowCall[] => {
  const args: FlowCall[] = [];
  const pattern = new RegExp(`(^|[^\\w$.])${name}\\s*\\(`, 'g');
  while (pattern.exec(text)) {
    let depth = 1;
    let i = pattern.lastIndex;
    const start = i;
    let quote = '';
    const commas: number[] = [];
    for (; i < text.length && depth > 0; i++) {
      const c = text[i];
      if (quote) {
        if (c === '\\') i++;
        else if (c === quote) quote = '';
      } else if (c === "'" || c === '"' || c === '`') quote = c;
      else if (c === '(' || c === '[' || c === '{') depth++;
      else if (c === ')' || c === ']' || c === '}') depth--;
      else if (c === ',' && depth === 1) commas.push(i);
    }
    const end = i - 1;
    const source = text.slice(start, commas[0] ?? end).trim();
    const limit =
      commas.length > 0
        ? text.slice(commas[0] + 1, commas[1] ?? end).trim()
        : '';
    args.push({
      source,
      key: source.replace(/\s+/g, ''),
      ...(limit ? { limit } : {}),
    });
  }
  return args;
};

const shapeName = (body: XmlElement): string => {
  const shape = body.parentNode as XmlElement;
  const cNvPr = shape?.getElementsByTagName('p:cNvPr')[0];
  return cNvPr?.getAttribute('name') || 'text box';
};

const textOf = (paragraph: XmlElement): string =>
  Array.from(
    { length: paragraph.getElementsByTagName('a:t').length },
    (_, i) => paragraph.getElementsByTagName('a:t')[i].textContent,
  ).join('');

/**
 * Distributes the entries of a list over the text boxes that call
 * `flow(list)`: the boxes are filled in reading order (left to right, then
 * top to bottom), each entry goes as a whole into the first box that still
 * has room for it. Heights are measured with the fonts' real metrics.
 */
export class TextFlow {
  /** Flow boxes (their p:txBody) and the source of their flow() argument. */
  private readonly keys = new Map<XmlElement, string>();
  private readonly plans = new Map<string, Promise<Plan>>();
  private readonly resolved = new Map<string, Plan>();
  private readonly fonts: FontRegistry;
  private readonly serializer = new XMLSerializer();

  constructor(
    boxes: Map<XmlElement, string>,
    private readonly host: FlowHost,
    private readonly options: FlowOptions = {},
  ) {
    boxes.forEach((key, body) => this.keys.set(body, key));
    this.fonts = new FontRegistry(options.fonts ?? []);
  }

  /** The sources of the flow() lists, one per group of boxes. */
  groups(): string[] {
    return [...new Set(this.keys.values())];
  }

  /**
   * Computes where the entries of `list` go, for the boxes of `flow(key)`.
   * A `flowRest()` list that is no group's falls back to the only group.
   * Must be awaited before `entries()`/`rest()`, which are synchronous.
   */
  async prepare(
    key: string,
    list: unknown,
    limits = new Map<XmlElement, number>(),
  ): Promise<void> {
    const groups = this.groups();
    const group = groups.includes(key)
      ? key
      : groups.length === 1
        ? groups[0]
        : undefined;
    if (group === undefined) {
      throw new Error(`flowRest(${key}): no flow(${key}) on the slide`);
    }
    if (!this.plans.has(group)) {
      const entries = this.toList(list);
      this.plans.set(
        group,
        this.computePlan(group, entries, limits).then((plan) => {
          this.resolved.set(group, plan);
          this.options.onFlowPages?.(this.pages);
          return plan;
        }),
      );
    }
    await this.plans.get(group);
  }

  /** The entries of `list` that go into the text box `body`. */
  entries(body: XmlElement | undefined, list: unknown): unknown[] {
    const key = body && this.keys.get(body);
    if (key === undefined) {
      throw new Error(
        'flow() works only in the loop of a text box, not in a table',
      );
    }
    const plan = this.resolved.get(key);
    if (!plan) throw new Error(`flow(${key}) was not prepared`);
    const entries = this.toList(list);
    const page = plan.pages[this.page] ?? new Map<XmlElement, number[]>();
    return (page.get(body) ?? []).map((i) =>
      i < entries.length ? entries[i] : plan.list[i],
    );
  }

  /**
   * The entries of `list` after the rendered slide: on the next slides, or
   * in none (they fit into no box).
   */
  rest(list: unknown): unknown[] {
    const entries = this.toList(list);
    const plans = [...this.resolved.values()];
    const plan =
      plans.find(
        (p) =>
          p.list.length === entries.length &&
          p.list.every((e, i) => e === entries[i]),
      ) ?? (plans.length === 1 ? plans[0] : undefined);
    if (!plan) {
      throw new Error('flowRest(): no flow() over this list on the slide');
    }
    const end = plan.ends[Math.min(this.page, plan.ends.length - 1)] ?? 0;
    return plan.list
      .map((_, i) => i)
      .slice(end)
      .map((i) => (i < entries.length ? entries[i] : plan.list[i]));
  }

  /** How many slides the lists need (the longest one counts). */
  get pages(): number {
    const counts = [...this.resolved.values()].map((p) => p.pages.length);
    return Math.max(1, ...counts);
  }

  private get page(): number {
    return this.options.flowPage ?? 0;
  }

  private toList(value: unknown): unknown[] {
    if (value === null || value === undefined) return [];
    if (typeof (value as Iterable<unknown>)[Symbol.iterator] !== 'function') {
      throw new Error('flow() needs a list');
    }
    return Array.from(value as Iterable<unknown>);
  }

  private async computePlan(
    key: string,
    list: unknown[],
    limits: Map<XmlElement, number>,
  ): Promise<Plan> {
    const bodies = [...this.keys]
      .filter(([, k]) => k === key)
      .map(([body]) => body);
    const styles = await this.host.styles();
    const boxes: Box[] = [];
    for (const body of bodies) {
      boxes.push(await this.measureBox(body, list, styles));
    }

    // reading order: columns left to right, inside a column top to bottom
    boxes.sort((a, b) =>
      Math.abs(a.geometry.x - b.geometry.x) > 1
        ? a.geometry.x - b.geometry.x
        : a.geometry.y - b.geometry.y,
    );

    const pages: Map<XmlElement, number[]>[] = [];
    const ends: number[] = [];
    let next = 0;
    // Fill the boxes, then the same boxes on the next slide, … until an
    // entry fits into none of them (or all are placed)
    for (;;) {
      const start = next;
      const page = new Map<XmlElement, number[]>();
      for (const box of boxes) {
        const taken: number[] = [];
        const limit = limits.get(box.body) ?? Infinity;
        while (next < list.length && taken.length < limit) {
          const paragraphs = [
            ...box.head,
            ...[...taken, next].flatMap((i) => box.entries[i]),
            ...box.tail,
          ];
          // an entry too large for this box tries the next one
          if (stackHeight(paragraphs) > box.geometry.height + TOLERANCE) break;
          taken.push(next++);
        }
        page.set(box.body, taken);
      }
      // a slide without progress: the next entry fits into no box
      if (next === start && pages.length) break;
      pages.push(page);
      ends.push(next);
      if (next >= list.length || next === start) break;
    }

    const unplaced = list.length - next;
    const shown = this.options.flowPage === undefined ? ends[0] : next;
    const missing = list.length - shown;
    if (missing) {
      const where = `the text boxes ${boxes.map((b) => `"${b.name}"`).join(', ')}`;
      const message = unplaced
        ? `flow(${key}): ${unplaced} of ${list.length} entries fit into none of ${where}.`
        : `flow(${key}): ${missing} of ${list.length} entries do not fit into ${where}; ` +
          `the list needs ${pages.length} slides. flowRest(…) returns the entries ` +
          'that are left, renderDeck adds the slides (see countFlowPages).';
      if (this.options.flowOverflow === 'error') throw new Error(message);
      log.warn(message);
    }
    return { list, pages, ends };
  }

  private async measureBox(
    body: XmlElement,
    list: unknown[],
    slide: SlideStyles,
  ): Promise<Box> {
    const name = shapeName(body);
    const styles = new ShapeStyles(
      body.parentNode as XmlElement,
      slide,
      this.fonts,
    );
    const geometry = styles.geometry();
    if (!geometry) {
      throw new Error(
        `flow(): "${name}" has no position and size of its own ` +
          '(a placeholder that inherits them is not supported)',
      );
    }
    const unsupported = styles.unsupported();
    if (unsupported)
      throw new Error(
        `flow(): "${name}" has ${unsupported}, which flow() cannot measure`,
      );

    const measure = (paragraphs: XmlElement[]) =>
      styles
        .paragraphs(paragraphs)
        .map((p) => measureParagraph(p, geometry.width));
    const xml = (p: XmlElement) => this.serializer.serializeToString(p);

    // Without entries: the paragraphs around the loop. A body left empty
    // keeps one empty paragraph, which is not there once entries are in.
    let fixed = await this.host.renderCopy(body, []);
    if (fixed.length === 1 && !textOf(fixed[0]).trim()) fixed = [];
    const fixedXml = fixed.map(xml);

    const head: XmlElement[] = [];
    const tail: XmlElement[] = [];
    const entries: ParagraphMetrics[][] = [];
    let split: number | undefined;
    for (const entry of list) {
      const rendered = await this.host.renderCopy(body, [entry]);
      const renderedXml = rendered.map(xml);
      // the loop's output sits between a common head and a common tail
      if (split === undefined) {
        for (let k = 0; k <= fixed.length; k++) {
          const headMatches = fixedXml
            .slice(0, k)
            .every((x, i) => renderedXml[i] === x);
          const tailLength = fixed.length - k;
          const tailMatches = fixedXml
            .slice(k)
            .every(
              (x, i) => renderedXml[rendered.length - tailLength + i] === x,
            );
          if (headMatches && tailMatches && rendered.length >= fixed.length) {
            split = k;
            break;
          }
        }
        if (split === undefined) {
          throw new Error(
            `flow(): in "${name}", put {{FOR …}} and {{END-FOR}} into ` +
              'paragraphs of their own, around whole paragraphs',
          );
        }
        head.push(...fixed.slice(0, split));
        tail.push(...fixed.slice(split));
      }
      entries.push(
        measure(
          rendered.slice(split, rendered.length - (fixed.length - split)),
        ),
      );
    }

    return {
      body,
      name,
      geometry,
      head: measure(head),
      tail: measure(tail),
      entries,
    };
  }
}
