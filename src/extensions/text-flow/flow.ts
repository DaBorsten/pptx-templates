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
  /**
   * Boxes with the same key render and measure alike, so they share one
   * measurement; undefined measures the box by itself.
   */
  measureKey?(body: XmlElement): string | undefined;
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
  /** Keeps the distribution for the other copies of the slide. */
  flowCache?: FlowCache;
  /** Called with the distribution of each list (renderDeck: flowContinue). */
  onFlowPlan?: (plan: FlowPlan, list: unknown) => void;
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
  /** false: `{ measure: false }`, entries are counted only. */
  measured: boolean;
};

/** What measuring a box yields, shared by boxes with the same measureKey. */
type Measurement = Pick<Box, 'head' | 'tail' | 'entries'>;

/**
 * Boxes that take the same entries: the flow boxes of one list inside one
 * PowerPoint group (a card), or a single box outside of groups.
 */
type Slot = {
  boxes: Box[];
  /** The p:grpSp of a card. */
  group?: XmlElement;
  /** Top left corner on the slide, in points (reading order). */
  x: number;
  y: number;
};

/**
 * Where the entries of a list go. It names a box by its position among the
 * flow boxes of the slide, so it holds for every copy of the slide.
 */
export type FlowPlan = {
  list: unknown[];
  /** Per slide (page): the entries of each box. */
  pages: Map<number, number[]>[];
  /** Per page: the index of the first entry after it. */
  ends: number[];
  /** Names of the boxes, for messages. */
  names: string[];
};

/**
 * Measuring a list takes most of the time of `flow()`. A cache passed as
 * `flowCache` to every render of one template slide (`countFlowPages` and
 * the `flowPage` copies) lets them share one measurement. A render whose
 * boxes, limits, text around the loop or list differ measures again. Use
 * one cache per template slide and set of data.
 */
export class FlowCache {
  /** @internal */
  readonly plans = new Map<string, FlowPlan>();
}

const isObject = (value: unknown): value is object =>
  (typeof value === 'object' && value !== null) || typeof value === 'function';

/** The list a `flowContinue()` result was taken from. */
const continuedFrom = new WeakMap<object, object>();

/** The list of the data that `list` stands for. */
const originOf = (list: unknown): object | undefined =>
  isObject(list) ? (continuedFrom.get(list) ?? list) : undefined;

const count = (counts: Map<unknown, number>, entry: unknown): void => {
  counts.set(entry, (counts.get(entry) ?? 0) + 1);
};

/**
 * The entries that `flow()` lists of earlier slides of a deck show, for
 * `flowContinue(list)`. `renderDeck` keeps one by itself. Objects are
 * compared with `===`, wherever they were shown. Strings and numbers count
 * only for the same list (the array of the data, or a `flowContinue()` of
 * it); of equal ones, as many are skipped as were shown.
 */
export class FlowShown {
  private readonly objects = new Map<unknown, number>();
  /** Strings, numbers, …: per list of the data. */
  private readonly values = new Map<object | undefined, Map<unknown, number>>();

  /**
   * Marks `entries` as shown; `list` is the list `flow()` got, for the
   * strings and numbers among them.
   */
  add(entries: Iterable<unknown>, list?: unknown): void {
    const origin = originOf(list);
    for (const entry of entries) {
      if (isObject(entry)) {
        count(this.objects, entry);
      } else {
        if (!this.values.has(origin)) this.values.set(origin, new Map());
        count(this.values.get(origin), entry);
      }
    }
  }

  /** The entries of `list` that are not shown yet, in their order. */
  continue(list: unknown): unknown[] {
    if (list === null || list === undefined) return [];
    if (typeof (list as Iterable<unknown>)[Symbol.iterator] !== 'function') {
      throw new Error('flowContinue() needs a list');
    }
    const origin = originOf(list);
    const values = this.values.get(origin);
    const used = new Map<unknown, number>();
    const result = Array.from(list as Iterable<unknown>).filter((entry) => {
      const shown = (isObject(entry) ? this.objects : values)?.get(entry) ?? 0;
      const skipped = used.get(entry) ?? 0;
      if (skipped >= shown) return true;
      used.set(entry, skipped + 1);
      return false;
    });
    if (origin) continuedFrom.set(result, origin);
    return result;
  }

  /** A copy that later `add` calls on this one do not change. */
  clone(): FlowShown {
    const copy = new FlowShown();
    this.objects.forEach((n, entry) => copy.objects.set(entry, n));
    this.values.forEach((counts, origin) =>
      copy.values.set(origin, new Map(counts)),
    );
    return copy;
  }
}

/** The entries a plan puts onto one of its pages (0-based). */
export const flowPlanEntries = (plan: FlowPlan, page: number): unknown[] => {
  const indexes = new Set([...(plan.pages[page]?.values() ?? [])].flat());
  return [...indexes].sort((a, b) => a - b).map((i) => plan.list[i]);
};

const sameList = (a: unknown[], b: unknown[]): boolean =>
  a.length === b.length && a.every((entry, i) => entry === b[i]);

const TOLERANCE = 0.01;
const EMU_PER_PT = 12700;

const childOf = (parent: XmlElement | undefined, tagName: string) => {
  for (let n = parent?.firstChild; n; n = n.nextSibling) {
    if ((n as XmlElement).tagName === tagName) return n as XmlElement;
  }
  return undefined;
};

const emuAttr = (el: XmlElement | undefined, name: string): number =>
  Number(el?.getAttribute(name) ?? 0) / EMU_PER_PT;

/** The nearest p:grpSp around a shape, if it is grouped. */
const groupOf = (shape: XmlElement): XmlElement | undefined => {
  const parent = shape?.parentNode as XmlElement;
  return parent?.tagName === 'p:grpSp' ? parent : undefined;
};

/**
 * A point in the coordinates of `shape`'s parent, mapped onto the slide
 * through the transforms of the groups around it.
 */
const toSlide = (
  shape: XmlElement,
  point: { x: number; y: number },
): { x: number; y: number } => {
  let { x, y } = point;
  for (let group = groupOf(shape); group; group = groupOf(group)) {
    const xfrm = childOf(childOf(group, 'p:grpSpPr'), 'a:xfrm');
    if (!xfrm) continue;
    const off = childOf(xfrm, 'a:off');
    const ext = childOf(xfrm, 'a:ext');
    const chOff = childOf(xfrm, 'a:chOff');
    const chExt = childOf(xfrm, 'a:chExt');
    const scale = (e: string) =>
      emuAttr(chExt, e) ? emuAttr(ext, e) / emuAttr(chExt, e) : 1;
    x = emuAttr(off, 'x') + (x - emuAttr(chOff, 'x')) * scale('cx');
    y = emuAttr(off, 'y') + (y - emuAttr(chOff, 'y')) * scale('cy');
  }
  return { x, y };
};

/** A `flow(…)` call: its argument as written, and without whitespace. */
export type FlowCall = {
  source: string;
  key: string;
  /** Source of the second argument: at most this many entries per box. */
  limit?: string;
  /** Source of the third argument, such as `{ measure: false }`. */
  options?: string;
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
    const options =
      commas.length > 1
        ? text.slice(commas[1] + 1, commas[2] ?? end).trim()
        : '';
    args.push({
      source,
      key: source.replace(/\s+/g, ''),
      ...(limit ? { limit } : {}),
      ...(options ? { options } : {}),
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
  /** Position of each flow box among the flow boxes of the slide. */
  private readonly indexes = new Map<XmlElement, number>();
  /** Each box as it is in the template: its shape and the groups around. */
  private readonly fingerprints = new Map<XmlElement, string>();
  private readonly plans = new Map<string, Promise<FlowPlan>>();
  private readonly resolved = new Map<string, FlowPlan>();
  private readonly fonts: FontRegistry;
  private readonly serializer = new XMLSerializer();

  constructor(
    boxes: Map<XmlElement, string>,
    private readonly host: FlowHost,
    private readonly options: FlowOptions = {},
  ) {
    boxes.forEach((key, body) => {
      this.keys.set(body, key);
      this.indexes.set(body, this.indexes.size);
      if (options.flowCache) {
        this.fingerprints.set(body, this.fingerprint(body));
      }
    });
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
    /** Boxes of `flow(list, n, { measure: false })`. */
    unmeasured = new Set<XmlElement>(),
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
        this.plan(group, entries, limits, unmeasured).then((plan) => {
          this.reportOverflow(group, plan);
          this.resolved.set(group, plan);
          this.options.onFlowPlan?.(plan, list);
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
    const page = plan.pages[this.page];
    return (page?.get(this.indexes.get(body)) ?? []).map((i) =>
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

  /**
   * The flow box of the card `shape` is in: a box in the nearest group
   * around `shape` that has one. Its entries are the card's entries.
   */
  boxOfCard(shape: XmlElement): XmlElement | undefined {
    for (let group = groupOf(shape); group; group = groupOf(group)) {
      for (const body of this.keys.keys()) {
        if (groupOf(body.parentNode as XmlElement) === group) return body;
      }
    }
    return undefined;
  }

  /**
   * Cards (PowerPoint groups with flow boxes) that stay empty on the
   * rendered slide, such as the last columns of the last slide.
   */
  emptyGroups(): XmlElement[] {
    const used = new Map<XmlElement, boolean>();
    this.keys.forEach((key, body) => {
      const group = groupOf(body.parentNode as XmlElement);
      if (!group) return;
      const plan = this.resolved.get(key);
      const page = plan?.pages[this.page];
      const filled = !plan || !!page?.get(this.indexes.get(body))?.length;
      used.set(group, (used.get(group) ?? false) || filled);
    });
    return [...used].filter(([, filled]) => !filled).map(([group]) => group);
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

  private bodiesOf(key: string): XmlElement[] {
    return [...this.keys].filter(([, k]) => k === key).map(([body]) => body);
  }

  /** The shape of a box and the transforms of the groups around it. */
  private fingerprint(body: XmlElement): string {
    const shape = body.parentNode as XmlElement;
    const parts = [this.serializer.serializeToString(shape)];
    for (let group = groupOf(shape); group; group = groupOf(group)) {
      const properties = childOf(group, 'p:grpSpPr');
      parts.push(
        properties ? this.serializer.serializeToString(properties) : '',
      );
    }
    return parts.join('\n');
  }

  /**
   * The plan of a list: from the cache if another copy of the slide has
   * distributed the same entries over the same boxes, else computed.
   */
  private async plan(
    key: string,
    list: unknown[],
    limits: Map<XmlElement, number>,
    unmeasured: Set<XmlElement>,
  ): Promise<FlowPlan> {
    const cache = this.options.flowCache;
    if (!cache) return this.computePlan(key, list, limits, unmeasured);

    // Everything the distribution depends on, besides the entries: the
    // boxes, their limits and the paragraphs around the loop, which can
    // differ from copy to copy ({{IF flowPage > 1}}…)
    const parts = [key];
    for (const body of this.bodiesOf(key)) {
      const measured = !unmeasured.has(body);
      const fixed = measured ? await this.host.renderCopy(body, []) : [];
      parts.push(
        String(this.indexes.get(body)),
        String(limits.get(body) ?? ''),
        String(measured),
        this.fingerprints.get(body),
        ...fixed.map((p) => this.serializer.serializeToString(p)),
      );
    }
    const signature = parts.join('\n');
    const cached = cache.plans.get(signature);
    if (cached && sameList(cached.list, list)) return cached;
    const plan = await this.computePlan(key, list, limits, unmeasured);
    cache.plans.set(signature, plan);
    return plan;
  }

  private async computePlan(
    key: string,
    list: unknown[],
    limits: Map<XmlElement, number>,
    unmeasured: Set<XmlElement>,
  ): Promise<FlowPlan> {
    const bodies = this.bodiesOf(key);
    const styles = await this.host.styles();
    const boxes: Box[] = [];
    const measurements = new Map<string, Measurement>();
    for (const body of bodies) {
      boxes.push(
        unmeasured.has(body)
          ? this.unmeasuredBox(body, styles, limits)
          : await this.measureBox(body, list, styles, measurements),
      );
    }

    const slots = this.slotsOf(boxes);

    const pages: Map<number, number[]>[] = [];
    const ends: number[] = [];
    let next = 0;
    // Fill the slots, then the same slots on the next slide, … until an
    // entry fits into none of them (or all are placed)
    for (;;) {
      const start = next;
      const page = new Map<number, number[]>();
      for (const slot of slots) {
        const taken: number[] = [];
        const limit = Math.min(
          ...slot.boxes.map((b) => limits.get(b.body) ?? Infinity),
        );
        while (next < list.length && taken.length < limit) {
          // an entry too large for this slot tries the next one; in a
          // card, its part has to fit into every box
          const fits = slot.boxes.every(
            (box) =>
              !box.measured ||
              stackHeight([
                ...box.head,
                ...[...taken, next].flatMap((i) => box.entries[i]),
                ...box.tail,
              ]) <=
                box.geometry.height + TOLERANCE,
          );
          if (!fits) break;
          taken.push(next++);
        }
        slot.boxes.forEach((box) =>
          page.set(this.indexes.get(box.body), taken),
        );
      }
      // a slide without progress: the next entry fits into no box
      if (next === start && pages.length) break;
      pages.push(page);
      ends.push(next);
      if (next >= list.length || next === start) break;
    }

    return { list, pages, ends, names: boxes.map((b) => b.name) };
  }

  /** Entries this render leaves out: a warning, or an error. */
  private reportOverflow(key: string, plan: FlowPlan): void {
    const { list, pages, ends } = plan;
    const placed = ends[ends.length - 1];
    const unplaced = list.length - placed;
    const shown = this.options.flowPage === undefined ? ends[0] : placed;
    const missing = list.length - shown;
    if (!missing) return;
    const where = `the text boxes ${plan.names.map((n) => `"${n}"`).join(', ')}`;
    const message = unplaced
      ? `flow(${key}): ${unplaced} of ${list.length} entries fit into none of ${where}.`
      : `flow(${key}): ${missing} of ${list.length} entries do not fit into ${where}; ` +
        `the list needs ${pages.length} slides. flowRest(…) returns the entries ` +
        'that are left, renderDeck adds the slides (see countFlowPages).';
    if (this.options.flowOverflow === 'error') throw new Error(message);
    log.warn(message);
  }

  /**
   * The slots in reading order: left to right, of slots on top of each
   * other the upper one first. Boxes in one group form one slot.
   */
  private slotsOf(boxes: Box[]): Slot[] {
    const slots = new Map<XmlElement, Slot>();
    for (const box of boxes) {
      const shape = box.body.parentNode as XmlElement;
      const group = groupOf(shape);
      const { x, y } = toSlide(shape, box.geometry);
      const slot = slots.get(group ?? box.body);
      if (slot) {
        slot.boxes.push(box);
        slot.x = Math.min(slot.x, x);
        slot.y = Math.min(slot.y, y);
      } else {
        slots.set(group ?? box.body, { boxes: [box], group, x, y });
      }
    }
    return [...slots.values()].sort((a, b) =>
      Math.abs(a.x - b.x) > 1 ? a.x - b.x : a.y - b.y,
    );
  }

  /** A box of `{ measure: false }`: only its position counts. */
  private unmeasuredBox(
    body: XmlElement,
    slide: SlideStyles,
    limits: Map<XmlElement, number>,
  ): Box {
    const name = shapeName(body);
    if (!limits.has(body)) {
      throw new Error(
        `flow(): "${name}" has { measure: false } without a number of ` +
          'entries per box, such as flow(projects, 1, { measure: false })',
      );
    }
    const geometry = new ShapeStyles(
      body.parentNode as XmlElement,
      slide,
      this.fonts,
    ).geometry();
    return {
      body,
      name,
      geometry: geometry ?? { x: 0, y: 0, width: 0, height: 0 },
      head: [],
      tail: [],
      entries: [],
      measured: false,
    };
  }

  private async measureBox(
    body: XmlElement,
    list: unknown[],
    slide: SlideStyles,
    /** Measurements of this list so far, by measureKey. */
    measurements: Map<string, Measurement>,
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

    const key = this.host.measureKey?.(body);
    const known = key === undefined ? undefined : measurements.get(key);
    if (known) return { body, name, geometry, ...known, measured: true };

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
      // the loop's output sits between a common head and a common tail
      if (split === undefined) {
        const renderedXml = rendered.map(xml);
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

    const measurement: Measurement = {
      head: measure(head),
      tail: measure(tail),
      entries,
    };
    if (key !== undefined) measurements.set(key, measurement);
    return { body, name, geometry, ...measurement, measured: true };
  }
}
