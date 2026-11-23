import fs from 'fs';
import JSZip from 'jszip';
import ArchiveJszip from '../../helper/archive/archive-jszip';
import { PptPaths } from '../../helper/ppt-paths';
import type { RootPresTemplate } from '../../interfaces/root-pres-template';
import { removeUnreachableParts } from './prune';
import { generatePptx } from './zip-writer';
import { readText } from './zip-text';
import type {
  RenderTemplateOptions,
  TemplateData,
} from '../template-logic/types';
import { FlowCache, FlowPlan, flowPlanEntries, FlowShown } from '../text-flow';
import {
  evaluateHideConditions,
  hideSlideConditions,
  stripHideSlideTagsInDeck,
  withSlideNumber,
} from '../hide-slide';

export type RenderDeckOptions = RenderTemplateOptions & {
  /**
   * 1-based numbers of the slides to keep, in this order. Default: every
   * slide of the deck, in its order.
   */
  slides?: number[];
};

const toBuffer = (pptx: Buffer | Uint8Array | ArrayBuffer): Buffer =>
  Buffer.isBuffer(pptx)
    ? pptx
    : pptx instanceof ArrayBuffer
      ? Buffer.from(pptx)
      : Buffer.from(pptx.buffer, pptx.byteOffset, pptx.byteLength);

type DeckSlide = {
  number: number;
  page?: number;
  flowCache?: FlowCache;
  /** What the slides before it show, for flowContinue(). */
  flowShown?: FlowShown;
};

const loadContent = async (
  pptx: string | Buffer | Uint8Array | ArrayBuffer,
): Promise<Buffer> =>
  typeof pptx === 'string' ? fs.promises.readFile(pptx) : toBuffer(pptx);

type SelectSlides = (available: number[]) => Promise<DeckSlide[]> | DeckSlide[];

/** Renders `slides` of a deck into a new deck (renderDeck's engine). */
async function buildDeck(
  content: Buffer,
  data: TemplateData,
  renderOptions: RenderTemplateOptions,
  select: SelectSlides,
): Promise<Buffer> {
  // Loaded here, through the package entry: upstream helpers import
  // src/index.ts, so loading automizer.ts or the template logic first (as
  // the CLI does) breaks on the import cycle
  const {
    default: Automizer,
    CallbackError,
    modify,
    TemplateRenderError,
  } = await import('../../index');
  const pres = new Automizer({
    removeExistingSlides: true,
    // the template's own slides leave the archive in removeUnreachableParts
    verbosity: 0,
  })
    .loadRoot(content)
    .load(content, 'deck');

  const available = await pres.getTemplate('deck').getAllSlideNumbers();
  const slides = await select(available);
  const missing = slides
    .map((s) => s.number)
    .filter((n) => !available.includes(n));
  if (missing.length) {
    throw new Error(
      `The deck has no slide ${missing.join(', ')} (slides: ${available.join(', ')})`,
    );
  }

  for (const { number, page, flowCache, flowShown } of slides) {
    pres.addSlide('deck', number, (slide) => {
      const render = modify.renderTemplate(data, {
        ...renderOptions,
        ...(flowCache ? { flowCache } : {}),
        ...(flowShown ? { flowShown } : {}),
        ...(page === undefined ? {} : { flowPage: page }),
      });
      slide.modify(async (document, context) => {
        try {
          await render(document, context);
        } catch (error) {
          const message =
            error instanceof Error ? error.message : String(error);
          throw new TemplateRenderError(`slide ${number}: ${message}`);
        }
      });
    });
  }

  try {
    const zip = await pres.getJSZip();
    await removeUnreachableParts(zip);
    // HIDE-SLIDE has been applied
    await stripHideSlideTagsInDeck(zip, renderOptions);
    return await generatePptx(zip);
  } catch (error) {
    // Automizer wraps callback errors; the tag error is what users need
    if (error instanceof CallbackError && error.cause instanceof Error) {
      throw error.cause;
    }
    throw error;
  }
}

/** Whether a slide of the deck calls flow() (cheap text check). */
const usesFlow = async (zip: JSZip, number: number): Promise<boolean> => {
  const xml = await readText(zip, PptPaths.slide(number));
  if (!xml) return false;
  // Runs joined: a tag can be split across several a:t
  const text = [...xml.matchAll(/<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g)]
    .map((m) => m[1])
    .join('');
  return /flow\s*\(/.test(text);
};

/**
 * How many slides the `flow()` lists of one template slide need with this
 * data: render the slide that many times, with `flowPage: 0`, `1`, ….
 * `renderDeck` does this by itself.
 *
 * @param pptx Path to a .pptx file or its contents.
 * @param slide 1-based number of the slide in the template.
 */
export const countFlowPages = async (
  pptx: string | Buffer | Uint8Array | ArrayBuffer,
  slide: number,
  data: TemplateData,
  options: RenderTemplateOptions = {},
): Promise<number> => {
  const content = await loadContent(pptx);
  return flowPagesOf(
    openCountDeck(content, await JSZip.loadAsync(content)),
    slide,
    data,
    options,
  );
};

/**
 * The deck opened once for countFlowPages: layouts, masters and themes are
 * parsed once and shared by the counts of all its slides.
 */
type CountDeck = { zip: JSZip; archive: ArchiveJszip };

const openCountDeck = (content: Buffer, zip: JSZip): CountDeck => {
  const archive = new ArchiveJszip(content, { mode: 'jszip' });
  // the zip is loaded already
  archive.archive = zip;
  return { zip, archive };
};

/**
 * Media files of IMAGE tags are not written while counting: the count
 * needs the text only. Every picture "exists" already.
 */
const countingTemplate = {
  mediaDeduplicator: {
    find: async (): Promise<string> => 'flow-count.png',
    add: (): void => undefined,
  },
} as unknown as RootPresTemplate;

/**
 * countFlowPages with the deck loaded already. Renders the slide by itself,
 * the way slide.modify() gets it in renderDeck (same slide XML, layout,
 * master and theme), without building a deck: the slide is rendered into
 * a copy that is thrown away.
 */
const flowPagesOf = async (
  deck: CountDeck,
  slide: number,
  data: TemplateData,
  options: RenderTemplateOptions,
): Promise<number> => {
  if (!(await usesFlow(deck.zip, slide))) return 1;
  // Loaded here, through the package entry (see buildDeck)
  const { modify, TemplateRenderError } = await import('../../index');
  const slidePath = PptPaths.slide(slide);
  const document = deck.archive.parseXml(await readText(deck.zip, slidePath));
  let pages = 1;
  const render = modify.renderTemplate(data, {
    ...options,
    // the overflow is what is being counted, not an error
    flowOverflow: 'warn',
    flowPage: Number.MAX_SAFE_INTEGER,
    onFlowPages: (count) => (pages = Math.max(pages, count)),
  });
  try {
    await render(document, {
      targetArchive: deck.archive,
      targetPath: slidePath,
      targetTemplate: countingTemplate,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new TemplateRenderError(`slide ${slide}: ${message}`);
  }
  return pages;
};

/**
 * Renders the tags of every slide of a template with one set of data and
 * returns the new .pptx: `renderTemplate` applied to each slide, the masters
 * and layouts of the template kept as they are.
 *
 * A failing tag rejects with a `TemplateRenderError` whose message starts
 * with the slide number ("slide 2: {{euro(p.preis)}}: Not a function …").
 *
 * A slide whose `flow()` lists do not fit is repeated as often as needed,
 * each copy with the next entries (`flowPage` 1, 2, … in the template).
 *
 * A slide with `{{HIDE-SLIDE condition}}` in its speaker notes is left out
 * when the condition is true; the tag itself is removed from the notes.
 *
 * `flowContinue(list)` in a template is `list` without the entries that
 * `flow()` shows on the slides before it, hidden copies not counted.
 *
 * @param pptx Path to a .pptx file or its contents.
 */
export const renderDeck = async (
  pptx: string | Buffer | Uint8Array | ArrayBuffer,
  data: TemplateData,
  options: RenderDeckOptions = {},
): Promise<Buffer> => {
  const { slides, ...renderOptions } = options;
  const content = await loadContent(pptx);
  const zip = await JSZip.loadAsync(content);
  const deck = openCountDeck(content, zip);
  return buildDeck(content, data, renderOptions, async (available) => {
    const result: DeckSlide[] = [];
    // The copies of a slide share the measurement of its lists
    const caches = new Map<number, FlowCache>();
    // flowContinue(): the entries of the flow() lists on the slides so far
    const shown = new FlowShown();
    for (const number of slides ?? available) {
      if (!caches.has(number)) caches.set(number, new FlowCache());
      const flowCache = caches.get(number);
      const flowShown = shown.clone();
      const slideOptions = { ...renderOptions, flowShown };
      const exists = available.includes(number);
      // {{HIDE-SLIDE …}} in the notes, evaluated per copy of the slide
      const hide = exists
        ? await hideSlideConditions(zip, number, renderOptions)
        : [];
      const hidden = async (page: number, pages: number) => {
        try {
          return await evaluateHideConditions(
            hide,
            data,
            slideOptions,
            renderOptions.flowPage ?? page,
            pages,
          );
        } catch (error) {
          throw withSlideNumber(error, number);
        }
      };
      // Hidden whatever flowPage is: no need to measure its lists
      const perPage = hide.some((c) => /\bflowPages?\b/.test(c.code));
      if (hide.length && !perPage && (await hidden(0, 1))) continue;

      // a slide whose flow() lists need more room repeats. Measured also
      // with a fixed flowPage: flowContinue() needs what the slide shows
      const plans: { plan: FlowPlan; list: unknown }[] = [];
      const counted = exists
        ? await flowPagesOf(deck, number, data, {
            ...slideOptions,
            flowCache,
            onFlowPlan: (plan, list) => plans.push({ plan, list }),
          })
        : 1;
      const pages = renderOptions.flowPage === undefined ? counted : 1;
      for (let page = 0; page < pages; page++) {
        if (perPage && (await hidden(page, counted))) continue;
        result.push({
          number,
          page: pages > 1 ? page : undefined,
          flowCache,
          flowShown,
        });
        const onSlide = renderOptions.flowPage ?? page;
        plans.forEach(({ plan, list }) =>
          shown.add(flowPlanEntries(plan, onSlide), list),
        );
      }
    }
    return result;
  });
};
