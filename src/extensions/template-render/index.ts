import fs from 'fs';
import JSZip from 'jszip';
import type {
  RenderTemplateOptions,
  TemplateData,
} from '../template-logic/types';
import { FlowCache } from '../text-flow';

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

type DeckSlide = { number: number; page?: number; flowCache?: FlowCache };

const loadContent = async (
  pptx: string | Buffer | Uint8Array | ArrayBuffer,
): Promise<Buffer> =>
  typeof pptx === 'string' ? fs.promises.readFile(pptx) : toBuffer(pptx);

/** Renders `slides` of a deck into a new deck (renderDeck's engine). */
const buildDeck = async (
  content: Buffer,
  data: TemplateData,
  renderOptions: RenderTemplateOptions,
  select: (available: number[]) => Promise<DeckSlide[]> | DeckSlide[],
): Promise<Buffer> => {
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
    // drops the template's own slides from the archive, not only the list
    cleanup: true,
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

  for (const { number, page, flowCache } of slides) {
    pres.addSlide('deck', number, (slide) => {
      const render = modify.renderTemplate(data, {
        ...renderOptions,
        ...(flowCache ? { flowCache } : {}),
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
    return await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
    });
  } catch (error) {
    // Automizer wraps callback errors; the tag error is what users need
    if (error instanceof CallbackError && error.cause instanceof Error) {
      throw error.cause;
    }
    throw error;
  }
};

/** Whether a slide of the deck calls flow() (cheap text check). */
const usesFlow = async (content: Buffer, number: number): Promise<boolean> => {
  const zip = await JSZip.loadAsync(content);
  const xml = await zip.file(`ppt/slides/slide${number}.xml`)?.async('string');
  return !!xml && /flow\s*\(/.test(xml.replace(/<[^>]+>/g, ''));
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
  if (!(await usesFlow(content, slide))) return 1;
  let pages = 1;
  await buildDeck(
    content,
    data,
    {
      ...options,
      // the overflow is what is being counted, not an error
      flowOverflow: 'warn',
      flowPage: Number.MAX_SAFE_INTEGER,
      onFlowPages: (count) => (pages = Math.max(pages, count)),
    },
    () => [{ number: slide }],
  );
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
 * @param pptx Path to a .pptx file or its contents.
 */
export const renderDeck = async (
  pptx: string | Buffer | Uint8Array | ArrayBuffer,
  data: TemplateData,
  options: RenderDeckOptions = {},
): Promise<Buffer> => {
  const { slides, ...renderOptions } = options;
  const content = await loadContent(pptx);
  return buildDeck(content, data, renderOptions, async (available) => {
    const result: DeckSlide[] = [];
    // The copies of a slide share the measurement of its lists
    const caches = new Map<number, FlowCache>();
    for (const number of slides ?? available) {
      if (!caches.has(number)) caches.set(number, new FlowCache());
      const flowCache = caches.get(number);
      // a slide whose flow() lists need more room repeats
      const pages =
        available.includes(number) && renderOptions.flowPage === undefined
          ? await countFlowPages(content, number, data, {
              ...renderOptions,
              flowCache,
            })
          : 1;
      for (let page = 0; page < pages; page++) {
        result.push({
          number,
          page: pages > 1 ? page : undefined,
          flowCache,
        });
      }
    }
    return result;
  });
};
