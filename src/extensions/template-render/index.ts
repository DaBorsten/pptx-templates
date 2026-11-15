import fs from 'fs';
import type {
  RenderTemplateOptions,
  TemplateData,
} from '../template-logic/types';

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

/**
 * Renders the tags of every slide of a template with one set of data and
 * returns the new .pptx: `renderTemplate` applied to each slide, the masters
 * and layouts of the template kept as they are.
 *
 * A failing tag rejects with a `TemplateRenderError` whose message starts
 * with the slide number ("slide 2: {{euro(p.preis)}}: Not a function …").
 *
 * @param pptx Path to a .pptx file or its contents.
 */
export const renderDeck = async (
  pptx: string | Buffer | Uint8Array | ArrayBuffer,
  data: TemplateData,
  options: RenderDeckOptions = {},
): Promise<Buffer> => {
  const { slides, ...renderOptions } = options;
  const content =
    typeof pptx === 'string'
      ? await fs.promises.readFile(pptx)
      : toBuffer(pptx);

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
  const numbers = slides ?? available;
  const missing = numbers.filter((n) => !available.includes(n));
  if (missing.length) {
    throw new Error(
      `The deck has no slide ${missing.join(', ')} (slides: ${available.join(', ')})`,
    );
  }

  for (const number of numbers) {
    pres.addSlide('deck', number, (slide) => {
      const render = modify.renderTemplate(data, renderOptions);
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
