/**
 * `{{HIDE-SLIDE condition}}` in the speaker notes of a slide: `renderDeck`
 * leaves the slide out when the condition is true. See docs/hide-slide.md.
 */
import fs from 'fs';
import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { XmlDocument, XmlElement } from '../../types/xml-types';
import { PptPaths } from '../../helper/ppt-paths';
import {
  fixSmartQuotes,
  isTableMarker,
  parseCommand,
  tagSyntax,
} from '../template-logic/commands';
import {
  TemplateExpression,
  TemplateScope,
} from '../template-logic/expression';
import type {
  RenderTemplateOptions,
  TemplateData,
} from '../template-logic/types';
import { FlowShown } from '../text-flow/flow';

/** One `{{HIDE-SLIDE …}}` of a slide's notes. */
export type HideSlideCondition = {
  /** The tag as written, e.g. `{{HIDE-SLIDE !projects.length}}`. */
  tag: string;
  /** The condition; empty: always hide. */
  code: string;
};

export type IsSlideHiddenOptions = RenderTemplateOptions & {
  /**
   * How many slides the `flow()` lists of the slide need, for `flowPages`
   * in the condition (`countFlowPages`). Default: 1
   */
  flowPages?: number;
};

const NOTES_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide';

const parseXml = (xml: string): XmlDocument =>
  new DOMParser().parseFromString(
    xml,
    'application/xml',
  ) as unknown as XmlDocument;

const elements = (
  parent: XmlDocument | XmlElement,
  tagName: string,
): XmlElement[] =>
  Array.from(
    parent.getElementsByTagName(tagName) as unknown as ArrayLike<XmlElement>,
  );

/** The text of a paragraph, runs joined (a tag can be split across runs). */
const paragraphText = (paragraph: XmlElement): string =>
  elements(paragraph, 'a:t')
    .map((t) => t.textContent)
    .join('');

type Syntax = ReturnType<typeof tagSyntax>;

/** The HIDE-SLIDE tags in `text`; other tags and invalid ones are skipped. */
const hideTagsIn = (
  text: string,
  syntax: Syntax,
  options: RenderTemplateOptions,
): HideSlideCondition[] =>
  Array.from(text.matchAll(syntax.pattern)).flatMap((m) => {
    if (isTableMarker(m[1])) return [];
    const inner = options.fixSmartQuotes ? fixSmartQuotes(m[1]) : m[1];
    try {
      const command = parseCommand(inner, m[0]);
      return command.type === 'HIDE-SLIDE'
        ? [{ tag: m[0], code: command.code }]
        : [];
    } catch {
      return [];
    }
  });

/** Path of the notes part of `ppt/slides/slide<n>.xml`, if it has notes. */
export const notesPartOf = async (
  zip: JSZip,
  slide: number,
): Promise<string | undefined> => {
  const rels = await zip.file(PptPaths.slideRels(slide))?.async('string');
  if (!rels) return undefined;
  const rel = elements(parseXml(rels), 'Relationship').find(
    (r) => r.getAttribute('Type') === NOTES_REL,
  );
  const target = rel?.getAttribute('Target');
  if (!target) return undefined;
  if (target.startsWith('/')) return target.slice(1);
  const n = /notesSlide(\d+)\.xml$/.exec(target)?.[1];
  return n ? PptPaths.notesSlide(Number(n)) : undefined;
};

/** The `{{HIDE-SLIDE …}}` tags in the speaker notes of a slide. */
export const hideSlideConditions = async (
  zip: JSZip,
  slide: number,
  options: RenderTemplateOptions = {},
): Promise<HideSlideCondition[]> => {
  const part = await notesPartOf(zip, slide);
  const xml = part && (await zip.file(part)?.async('string'));
  if (!xml) return [];
  const syntax = tagSyntax(options);
  return elements(parseXml(xml), 'a:p').flatMap((p) =>
    hideTagsIn(paragraphText(p), syntax, options),
  );
};

/** A read-only scope: built-ins, then additionalJsContext, then data. */
const scopeOf = (layers: Record<string, unknown>[]): TemplateScope => ({
  lookup(name: string) {
    for (const vars of layers) {
      if (vars && Object.prototype.hasOwnProperty.call(vars, name)) {
        return { found: true, value: vars[name] };
      }
    }
    return { found: false };
  },
});

/**
 * Whether one of the conditions is true. `page` is 0-based, as the
 * `flowPage` option; the condition sees it 1-based, as in the slide.
 */
export const evaluateHideConditions = async (
  conditions: HideSlideCondition[],
  data: TemplateData,
  options: RenderTemplateOptions,
  page: number,
  pages: number,
): Promise<boolean> => {
  if (!conditions.length) return false;
  const scope = scopeOf([
    data ?? {},
    options.additionalJsContext ?? {},
    {
      flowPage: page + 1,
      flowPages: pages,
      // The lists are not distributed here: flow() is the list itself
      flow: (list: unknown) => list,
      flowRest: (): unknown[] => [],
      flowContinue: (list: unknown) =>
        (options.flowShown ?? new FlowShown()).continue(list),
    },
  ]);
  const evaluation = {
    strict: options.strict ?? false,
    allowStringPatterns: options.allowStringPatterns ?? false,
  };
  for (const { tag, code } of conditions) {
    if (!code) return true;
    let value: unknown;
    try {
      value = await new TemplateExpression(code).evaluate(scope, evaluation);
    } catch (error) {
      if (!options.errorHandler) {
        const message = error instanceof Error ? error.message : String(error);
        // Loaded here: the renderer imports src/index.ts (import cycle)
        const { TemplateRenderError } =
          await import('../template-logic/renderer');
        throw new TemplateRenderError(`${tag}: ${message}`);
      }
      value = await options.errorHandler(error as Error, code);
    }
    if (value) return true;
  }
  return false;
};

/** A TemplateRenderError of a condition, its message led by the slide. */
export const withSlideNumber = (error: unknown, slide: number): unknown => {
  if (error instanceof Error && error.name === 'TemplateRenderError') {
    error.message = `slide ${slide}: ${error.message}`;
  }
  return error;
};

/**
 * Whether `renderDeck` leaves a slide out: one of the `{{HIDE-SLIDE …}}`
 * tags in its speaker notes is true with this data. For slides added with
 * `addSlide`: skip the `addSlide` call when it returns true.
 *
 * @param pptx Path to a .pptx file or its contents.
 * @param slide 1-based number of the slide in the template.
 */
export const isSlideHidden = async (
  pptx: string | Buffer | Uint8Array | ArrayBuffer,
  slide: number,
  data: TemplateData,
  options: IsSlideHiddenOptions = {},
): Promise<boolean> => {
  const content =
    typeof pptx === 'string' ? await fs.promises.readFile(pptx) : pptx;
  const zip = await JSZip.loadAsync(content);
  const conditions = await hideSlideConditions(zip, slide, options);
  try {
    return await evaluateHideConditions(
      conditions,
      data,
      options,
      options.flowPage ?? 0,
      options.flowPages ?? 1,
    );
  } catch (error) {
    throw withSlideNumber(error, slide);
  }
};

/**
 * Removes the `{{HIDE-SLIDE …}}` tags from a notes part; a paragraph that
 * held nothing else goes as well. Returns whether anything changed.
 */
export const stripHideSlideTags = (
  doc: XmlDocument,
  options: RenderTemplateOptions = {},
): boolean => {
  const syntax = tagSyntax(options);
  const strip = (text: string) =>
    text.replace(syntax.pattern, (raw) =>
      hideTagsIn(raw, syntax, options).length ? '' : raw,
    );
  let changed = false;

  elements(doc, 'a:p').forEach((paragraph) => {
    const before = paragraphText(paragraph);
    if (!hideTagsIn(before, syntax, options).length) return;
    changed = true;
    const texts = elements(paragraph, 'a:t');
    // Tags within one run: removed there, the formatting stays
    texts.forEach((t) => {
      t.textContent = strip(t.textContent);
    });
    // Tags split across runs: the paragraph's text goes into its first run
    const joined = paragraphText(paragraph);
    if (hideTagsIn(joined, syntax, options).length) {
      texts.forEach((t, i) => {
        t.textContent = i === 0 ? strip(joined) : '';
      });
    }
    const body = paragraph.parentNode as XmlElement;
    const isOnly = elements(body, 'a:p').length === 1;
    if (paragraphText(paragraph).trim() === '' && !isOnly) {
      body.removeChild(paragraph);
    }
  });
  return changed;
};

/** stripHideSlideTags for every notes part of a deck. */
export const stripHideSlideTagsInDeck = async (
  zip: JSZip,
  options: RenderTemplateOptions = {},
): Promise<void> => {
  const parts = Object.keys(zip.files).filter((name) =>
    /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name),
  );
  for (const part of parts) {
    const xml = await zip.file(part).async('string');
    // Cheap pre-check; the keyword is case-insensitive
    if (!/hide/i.test(xml)) continue;
    const doc = parseXml(xml);
    if (stripHideSlideTags(doc, options)) {
      zip.file(
        part,
        new XMLSerializer().serializeToString(
          doc as unknown as Parameters<XMLSerializer['serializeToString']>[0],
        ),
      );
    }
  }
};
